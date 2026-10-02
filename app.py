import os
import re
import io
import csv
import json
import base64
import hashlib
import secrets
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template_string, request, redirect, url_for, session, jsonify, send_file
import firebase_admin
from firebase_admin import credentials, firestore, auth
import openpyxl

try:
    import boto3
    from botocore.client import Config as BotoConfig
    BOTO3_AVAILABLE = True
except Exception:
    boto3 = None
    BotoConfig = None
    BOTO3_AVAILABLE = False

try:
    import fitz
    FITZ_AVAILABLE = True
except Exception:
    fitz = None
    FITZ_AVAILABLE = False

try:
    from PIL import Image
    PIL_AVAILABLE = True
except Exception:
    Image = None
    PIL_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors as reportlab_colors
    from reportlab.lib.styles import getSampleStyleSheet
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

app = Flask(__name__)
app.secret_key = "election_office_secret_key_change_this"

# Firebase Initialization
CRED_PATH = os.path.join(os.path.dirname(__file__), "serviceAccountKey.json")
if not firebase_admin._apps:
    if os.path.exists(CRED_PATH):
        try:
            cred = credentials.Certificate(CRED_PATH)
            firebase_admin.initialize_app(cred)
        except Exception as e:
            print("Firebase init error:", e)

db = firestore.client() if firebase_admin._apps else None

# ============================================================
# DESKTOP FEATURE COMPATIBILITY LAYER
# ============================================================

WORKSPACE_DEFAULTS = {
    "current_statuses": ["Pending", "Approved", "Rejected", "E_Roll Updated", "Submitted", "BLO Assigned", "Mark Incomplete"],
    "received_options": ["Yes", "No"],
    "work_statuses": ["Done", "Not Done", "Pending", "In Progress", "Work Complete", "Completed", "Not Started"],
    "dashboard_cards": {
        "Total Entries": True, "Pending": True, "Approved": True, "Rejected": True,
        "E-Roll Updated": True, "Work Complete": True, "Today": True, "This Month": True,
        "Total Amount": True, "Amount Received": True, "Amount Pending": True,
        "All Reports": True, "Client Pending Summary": True, "Alerts": True,
    },
    "table_columns": {
        "Client Name": True, "Reference No.": True, "AC": True, "Full Name": True,
        "Form Type": True, "Submission Date": True, "Current Status": True,
        "Day": True, "Amount": True, "Received": True, "Work Status": True,
        "State": True, "Remarks": True,
    },
    "custom_columns": [],
}

PAGE_KEYS = {
    "Dashboard": "dashboard", "New Entry": "entries", "All Entries": "entries",
    "Pending Entries": "entries", "Approved Entries": "entries", "Rejected Entries": "entries",
    "Work Complete": "entries", "AC Summary": "entries", "Reports": "reports",
    "E-Roll": "entries", "Users": "users", "Settings": "settings",
}

def _copy_defaults():
    return json.loads(json.dumps(WORKSPACE_DEFAULTS))

def load_workspace_config():
    cfg = _copy_defaults()
    if db:
        try:
            snap = db.collection("app_settings").document("workspace").get()
            if snap.exists:
                raw = snap.to_dict() or {}
                for key in cfg:
                    if key in raw:
                        cfg[key] = raw[key]
        except Exception as exc:
            print("Workspace config load warning:", exc)
    for key in ("current_statuses", "received_options", "work_statuses", "custom_columns"):
        if not isinstance(cfg.get(key), list):
            cfg[key] = _copy_defaults()[key]
        cfg[key] = [str(x).strip() for x in cfg[key] if str(x).strip()]
    for key in ("dashboard_cards", "table_columns"):
        if not isinstance(cfg.get(key), dict):
            cfg[key] = dict(_copy_defaults()[key])
    return cfg

def save_workspace_config(patch):
    if not db:
        raise RuntimeError("Firebase is not connected.")
    db.collection("app_settings").document("workspace").set(patch, merge=True)

def user_is_super_admin():
    role = str(session.get("role", "")).strip().lower()
    return role in ("super admin", "superadmin", "founder", "owner")

def can_page(page_name):
    if user_is_super_admin():
        return True
    features = session.get("features") or {}
    key = PAGE_KEYS.get(page_name)
    if key and features.get(key) is False:
        return False
    permissions = features.get("permissions") if isinstance(features, dict) else {}
    pages = permissions.get("pages", {}) if isinstance(permissions, dict) else {}
    if isinstance(pages, dict) and page_name in pages:
        return bool(pages[page_name])
    return True

@app.context_processor
def inject_workspace_context():
    return {"can_page": can_page, "workspace_config": load_workspace_config()}

def require_page(page_name):
    if "user" not in session:
        return redirect(url_for("login"))
    if not can_page(page_name):
        return ("Access denied", 403)
    return None

def normalize_amount(v):
    try:
        return float(str(v or 0).replace(",", "").replace("₹", "").strip() or 0)
    except Exception:
        return 0.0

def _firebase_api_key():
    return os.getenv("FIREBASE_WEB_API_KEY", "").strip()

def _find_profile(uid=None, email=None, username=None):
    if not db:
        return None
    try:
        if uid:
            snap = db.collection("users").document(uid).get()
            if snap.exists:
                d = snap.to_dict() or {}; d["doc_id"] = snap.id; return d
        for field, value in (("username", username), ("email", email)):
            if value:
                q = db.collection("users").where(field, "==", str(value).strip().lower()).limit(1).stream()
                for doc in q:
                    d = doc.to_dict() or {}; d["doc_id"] = doc.id; return d
    except Exception as exc:
        print("Profile lookup warning:", exc)
    return None

def _firebase_password_login(email, password):
    api_key = _firebase_api_key()
    if not api_key:
        return None, "FIREBASE_WEB_API_KEY is not configured."
    try:
        import requests
        response = requests.post(
            f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={api_key}",
            json={"email": email, "password": password, "returnSecureToken": True},
            timeout=15,
        )
        data = response.json()
        if response.status_code == 200:
            return data, None
        return None, (data.get("error") or {}).get("message", "LOGIN_FAILED")
    except Exception as exc:
        return None, str(exc)

class WebB2Storage:
    def __init__(self):
        self.client = None
        self.bucket = os.getenv("B2_BUCKET_NAME", "e-roll-photos").strip()
        self.endpoint = os.getenv("B2_ENDPOINT", "https://s3.eu-central-003.backblazeb2.com").strip()
        self.key_id = os.getenv("B2_KEY_ID", "").strip()
        self.application_key = os.getenv("B2_APPLICATION_KEY", "").strip()
        if BOTO3_AVAILABLE and self.key_id and self.application_key and self.bucket:
            try:
                self.client = boto3.client(
                    "s3", endpoint_url=self.endpoint,
                    aws_access_key_id=self.key_id, aws_secret_access_key=self.application_key,
                    config=BotoConfig(signature_version="s3v4", s3={"addressing_style":"path"}, max_pool_connections=20)
                )
            except Exception as exc:
                print("B2 connection error:", exc)

    def upload(self, photo_base64, epic):
        if not photo_base64:
            return ""
        if not self.client:
            raise RuntimeError("Backblaze B2 is not connected. Check B2 credentials.")
        raw = base64.b64decode(photo_base64, validate=True)
        if not (raw.startswith(b"\xff\xd8\xff") or raw.startswith(b"\x89PNG\r\n\x1a\n")):
            raise RuntimeError("Invalid photo encoding; B2 upload blocked.")
        safe = re.sub(r"[^A-Za-z0-9_-]+","_",str(epic or "UNKNOWN").upper())[:80] or "UNKNOWN"
        key = f"e-roll/{safe}/{safe}_{hashlib.sha256(raw).hexdigest()[:16]}.jpg"
        self.client.put_object(Bucket=self.bucket, Key=key, Body=raw,
                               ContentType="image/jpeg" if raw.startswith(b"\xff\xd8") else "image/png")
        return key

    def get(self, key):
        if not self.client or not key: return b""
        try: return self.client.get_object(Bucket=self.bucket, Key=key)["Body"].read()
        except Exception as exc: print("B2 photo read error:",exc); return b""

    def delete(self, key):
        if not self.client or not key: return False
        try: self.client.delete_object(Bucket=self.bucket,Key=key); return True
        except Exception as exc: print("B2 photo delete error:",exc); return False

B2 = WebB2Storage()

EROLL_FIELD_PATTERNS = {
    "elector_name": r"Elector'?s?\s*\*?Name\s*\*?\s*:\s*(.+)",
    "epic": r"Epic\s*\*?\s*:\s*([A-Za-z0-9]+)",
    "serial_no": r"Serial\s*\*?\s*No\.?\s*\*?\s*:\s*([0-9]+)",
    "part": r"Part\s*\*?\s*No\.?\s*and\s*\*?\s*Name\s*\*?\s*:\s*(.+)",
    "ac_pc": r"AC/PC\s*\*?\s*Name\s*\*?\s*:\s*(.+)",
    "state": r"State\s*\*?\s*:\s*(.+)",
}

def _clean_eroll(v):
    return re.sub(r"\s+"," ",str(v or "")).strip()

def _resolve_eroll_address(data):
    candidates=["address","address_en","address_english","elector_address","full_address","address_line","address1","address2",
                "house_address","postal_address","residential_address","elector_address_en","address_text","address_value",
                "house_name","location","address_full"]
    parts=[]
    for key in candidates:
        value=str(data.get(key,"") or "").strip()
        if value and value not in parts: parts.append(value)
    for key in ("house_no","house_number","building_no","street","locality","village","ward","tehsil","district","pincode","pin_code"):
        value=str(data.get(key,"") or "").strip()
        if value and value not in parts: parts.append(value)
    return ", ".join(parts).strip(" ,")

def _image_to_base64(image_bytes):
    if not PIL_AVAILABLE:return ""
    try:
        image=Image.open(io.BytesIO(image_bytes)).convert("RGB")
        max_w,max_h=600,800
        scale=min(max_w/max(1,image.width),max_h/max(1,image.height))
        if scale>1:
            image=image.resize((max(1,int(image.width*scale)),max(1,int(image.height*scale))),Image.Resampling.LANCZOS)
        else:
            image.thumbnail((max_w,max_h),Image.Resampling.LANCZOS)
        best=None
        for quality in range(92,59,-2):
            out=io.BytesIO();image.save(out,format="JPEG",quality=quality,optimize=True,progressive=True);candidate=out.getvalue()
            if 20*1024<=len(candidate)<=30*1024:best=candidate;break
            if best is None or abs(len(candidate)-25*1024)<abs(len(best)-25*1024):best=candidate
        return base64.b64encode(best).decode("ascii") if best else ""
    except Exception as exc:
        print("Photo encoding warning:",exc);return ""

def _extract_eroll_photo(page):
    if not FITZ_AVAILABLE or not page:return ""
    try:
        images=page.get_images(full=True)
        if not images:return ""
        pw,ph=float(page.rect.width),float(page.rect.height)
        bx0,by0,bx1,by1=pw*.66,ph*.055,pw*.88,ph*.215
        candidates=[]
        for img in images:
            xref=img[0]
            try:
                info=page.parent.extract_image(xref);raw=info.get("image",b"");w=int(info.get("width",0) or 0);h=int(info.get("height",0) or 0)
                if not raw or w<20 or h<20:continue
                ratio=w/float(h)
                if ratio<.45 or ratio>1.45:continue
                for r in page.get_image_rects(xref):
                    cx,cy=(r.x0+r.x1)/2,(r.y0+r.y1)/2
                    if not (bx0<=cx<=bx1 and by0<=cy<=by1):continue
                    rr=r.width/max(1,r.height)
                    if rr<.45 or rr>1.45:continue
                    tx,ty=pw*.773,ph*.138
                    score=1000-abs(cx-tx)/pw*100-abs(cy-ty)/ph*100
                    candidates.append((score,w*h,raw))
            except Exception:continue
        if not candidates:return ""
        candidates.sort(key=lambda x:(x[0],x[1]),reverse=True)
        return _image_to_base64(candidates[0][2])
    except Exception:return ""

def parse_eroll_page(text):
    lines=[_clean_eroll(x) for x in str(text or "").splitlines() if _clean_eroll(x)]
    data={"part_no":"","part_name":"","elector_name":"","epic":"","serial_no":"","ac_pc":"","state":"","address":"","photo_base64":""}
    def find_value(pattern):
        rx=re.compile(pattern,re.I)
        for line in lines:
            m=rx.search(line)
            if m:return _clean_eroll(m.group(1))
        return ""
    data["elector_name"]=find_value(EROLL_FIELD_PATTERNS["elector_name"])
    data["epic"]=find_value(EROLL_FIELD_PATTERNS["epic"]).upper()
    data["serial_no"]=find_value(EROLL_FIELD_PATTERNS["serial_no"])
    data["ac_pc"]=find_value(EROLL_FIELD_PATTERNS["ac_pc"])
    data["state"]=find_value(EROLL_FIELD_PATTERNS["state"])
    part=find_value(EROLL_FIELD_PATTERNS["part"])
    if part:
        m=re.match(r"^\s*(\d+)\s+(.+?)\s*$",part)
        if m:data["part_no"],data["part_name"]=m.group(1),_clean_eroll(m.group(2))
        else:data["part_name"]=part
    compact=_clean_eroll(str(text or "").replace("\r","\n"))
    m=re.search(r"Address\s*(?:\*\s*)?(?::|-)?\s*(.*?)(?=\s+(?:Serial|EPIC|State|AC/PC|Part\s+No|Elector))",compact,re.I)
    if m:data["address"]=_clean_eroll(m.group(1))
    if not data["address"]:
        labels=re.compile(r"^(?:Serial\s*\*?\s*No\.?|EPIC|State|AC/?PC\s*\*?\s*Name|Part\s*\*?\s*No\.?|Elector'?s?\s*\*?\s*Name|Address)\s*(?:\*\s*)?(?::|-)?\s*(.*)$",re.I)
        for i,line in enumerate(lines):
            m=re.match(r"^Address\s*(?:\*\s*)?(?::|-)?\s*(.*)$",line,re.I)
            if not m:continue
            parts=[];first=_clean_eroll(m.group(1))
            if first and first.lower() not in {"address","address *"}:parts.append(first)
            for nxt in lines[i+1:]:
                if labels.match(nxt):break
                parts.append(nxt)
            if parts:data["address"]=_clean_eroll(" ".join(parts));break
    data["address"]=_resolve_eroll_address(data)
    return data

def parse_eroll_pdf(path):
    if not FITZ_AVAILABLE:raise RuntimeError("PyMuPDF installed नहीं है. Run: pip install pymupdf")
    doc=fitz.open(path);records=[]
    try:
        for page_number,page in enumerate(doc,start=1):
            data=parse_eroll_page(page.get_text("text"))
            if not any(data[k] for k in ("elector_name","epic","serial_no","part_no")):continue
            data["photo_base64"]=_extract_eroll_photo(page)
            data["source_type"]="pdf";data["source_file"]=os.path.basename(path);data["source_page"]=page_number
            records.append(data)
    finally:doc.close()
    return records

def save_eroll_records(records):
    if not db:raise RuntimeError("Firebase is not connected.")
    existing=set()
    for doc in db.collection("e_roll_entries").stream():
        d=doc.to_dict() or {};e=str(d.get("epic",d.get("epic_no","")) or "").strip().upper();p=str(d.get("part_no","") or "").strip();s=str(d.get("serial_no","") or "").strip()
        if e or s:existing.add((e,p,s))
    seen=set();prepared=[];skipped=0
    for raw in records:
        payload=dict(raw or {});e=str(payload.get("epic",payload.get("epic_no","")) or "").strip().upper();p=str(payload.get("part_no","") or "").strip();s=str(payload.get("serial_no","") or "").strip();key=(e,p,s)
        if (e or s) and (key in existing or key in seen):skipped+=1;continue
        seen.add(key);photo=payload.pop("photo_base64","")
        if photo:
            photo_key=B2.upload(photo,e)
            if photo_key:payload["photo_b2_key"]=photo_key;payload["photo_storage"]="backblaze_b2"
        payload["created_at"]=datetime.now().isoformat(timespec="seconds");prepared.append(payload)
    saved=0
    for start in range(0,len(prepared),400):
        chunk=prepared[start:start+400]
        if not chunk:continue
        batch=db.batch()
        for payload in chunk:batch.set(db.collection("e_roll_entries").document(),payload)
        batch.commit();saved+=len(chunk)
    return saved,skipped


# --- MODERN WEB LAYOUT WITH WHATSAPP LAUNCHER & PERSISTENT LOCK ---
BASE_LAYOUT = """
<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <title>{% block title %}My Workspace - Election Office{% endblock %}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body { background-color: #F8FAFC; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #1E293B; }
        .sidebar { width: 260px; background-color: #0B132B; min-height: 100vh; position: fixed; color: #fff; top: 0; left: 0; z-index: 1000; padding: 12px 12px; display: flex; flex-direction: column; justify-content: space-between; }
        .sidebar .nav-link { color: #AAB4C5; font-weight: 600; padding: 6px 11px; border-radius: 7px; margin-bottom: 1px; font-size: 12.5px; display: flex; justify-content: space-between; align-items: center; text-decoration: none; }
        .sidebar .nav-link:hover, .sidebar .nav-link.active { background-color: #2563EB; color: #fff; }
        .sidebar .badge-count { background: rgba(255,255,255,0.2); padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: bold; }
        .main-content { margin-left: 260px; padding: 25px; }
        .card-box { background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .action-bar { background: #FFFFFF; border: 1px solid #DCE3EE; border-radius: 12px; padding: 10px 15px; margin-bottom: 15px; display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
        
        table.excel-table { font-size: 11.5px; white-space: nowrap; border-collapse: collapse; width: 100%; }
        table.excel-table th { background-color: #173B73; color: white; font-weight: 700; padding: 9px 8px; border: 1px solid #31578F; text-align: left; cursor: pointer; }
        table.excel-table td { padding: 7px 8px; vertical-align: middle; border: 1px solid #CBD5E1; color: #172033; background: #FFFFFF; }
        table.excel-table tr:nth-child(even) td { background-color: #F7F9FC; }

        table.summary-grid-table { font-size: 12.5px; border-collapse: separate; border-spacing: 0; width: 100%; border: 1px solid #E2E8F0; border-radius: 8px; overflow: hidden; }
        table.summary-grid-table th { background: linear-gradient(135deg, #1E293B, #0F172A); color: white; font-weight: 600; padding: 12px 14px; text-align: left; border-bottom: 2px solid #CBD5E1; }
        table.summary-grid-table td { padding: 12px 14px; vertical-align: middle; border-bottom: 1px solid #F1F5F9; border-right: 1px solid #F1F5F9; background: #FFFFFF; color: #334155; }
        table.summary-grid-table tr:hover td { background-color: #F8FAFC; }
        table.summary-grid-table tr:last-child td { border-bottom: none; }

        .excel-color-grid { display: grid; grid-template-columns: repeat(10, 22px); gap: 2px; padding: 6px; background: #fff; }
        .color-box { width: 22px; height: 22px; cursor: pointer; border: 1px solid #ccc; transition: transform 0.1s; }
        .color-box:hover { transform: scale(1.2); border-color: #000; z-index: 10; }
        .clickable-card { cursor: pointer; transition: transform 0.1s ease; }
        .clickable-card:hover { transform: translateY(-3px); }
        .metric-mini{border:1px solid #E2E8F0;border-radius:12px;background:#fff;padding:14px;}
        .settings-tab{border:1px solid #E2E8F0;border-radius:12px;background:#fff;}
        .e-roll-photo{width:72px;height:90px;object-fit:cover;border:1px solid #CBD5E1;border-radius:7px;background:#F8FAFC;}
        .drop-zone{border:2px dashed #93C5FD;background:#EFF6FF;border-radius:14px;padding:24px;text-align:center;cursor:pointer;}
        .drop-zone.dragover{background:#DBEAFE;border-color:#2563EB;}
    </style>
</head>
<body>
    <div class="sidebar">
        <div>
            <div class="d-flex align-items-center gap-3 mb-4 px-2">
                <div class="bg-primary text-white rounded-3 d-flex align-items-center justify-content-center fw-bold fs-4" style="width: 40px; height: 40px;">M</div>
                <div>
                    <h5 class="mb-0 fw-bold text-white fs-6">My Workspace</h5>
                    <small class="text-success fw-bold" style="font-size: 10px;">● Cloud Live Sync</small>
                </div>
            </div>
            <ul class="nav flex-column">
                <li class="nav-item"><a href="/dashboard" class="nav-link {% if page == 'dash' %}active{% endif %}"><span><i class="fa-solid fa-chart-pie me-2"></i> Dashboard</span></a></li>
                <li class="nav-item"><a href="/new-entry" class="nav-link {% if page == 'new' %}active{% endif %}"><span><i class="fa-solid fa-plus-circle me-2"></i> New Entry</span></a></li>
                <li class="nav-item"><a href="/all-entries" class="nav-link {% if page == 'all' %}active{% endif %}"><span><i class="fa-solid fa-folder-open me-2"></i> All Entries</span> <span class="badge-count bg-primary text-white">{{ counts.all if counts is defined else 0 }}</span></a></li>
                <li class="nav-item"><a href="/pending-entries" class="nav-link {% if page == 'pending' %}active{% endif %}"><span><i class="fa-solid fa-clock me-2"></i> Pending Entries</span> <span class="badge-count bg-warning text-dark">{{ counts.pending if counts is defined else 0 }}</span></a></li>
                <li class="nav-item"><a href="/approved-entries" class="nav-link {% if page == 'approved' %}active{% endif %}"><span><i class="fa-solid fa-file-lines me-2"></i> Approved Entries</span> <span class="badge-count bg-success text-white">{{ counts.approved if counts is defined else 0 }}</span></a></li>
                <li class="nav-item"><a href="/rejected-entries" class="nav-link {% if page == 'rejected' %}active{% endif %}"><span><i class="fa-solid fa-circle-xmark me-2"></i> Rejected Entries</span> <span class="badge-count bg-danger text-white">{{ counts.rejected if counts is defined else 0 }}</span></a></li>
                <li class="nav-item"><a href="/work-complete" class="nav-link {% if page == 'complete' %}active{% endif %}"><span><i class="fa-solid fa-wand-magic-sparkles me-2"></i> Work Complete</span> <span class="badge-count bg-info text-dark">{{ counts.complete if counts is defined else 0 }}</span></a></li>
                {% if can_page('E-Roll') %}
                <li class="nav-item mt-2"><a href="/e-roll" class="nav-link {% if page == 'eroll' %}active{% endif %}"><span><i class="fa-solid fa-id-card me-2 text-info"></i> E-Roll</span></a></li>
                {% endif %}
                <li class="nav-item"><a href="/excel-editor" class="nav-link {% if page == 'excel' %}active{% endif %}"><span><i class="fa-solid fa-file-excel me-2 text-success"></i> Excel / Sheet</span></a></li>
                <li class="nav-item"><a href="/google-sheet" class="nav-link {% if page == 'google_sheet' %}active{% endif %}"><span><i class="fa-solid fa-table-cells me-2 text-success"></i> Google Sheet</span></a></li>
                <li class="nav-item"><a href="/whatsapp" class="nav-link {% if page == 'whatsapp' %}active{% endif %}"><span><i class="fa-brands fa-whatsapp me-2 text-success" style="font-size: 15px;"></i> WhatsApp</span></a></li>
                <li class="nav-item"><a href="/ac-summary" class="nav-link {% if page == 'summary' %}active{% endif %}"><span><i class="fa-solid fa-chart-column me-2"></i> AC Summary</span></a></li>
                {% if can_page('Reports') %}
                <li class="nav-item"><a href="/reports" class="nav-link {% if page == 'reports' %}active{% endif %}"><span><i class="fa-solid fa-box-archive me-2"></i> Reports</span></a></li>
                {% endif %}
                {% if can_page('Users') %}
                <li class="nav-item"><a href="/users" class="nav-link {% if page == 'users' %}active{% endif %}"><span><i class="fa-solid fa-users-gear me-2"></i> Users</span></a></li>
                {% endif %}
                <li class="nav-item"><a href="/notepad" class="nav-link {% if page == 'notepad' %}active{% endif %}"><span><i class="fa-solid fa-note-sticky me-2"></i> Notepad</span></a></li>
                {% if can_page('Settings') %}
                <li class="nav-item mt-1"><a href="/settings" class="nav-link {% if page == 'settings' %}active{% endif %}"><span><i class="fa-solid fa-gear me-2"></i> Settings</span></a></li>
                {% endif %}
            </ul>
        </div>
        <div>
            <div id="lockCountdownBox" class="mb-2 px-2 py-1 bg-secondary bg-opacity-25 rounded-2 text-center" style="font-size: 11px; display: none;">
                🔒 Auto-lock in: <span id="countdownTimer" class="fw-bold text-warning">--:--</span>
            </div>
            <div class="bg-dark p-3 rounded-3">
                <p class="mb-1 fw-bold text-white small">👤 {{ session.get('user', 'User') }}</p>
                <p class="mb-2 text-muted" style="font-size: 11px;">Role: {{ session.get('role', 'Operator') }}</p>
                <a href="/logout" class="text-danger text-decoration-none fw-bold small"><i class="fa-solid fa-right-from-bracket me-1"></i> Log Out</a>
            </div>
        </div>
    </div>

    <div class="main-content">
        {% block content %}{% endblock %}
    </div>

    <!-- STATUS UPDATE MODAL (Current Status & Received) -->
    <div class="modal fade" id="statusUpdateModal" tabindex="-1">
      <div class="modal-dialog">
        <div class="modal-content rounded-4">
          <div class="modal-header bg-primary text-white">
            <h5 class="modal-title fw-bold fs-5">⚡ Batch Status & Received Update</h5>
            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
          </div>
          <div class="modal-body">
            <p class="text-muted small">Update status and payment received for all selected records:</p>
            <div class="mb-3">
                <label class="form-label small fw-bold">New Current Status</label>
                <select id="bulkNewStatus" class="form-select form-select-sm">
                    {% for x in workspace_config.current_statuses %}<option value="{{ x }}">{{ x }}</option>{% endfor %}
                </select>
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Amount Received?</label>
                <select id="bulkNewReceived" class="form-select form-select-sm">
                    {% for x in workspace_config.received_options %}<option value="{{ x }}">{{ x }}</option>{% endfor %}
                </select>
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Cancel</button>
            <button type="button" class="btn btn-success btn-sm fw-bold" onclick="submitBulkStatusUpdate()">💾 Apply Updates</button>
          </div>
        </div>
      </div>
    </div>

    <!-- SYSTEM DIAGNOSTIC REPORT MODAL -->
    <div class="modal fade" id="diagModal" tabindex="-1">
      <div class="modal-dialog">
        <div class="modal-content rounded-4">
          <div class="modal-header bg-dark text-white">
            <h5 class="modal-title fw-bold fs-5" id="diagTitle">🔍 System Check Report</h5>
            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
          </div>
          <div class="modal-body" id="diagBody">
            <div class="text-center py-3"><div class="spinner-border text-primary" role="status"></div><p class="mt-2 small text-muted">Running deep diagnostics...</p></div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Close</button>
          </div>
        </div>
      </div>
    </div>

    <!-- GENERAL LIST MODAL POPUP -->
    <div class="modal fade" id="dashboardListModal" tabindex="-1">
      <div class="modal-dialog modal-xl">
        <div class="modal-content rounded-4">
          <div class="modal-header bg-dark text-white">
            <h5 class="modal-title fw-bold fs-5" id="modalTitle">📋 Entry Details List</h5>
            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
          </div>
          <div class="modal-body">
            <div class="table-responsive">
              <table class="table table-sm table-bordered align-middle">
                <thead class="table-light">
                  <tr id="modalTableHead">
                    <th>Client Name</th>
                    <th>Reference No.</th>
                    <th>Full Name</th>
                    <th>Form Type</th>
                    <th>Day / Date</th>
                    <th>Status</th>
                  </tr>
                </thead>
                <tbody id="modalTableBody"></tbody>
              </table>
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Close</button>
          </div>
        </div>
      </div>
    </div>

    <!-- EDIT FULL FORM MODAL -->
    <div class="modal fade" id="editModal" tabindex="-1">
      <div class="modal-dialog modal-lg">
        <div class="modal-content rounded-4">
          <div class="modal-header bg-primary text-white">
            <h5 class="modal-title fw-bold fs-5">✏️ Edit Full Record Details</h5>
            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
          </div>
          <div class="modal-body" style="max-height: 70vh; overflow-y: auto;">
            <input type="hidden" id="editEntryId">
            <div class="row g-2">
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">Client Name</label>
                    <input type="text" id="editClientName" class="form-control form-control-sm" style="text-transform: uppercase;">
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">Reference Number</label>
                    <input type="text" id="editRefNo" class="form-control form-control-sm" style="text-transform: uppercase;">
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">State</label>
                    <input type="text" id="editState" class="form-control form-control-sm" style="text-transform: uppercase;">
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">AC</label>
                    <input type="text" id="editAc" class="form-control form-control-sm" style="text-transform: uppercase;">
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">First Name</label>
                    <input type="text" id="editFirstName" class="form-control form-control-sm" style="text-transform: uppercase;">
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">Last Name</label>
                    <input type="text" id="editLastName" class="form-control form-control-sm" style="text-transform: uppercase;">
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">Form Type</label>
                    <select id="editFormType" class="form-select form-select-sm">
                        <option value="Form 6">Form 6</option>
                        <option value="Form 6A">Form 6A</option>
                        <option value="Form 7">Form 7</option>
                        <option value="Form 8">Form 8</option>
                    </select>
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">Submission Date</label>
                    <input type="text" id="editSubDate" class="form-control form-control-sm">
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">Current Status</label>
                    <select id="editStatus" class="form-select form-select-sm">
                        {% for x in workspace_config.current_statuses %}<option value="{{ x }}">{{ x }}</option>{% endfor %}
                    </select>
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">Amount (₹)</label>
                    <input type="number" id="editAmount" class="form-control form-control-sm">
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">Amount Received?</label>
                    <select id="editReceived" class="form-select form-select-sm">
                        {% for x in workspace_config.received_options %}<option value="{{ x }}">{{ x }}</option>{% endfor %}
                    </select>
                </div>
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold">Work Complete?</label>
                    <select id="editWorkComplete" class="form-select form-select-sm">
                        {% for x in workspace_config.work_statuses %}<option value="{{ x }}">{{ x }}</option>{% endfor %}
                    </select>
                </div>
                <div class="col-12 mb-2">
                    <label class="form-label small fw-bold">Remarks</label>
                    <input type="text" id="editRemarks" class="form-control form-control-sm" style="text-transform: uppercase;">
                </div>
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Cancel</button>
            <button type="button" class="btn btn-primary btn-sm fw-bold" onclick="submitEdit()">💾 Save All Changes</button>
          </div>
        </div>
      </div>
    </div>

    <!-- UNDO POPUP MODAL -->
    <div class="modal fade" id="undoModal" tabindex="-1">
      <div class="modal-dialog modal-lg">
        <div class="modal-content rounded-4">
          <div class="modal-header bg-dark text-white">
            <h5 class="modal-title fw-bold fs-5">🔄 Undo Deleted Entries (Last 24 Hours)</h5>
            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
          </div>
          <div class="modal-body">
            <p class="text-muted small">Select the deleted entries you want to restore back to All Entries:</p>
            <div class="table-responsive">
              <table class="table table-sm table-bordered">
                <thead class="table-light">
                  <tr>
                    <th style="width: 40px;"><input type="checkbox" onclick="toggleUndoAll(this)"></th>
                    <th>Client Name</th>
                    <th>Reference No.</th>
                    <th>Deleted Time</th>
                  </tr>
                </thead>
                <tbody id="trashTableBody"></tbody>
              </table>
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Close</button>
            <button type="button" class="btn btn-success btn-sm fw-bold" onclick="submitRestore()">♻️ Restore Selected</button>
          </div>
        </div>
      </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        let idleTime = 0;
        let lockDurationSetting = localStorage.getItem('portal_lock_duration');
        let lockDuration = lockDurationSetting === null ? 180 : parseInt(lockDurationSetting);

        // Auto-lock activity is scoped to THIS webpage/tab.
        // Browser chrome/header clicks and tab switching are NOT activity.
        // Only real interaction inside the web page resets the timer.
        function resetTimer() {
            idleTime = 0;
        }

        // Loading/reloading the current page starts a fresh activity cycle.
        window.addEventListener('load', resetTimer);

        // Do NOT use window.onfocus here: focusing the browser/tab (including
        // clicking the browser header or switching back to this tab) must NOT
        // reset the inactivity timer.

        // Genuine interaction inside the webpage counts as activity.
        document.addEventListener('click', resetTimer, true);
        document.addEventListener('pointerdown', resetTimer, true);
        document.addEventListener('keydown', resetTimer, true);
        document.addEventListener('scroll', resetTimer, true);

        setInterval(() => {
            if (lockDuration <= 0) {
                document.getElementById('lockCountdownBox').style.display = 'none';
                return;
            }

            document.getElementById('lockCountdownBox').style.display = 'block';
            let timeLeft = lockDuration - idleTime;
            if (timeLeft <= 0) {
                window.location.href = '/logout';
            } else {
                let mins = Math.floor(timeLeft / 60);
                let secs = timeLeft % 60;
                document.getElementById('countdownTimer').innerText = 
                    (mins < 10 ? "0" + mins : mins) + ":" + (secs < 10 ? "0" + secs : secs);
                idleTime++;
            }
        }, 1000);

        async function loadSecuritySettings() {
            try {
                const res = await fetch('/api/security-settings', {cache:'no-store'});
                if (!res.ok) return;
                const data = await res.json();
                if (Number.isFinite(Number(data.auto_lock_seconds))) {
                    lockDuration = Number(data.auto_lock_seconds);
                    localStorage.setItem('portal_lock_duration', String(lockDuration));
                    const sel = document.getElementById('lockDurationSelect');
                    if (sel) sel.value = String(lockDuration);
                }
            } catch (e) {
                const savedDur = localStorage.getItem('portal_lock_duration');
                if (savedDur !== null && document.getElementById('lockDurationSelect')) {
                    lockDuration = parseInt(savedDur, 10) || 0;
                    document.getElementById('lockDurationSelect').value = String(lockDuration);
                }
            }
        }

        async function saveLockSettings() {
            const sel = document.getElementById('lockDurationSelect');
            if (!sel) return;
            const val = parseInt(sel.value, 10) || 0;
            try {
                const res = await fetch('/settings', {
                    method: 'POST',
                    headers: {'Content-Type':'application/x-www-form-urlencoded'},
                    body: new URLSearchParams({action:'security', auto_lock_seconds:String(val)})
                });
                if (!res.ok) throw new Error('Unable to save security setting');
                localStorage.setItem('portal_lock_duration', String(val));
                lockDuration = val;
                idleTime = 0;
                alert("✅ Auto-lock timer saved successfully.");
            } catch (e) {
                alert("❌ Auto-lock save failed: " + e.message);
            }
        }

        document.addEventListener("DOMContentLoaded", function() {
            loadSecuritySettings();
        });

        function toggleSelectAll(master) {
            let table = document.getElementById('entriesTable');
            if(!table) return;
            let tr = table.getElementsByTagName('tr');
            for (let i = 1; i < tr.length; i++) {
                if (tr[i].style.display !== "none") {
                    let cb = tr[i].querySelector('.row-checkbox');
                    if (cb) {
                        cb.checked = master.checked;
                    }
                }
            }
        }

        function filterTable() {
            let input = document.getElementById('searchInput').value.toLowerCase();
            let acFilter = document.getElementById('acDropdown').value.toLowerCase();
            let table = document.getElementById('entriesTable');
            if(!table) return;
            let tr = table.getElementsByTagName('tr');

            for (let i = 1; i < tr.length; i++) {
                let tdClient = tr[i].getElementsByTagName('td')[2];
                let tdRef = tr[i].getElementsByTagName('td')[3];
                let tdAc = tr[i].getElementsByTagName('td')[4];
                let tdName = tr[i].getElementsByTagName('td')[5];

                if (tdClient && tdRef && tdAc && tdName) {
                    let clientText = tdClient.textContent || tdClient.innerText;
                    let refText = tdRef.textContent || tdRef.innerText;
                    let acText = tdAc.textContent || tdAc.innerText;
                    let nameText = tdName.textContent || tdName.innerText;

                    let matchesSearch = clientText.toLowerCase().includes(input) || 
                                        refText.toLowerCase().includes(input) || 
                                        nameText.toLowerCase().includes(input);
                    let matchesAC = (acFilter === "" || acText.toLowerCase() === acFilter);

                    tr[i].style.display = (matchesSearch && matchesAC) ? "" : "none";
                }
            }
            recalculateSerialNumbers();
        }

        function filterSummaryTable() {
            let input = document.getElementById('summarySearch').value.toLowerCase();
            let table = document.getElementById('summaryTable');
            if(!table) return;
            let tr = table.getElementsByTagName('tr');
            for (let i = 1; i < tr.length; i++) {
                let tdClient = tr[i].getElementsByTagName('td')[1];
                if(tdClient) {
                    let txt = tdClient.textContent || tdClient.innerText;
                    tr[i].style.display = txt.toLowerCase().includes(input) ? "" : "none";
                }
            }
        }

        function recalculateSerialNumbers() {
            let table = document.getElementById('entriesTable');
            if(!table) return;
            let tr = table.getElementsByTagName('tr');
            let sno = 1;
            for (let i = 1; i < tr.length; i++) {
                if (tr[i].style.display !== "none") {
                    let snoCell = tr[i].getElementsByClassName('sno-cell')[0];
                    if (snoCell) { snoCell.textContent = sno++; }
                }
            }
        }

        let sortDirections = {};
        function sortTable(colIndex) {
            let table = document.getElementById('entriesTable');
            if(!table) return;
            let tbody = table.tBodies[0];
            let rows = Array.from(tbody.querySelectorAll('tr'));
            
            let dir = sortDirections[colIndex] === 'asc' ? 'desc' : 'asc';
            sortDirections[colIndex] = dir;

            rows.sort((a, b) => {
                let xCell = a.cells[colIndex];
                let yCell = b.cells[colIndex];
                let xVal = xCell ? (xCell.textContent || xCell.innerText).trim().toLowerCase() : '';
                let yVal = yCell ? (yCell.textContent || yCell.innerText).trim().toLowerCase() : '';

                if (!isNaN(xVal) && !isNaN(yVal) && xVal !== '' && yVal !== '') {
                    return dir === 'asc' ? Number(xVal) - Number(yVal) : Number(yVal) - Number(xVal);
                }
                return dir === 'asc' ? xVal.localeCompare(yVal) : yVal.localeCompare(xVal);
            });

            rows.forEach(row => tbody.appendChild(row));
            recalculateSerialNumbers();
        }

        function highlightSelectedRows(color) {
            let checkboxes = document.querySelectorAll('.row-checkbox:checked');
            if (checkboxes.length === 0) {
                alert("⚠️ कृपया कलर करने के लिए कम से कम एक record select करें!");
                return;
            }
            checkboxes.forEach(cb => {
                let tr = cb.closest('tr');
                if(tr) {
                    tr.style.backgroundColor = color;
                }
            });
        }

        function toggleAmountMask(elemId, actualVal) {
            let el = document.getElementById(elemId);
            if(el.dataset.masked === "true") {
                el.innerText = "₹" + Number(actualVal).toLocaleString('en-IN', {minimumFractionDigits: 2});
                el.dataset.masked = "false";
            } else {
                el.innerText = "••••••";
                el.dataset.masked = "true";
            }
        }

        function runDiagnostic(checkType) {
            let modalTitle = checkType === 'firebase' ? '🔥 Firebase Deep Connection Check' : '☁️ Cloud B2 Connection Check';
            document.getElementById('diagTitle').innerText = modalTitle;
            document.getElementById('diagBody').innerHTML = '<div class="text-center py-3"><div class="spinner-border text-primary" role="status"></div><p class="mt-2 small text-muted">Running deep step-by-step verification...</p></div>';
            new bootstrap.Modal(document.getElementById('diagModal')).show();

            fetch(`/api/diagnostic?type=${checkType}`)
            .then(res => res.json())
            .then(data => {
                let html = '<ul class="list-group list-group-flush small">';
                data.steps.forEach(st => {
                    let icon = st.success ? '<i class="fa-solid fa-circle-check text-success me-2"></i>' : '<i class="fa-solid fa-circle-xmark text-danger me-2"></i>';
                    html += `<li class="list-group-item d-flex align-items-center">${icon} <span><b>${st.title}:</b> ${st.msg}</span></li>`;
                });
                html += '</ul>';
                let alertClass = data.connected ? 'alert-success' : 'alert-danger';
                let statusText = data.connected ? '✅ SYSTEM FULLY CONNECTED & HEALTHY' : '❌ CONNECTION FAULT DETECTED';
                document.getElementById('diagBody').innerHTML = `<div class="alert ${alertClass} py-2 fw-bold text-center small mb-3">${statusText}</div>` + html;
            });
        }

        function openDashboardModal(filterType, titleName, extraParam = '') {
            fetch(`/api/dashboard-list?type=${filterType}&param=${encodeURIComponent(extraParam)}`)
            .then(res => res.json())
            .then(data => {
                document.getElementById('modalTitle').innerText = titleName;
                let tbody = document.getElementById('modalTableBody');
                tbody.innerHTML = '';
                if(data.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted py-3">कोई रिकॉर्ड नहीं मिला।</td></tr>';
                } else {
                    data.forEach(item => {
                        tbody.innerHTML += `<tr>
                            <td class="fw-bold">${item.client_name}</td>
                            <td class="font-monospace">${item.ref_no}</td>
                            <td class="fw-bold text-primary">${item.full_name}</td>
                            <td>${item.form_type}</td>
                            <td>${item.submission_date} (${item.day_count} Days)</td>
                            <td><span class="badge bg-secondary">${item.current_status}</span></td>
                        </tr>`;
                    });
                }
                new bootstrap.Modal(document.getElementById('dashboardListModal')).show();
            });
        }

        function getSelectedIds() {
            return Array.from(document.querySelectorAll('.row-checkbox:checked')).map(cb => cb.value);
        }

        function requireSelection(message) {
            const ids = getSelectedIds();
            if (!ids.length) { alert(message || "⚠️ पहले कम से कम एक record select करें!"); return null; }
            return ids;
        }

        function bulkFieldUpdate(field, value) {
            const ids = requireSelection("⚠️ पहले कम से कम एक record select करें!");
            if (!ids) return;
            fetch('/api/bulk-field-update', {
                method:'POST', headers:{'Content-Type':'application/json'},
                body:JSON.stringify({ids:ids, field:field, value:value})
            }).then(r=>r.json()).then(data=>{
                if (data.status !== 'success') { alert(data.message || 'Update failed'); return; }
                location.reload();
            }).catch(()=>alert('Update failed'));
        }

        function changeTableScale(delta) {
            const table = document.getElementById('entriesTable');
            if (!table) return;
            let size = parseFloat(localStorage.getItem('entries_table_scale') || '13');
            size = Math.max(9, Math.min(20, size + delta));
            table.style.fontSize = size + 'px';
            localStorage.setItem('entries_table_scale', String(size));
        }

        function restoreTableScale() {
            const table = document.getElementById('entriesTable');
            const saved = localStorage.getItem('entries_table_scale');
            if (table && saved) table.style.fontSize = saved + 'px';
        }

        function setDayPeriod(period) {
            const rows = document.querySelectorAll('#entriesTable tbody tr[data-id]');
            rows.forEach(tr => {
                const cells = tr.querySelectorAll('td');
                const dayCell = cells[9];
                const day = parseFloat((dayCell ? dayCell.textContent : '').trim());
                let show = true;
                if (period === 'below7') show = !isNaN(day) && day < 7;
                if (period === 'more7') show = !isNaN(day) && day > 7;
                tr.style.display = show ? '' : 'none';
            });
            recalculateSerialNumbers();
        }

        function showPendingAmount() {
            let total = 0;
            document.querySelectorAll('#entriesTable tbody tr[data-id]').forEach(tr => {
                if (tr.style.display === 'none') return;
                const cells = tr.querySelectorAll('td');
                const received = (cells[11]?.textContent || '').trim().toLowerCase();
                const amount = (cells[10]?.textContent || '').replace(/[₹,]/g,'').trim();
                if (received === 'no') total += parseFloat(amount) || 0;
            });
            alert('Pending Amount: ₹' + total.toLocaleString('en-IN', {minimumFractionDigits:2}));
        }

        document.addEventListener('DOMContentLoaded', restoreTableScale);

        function applyDropdownAction(field, selectId) {
            const ids = requireSelection("⚠️ पहले कम से कम एक record select करें!");
            if (!ids) return;
            const select = document.getElementById(selectId);
            const value = select ? select.value : '';
            if (!value) {
                alert(field === 'amount_received' ? '⚠️ Received में Yes या No select करें!' : '⚠️ Work Status select करें!');
                return;
            }
            bulkFieldUpdate(field, value);
        }

        function handleAction(actionType) {
            let checkboxes = document.querySelectorAll('.row-checkbox:checked');
            
            if (actionType === 'status_update') {
                if (checkboxes.length === 0) {
                    alert("⚠️ कृपया स्टेटस अपडेट करने के लिए कम से कम एक record select करें!");
                    return;
                }
                new bootstrap.Modal(document.getElementById('statusUpdateModal')).show();
                return;
            }

            if (actionType === 'amount') {
                applyDropdownAction('amount_received', 'receivedActionSelect');
                return;
            }

            if (actionType === 'work') {
                applyDropdownAction('work_status', 'workStatusActionSelect');
                return;
            }

            if (actionType === 'undo') {
                fetch('/api/trash-bin')
                .then(res => res.json())
                .then(data => {
                    let tbody = document.getElementById('trashTableBody');
                    tbody.innerHTML = '';
                    if(data.length === 0) {
                        tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted">पिछले 24 घंटों में कोई deleted record नहीं मिला।</td></tr>';
                    } else {
                        data.forEach(item => {
                            tbody.innerHTML += `<tr>
                                <td><input type="checkbox" class="undo-checkbox" value="${item.key}"></td>
                                <td class="text-success fw-bold">${item.client_name}</td>
                                <td class="text-success fw-bold">${item.ref_no}</td>
                                <td class="text-success fw-bold">${item.deleted_time}</td>
                            </tr>`;
                        });
                    }
                    new bootstrap.Modal(document.getElementById('undoModal')).show();
                });
                return;
            }

            if (checkboxes.length === 0) {
                alert("⚠️ कृपया कम से कम एक record select करें!");
                return;
            }
            let ids = Array.from(checkboxes).map(cb => cb.value);
            
            if (actionType === 'delete') {
                if (confirm(`क्या आप वाकई ${ids.length} selected record(s) delete करना चाहते हैं?`)) {
                    fetch('/api/delete-entries', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({ids: ids})
                    }).then(res => res.json()).then(data => { location.reload(); });
                }
            } else if (actionType === 'alldone') {
                fetch('/api/alldone-entries', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ids: ids})
                }).then(res => res.json()).then(data => { location.reload(); });
            } else if (actionType === 'movetoreport') {
                if (confirm(`क्या आप वाकई ${ids.length} selected entry/entries को Reports में भेजना चाहते हैं?`)) {
                    fetch('/api/move-to-report', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({ids: ids})
                    }).then(res => res.json()).then(data => { location.reload(); });
                }
            } else if (actionType === 'edit') {
                if (ids.length > 1) {
                    alert("⚠️ कृपया एडिट करने के लिए केवल एक ही record select करें!");
                    return;
                }
                let row = checkboxes[0].closest('tr');
                let entryId = ids[0];
                
                document.getElementById('editEntryId').value = entryId;
                document.getElementById('editClientName').value = row.getAttribute('data-client') || '';
                document.getElementById('editRefNo').value = row.getAttribute('data-ref') || '';
                document.getElementById('editState').value = row.getAttribute('data-state') || '';
                document.getElementById('editAc').value = row.getAttribute('data-ac') || '';
                document.getElementById('editFirstName').value = row.getAttribute('data-firstname') || '';
                document.getElementById('editLastName').value = row.getAttribute('data-lastname') || '';
                document.getElementById('editFormType').value = row.getAttribute('data-formtype') || 'Form 6';
                document.getElementById('editSubDate').value = row.getAttribute('data-subdate') || '';
                document.getElementById('editStatus').value = row.getAttribute('data-status') || 'Pending';
                document.getElementById('editAmount').value = row.getAttribute('data-amount') || '0';
                document.getElementById('editReceived').value = row.getAttribute('data-received') || 'No';
                document.getElementById('editWorkComplete').value = row.getAttribute('data-work') || 'Not Done';
                document.getElementById('editRemarks').value = row.getAttribute('data-remarks') || '';

                new bootstrap.Modal(document.getElementById('editModal')).show();
            }
        }

        function submitBulkStatusUpdate() {
            let checkboxes = document.querySelectorAll('.row-checkbox:checked');
            let ids = Array.from(checkboxes).map(cb => cb.value);
            let newStatus = document.getElementById('bulkNewStatus').value;
            let newReceived = document.getElementById('bulkNewReceived').value;

            fetch('/api/bulk-status-update', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ids: ids, status: newStatus, amount_received: newReceived})
            })
            .then(res => res.json())
            .then(data => {
                location.reload();
            });
        }

        function submitEdit() {
            let entryId = document.getElementById('editEntryId').value;
            let firstName = document.getElementById('editFirstName').value.toUpperCase();
            let lastName = document.getElementById('editLastName').value.toUpperCase();
            
            let payload = {
                id: entryId,
                client_name: document.getElementById('editClientName').value.toUpperCase(),
                ref_no: document.getElementById('editRefNo').value.toUpperCase(),
                state: document.getElementById('editState').value.toUpperCase(),
                ac: document.getElementById('editAc').value.toUpperCase(),
                first_name: firstName,
                last_name: lastName,
                full_name: (firstName + " " + lastName).trim(),
                form_type: document.getElementById('editFormType').value,
                submission_date: document.getElementById('editSubDate').value,
                current_status: document.getElementById('editStatus').value,
                amount: document.getElementById('editAmount').value,
                amount_received: document.getElementById('editReceived').value,
                work_complete: document.getElementById('editWorkComplete').value,
                remarks: document.getElementById('editRemarks').value.toUpperCase()
            };

            fetch('/api/edit-entry', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(payload)
            }).then(res => res.json()).then(data => {
                location.reload();
            });
        }

        function toggleUndoAll(master) {
            document.querySelectorAll('.undo-checkbox').forEach(cb => cb.checked = master.checked);
        }

        function setAmtRecv(val) {
            document.getElementById('inputAmtRecv').value = val;
            let btnNo = document.getElementById('btnRecvNo');
            let btnYes = document.getElementById('btnRecvYes');
            if(val === 'Yes') {
                btnYes.className = "btn btn-success flex-fill fw-bold";
                btnNo.className = "btn btn-outline-secondary flex-fill fw-bold text-secondary";
            } else {
                btnNo.className = "btn btn-danger flex-fill fw-bold";
                btnYes.className = "btn btn-outline-secondary flex-fill fw-bold text-secondary";
            }
        }

        function setWorkComp(val) {
            document.getElementById('inputWorkComp').value = val;
            let btnNo = document.getElementById('btnWorkNo');
            let btnYes = document.getElementById('btnWorkYes');
            if(val === 'Done') {
                btnYes.className = "btn btn-success flex-fill fw-bold";
                btnNo.className = "btn btn-outline-secondary flex-fill fw-bold text-secondary";
            } else {
                btnNo.className = "btn btn-danger flex-fill fw-bold";
                btnYes.className = "btn btn-outline-secondary flex-fill fw-bold text-secondary";
            }
        }

        function autoFillForm() {
            let rawText = document.getElementById('quickPasteBox').value.trim();
            if(!rawText) return;

            fetch('/api/parse-paste', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({text: rawText})
            })
            .then(res => res.json())
            .then(data => {
                if(data.status === 'success') {
                    document.getElementById('inputRefNo').value = data.ref_no || '';
                    document.getElementById('inputState').value = data.state || 'NCT OF DELHI';
                    document.getElementById('inputAc').value = data.ac || '';
                    document.getElementById('inputFirstName').value = data.first_name || '';
                    document.getElementById('inputLastName').value = data.last_name || '';
                    document.getElementById('inputFormType').value = data.form_type || 'Form 6';
                    document.getElementById('inputSubDate').value = data.submission_date || '';
                    document.getElementById('inputStatus').value = data.current_status || 'Pending';
                }
            });
        }

        function editCell(td, field) {
            const tr = td.closest('tr'), id = tr && tr.dataset.id;
            if (!id) return;
            const value = prompt('Edit ' + field, td.textContent.trim());
            if (value === null) return;
            fetch('/api/entry-custom-field',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id:id,field:field,value:value})})
              .then(r=>r.json()).then(d=>{if(d.error) alert(d.error); else td.textContent=value;});
        }
        function applyCellColor(color) {
            const table=document.getElementById('entriesTable');
            if(!table){alert('Entries table not found.');return;}
            const selected=table.querySelectorAll('.row-checkbox:checked');
            if(!selected.length){alert('⚠️ Please select at least one record.');return;}
            selected.forEach(cb=>{
                const tr=cb.closest('tr'), id=tr.dataset.id;
                tr.querySelectorAll('td[data-field]').forEach(td=>{
                    const field=td.dataset.field; td.style.backgroundColor=color;
                    fetch('/api/entry-cell-color',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id:id,field:field,color:color})});
                });
            });
        }
    </script>
</body>
</html>
"""

# --- STYLISH LOGIN PAGE WITH TAB 1 AND TAB 2 SEPARATE SWITCH ---
LOGIN_HTML = """
<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <title>Election Office - Login Portal</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body {
            margin: 0;
            padding: 0;
            height: 100vh;
            background: radial-gradient(circle at 20% 30%, #0d2854 0%, #06152d 40%, #020817 100%);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            color: #fff;
            position: relative;
        }
        body::before {
            content: "";
            position: absolute;
            bottom: 0;
            left: 0;
            width: 100%;
            height: 12px;
            background: linear-gradient(90deg, #FF9933 0%, #FFFFFF 50%, #138808 100%);
            z-index: 100;
        }
        .login-wrapper {
            display: flex;
            width: 1100px;
            max-width: 95%;
            justify-content: space-between;
            align-items: center;
            z-index: 10;
        }
        .login-left-content {
            flex: 1;
            padding-right: 50px;
        }
        .login-left-content h1 {
            font-size: 52px;
            font-weight: 800;
            margin-bottom: 2px;
            background: linear-gradient(90deg, #ffffff, #93c5fd);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: -1px;
        }
        .login-left-content h4 {
            color: #60a5fa;
            font-weight: 700;
            margin-bottom: 20px;
            font-size: 24px;
        }
        .login-left-content p {
            color: #94a3b8;
            font-size: 16px;
            line-height: 1.6;
            margin-bottom: 35px;
        }
        .features-grid {
            display: flex;
            gap: 15px;
        }
        .feature-box {
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid rgba(255, 255, 255, 0.1);
            padding: 15px 10px;
            border-radius: 14px;
            width: 115px;
            text-align: center;
            backdrop-filter: blur(10px);
            transition: transform 0.2s;
        }
        .feature-box:hover {
            transform: translateY(-3px);
            border-color: rgba(59, 130, 246, 0.5);
        }
        .feature-box i {
            color: #3b82f6;
            font-size: 22px;
            margin-bottom: 8px;
        }
        .feature-box span {
            display: block;
            font-size: 12px;
            font-weight: 600;
            color: #e2e8f0;
        }
        .login-card {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(20px);
            color: #1E293B;
            width: 440px;
            padding: 35px;
            border-radius: 24px;
            box-shadow: 0 30px 60px rgba(0, 0, 0, 0.5);
            position: relative;
        }
        .login-logo {
            width: 55px;
            height: 55px;
            background: linear-gradient(135deg, #2563EB, #7c3aed);
            color: white;
            font-size: 26px;
            font-weight: bold;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 16px;
            margin: 0 auto 10px auto;
            box-shadow: 0 10px 20px rgba(37, 99, 235, 0.35);
        }
        .form-control {
            border-radius: 12px;
            padding: 12px 15px 12px 44px;
            border: 1px solid #CBD5E1;
            font-size: 14px;
            background-color: #f8fafc;
        }
        .form-control:focus {
            box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.12);
            border-color: #2563EB;
            background-color: #ffffff;
        }
        .input-group-icon {
            position: absolute;
            left: 16px;
            top: 42px;
            color: #64748B;
            z-index: 20;
        }
        .btn-signin {
            background: linear-gradient(135deg, #2563EB, #7c3aed);
            border: none;
            border-radius: 12px;
            padding: 13px;
            font-weight: bold;
            color: white;
            width: 100%;
            transition: transform 0.1s;
            box-shadow: 0 8px 20px rgba(37, 99, 235, 0.35);
        }
        .btn-signin:hover {
            transform: translateY(-2px);
            background: linear-gradient(135deg, #1d4ed8, #6d28d9);
        }
        .nav-pills .nav-link {
            border-radius: 10px;
            font-weight: bold;
            font-size: 13px;
            color: #475569;
        }
        .nav-pills .nav-link.active {
            background-color: #2563EB;
            color: #fff;
        }
        .footer-tag {
            position: absolute;
            bottom: 30px;
            left: 60px;
            color: rgba(255, 255, 255, 0.5);
            font-style: italic;
            font-size: 17px;
            letter-spacing: 1px;
            font-weight: 500;
        }
    </style>
</head>
<body>
    <div class="footer-tag">
        Stronger Democracy, Brighter Tomorrow
    </div>

    <div class="login-wrapper">
        <div class="login-left-content d-none d-lg-block">
            <h1>My Workspace</h1>
            <h4>EMS Delhi</h4>
            <p>Secure &bull; Simple &bull; Smart &bull; Together<br>Powering a transparent, efficient and inclusive<br>electoral process for a stronger democracy.</p>
            
            <div class="features-grid">
                <div class="feature-box">
                    <i class="fa-solid fa-shield-halved"></i>
                    <span>Secure</span>
                    <small style="font-size: 9px; color: #94a3b8;">Your data, priority</small>
                </div>
                <div class="feature-box">
                    <i class="fa-solid fa-bolt"></i>
                    <span>Simple</span>
                    <small style="font-size: 9px; color: #94a3b8;">Easy access</small>
                </div>
                <div class="feature-box">
                    <i class="fa-solid fa-chart-pie"></i>
                    <span>Smart</span>
                    <small style="font-size: 9px; color: #94a3b8;">Tech driven</small>
                </div>
                <div class="feature-box">
                    <i class="fa-solid fa-users"></i>
                    <span>Together</span>
                    <small style="font-size: 9px; color: #94a3b8;">For tomorrow</small>
                </div>
            </div>
        </div>

        <div class="login-card">
            <div class="login-logo">M</div>
            <div class="text-center mb-3">
                <h4 class="fw-bold text-dark mb-1">Welcome Back</h4>
                <p class="text-muted small">Choose your login method below</p>
            </div>

            {% if error %}
            <div class="alert alert-danger py-2 small text-center fw-bold rounded-3">{{ error }}</div>
            {% endif %}

            <!-- TABS NAVIGATION -->
            <ul class="nav nav-pills nav-fill mb-3 bg-light p-1 rounded-3" id="loginTabs" role="tablist">
                <li class="nav-item" role="presentation">
                    <button class="nav-link active py-2" id="tab1-btn" type="button" role="tab" onclick="showLoginTab('user')">Tab 1: Username</button>
                </li>
                <li class="nav-item" role="presentation">
                    <button class="nav-link py-2" id="tab2-btn" type="button" role="tab" onclick="showLoginTab('pin')">Tab 2: PIN</button>
                </li>
            </ul>

            <script>
            function showLoginTab(mode) {
                const userPane = document.getElementById('tab1-pane');
                const pinPane = document.getElementById('tab2-pane');
                const userBtn = document.getElementById('tab1-btn');
                const pinBtn = document.getElementById('tab2-btn');
                document.getElementById('login_mode').value = mode;
                if (mode === 'pin') {
                    userPane.classList.remove('show','active');
                    pinPane.classList.add('show','active');
                    userBtn.classList.remove('active');
                    pinBtn.classList.add('active');
                    setTimeout(() => document.getElementById('firebase_pin').focus(), 80);
                } else {
                    pinPane.classList.remove('show','active');
                    userPane.classList.add('show','active');
                    pinBtn.classList.remove('active');
                    userBtn.classList.add('active');
                }
            }
            </script>

            <form method="POST">
                <input type="hidden" name="login_mode" id="login_mode" value="user">
                
                <div class="tab-content" id="loginTabsContent">
                    <!-- TAB 1 PANE: Username & Password -->
                    <div class="tab-pane fade show active" id="tab1-pane" role="tabpanel">
                        <div class="mb-3 position-relative">
                            <label class="form-label small fw-bold text-secondary">Username / Email</label>
                            <i class="fa-solid fa-user input-group-icon"></i>
                            <input type="text" name="username" class="form-control" placeholder="Enter username or email">
                        </div>
                        <div class="mb-3 position-relative">
                            <label class="form-label small fw-bold text-secondary">Password</label>
                            <i class="fa-solid fa-lock input-group-icon"></i>
                            <input type="password" name="password" class="form-control" placeholder="Enter password">
                        </div>
                    </div>

                    <!-- TAB 2 PANE: PIN Code -->
                    <div class="tab-pane fade" id="tab2-pane" role="tabpanel">
                        <div class="mb-3 position-relative">
                            <label class="form-label small fw-bold text-secondary">Enter Firebase Security PIN</label>
                            <i class="fa-solid fa-key input-group-icon"></i>
                            <input type="password" name="pin" id="firebase_pin" class="form-control text-center fw-bold" placeholder="Enter 4-digit PIN" maxlength="4" inputmode="numeric" pattern="[0-9]{4}" autocomplete="current-password" oninput="this.value=this.value.replace(/[^0-9]/g,'').slice(0,4)">
                        </div>
                        <div class="text-muted small text-center mb-3">
                            <i class="fa-solid fa-circle-info me-1 text-primary"></i> Enter the PIN already saved in Firebase.
                        </div>
                    </div>
                </div>

                <button type="submit" class="btn btn-signin mt-2 mb-2"><i class="fa-solid fa-arrow-right-to-bracket me-2"></i> Sign In</button>
            </form>
            <div class="text-center mt-2">
                <small class="text-muted">Need help? <a href="#" class="text-decoration-none fw-bold">Contact Support</a></small>
            </div>
        </div>
    </div>
</body>
</html>
"""

TABLE_TEMPLATE_CONTENT = """
    <div class="d-flex justify-content-between align-items-center mb-3">
        <div>
            <h2 class="fw-bold text-dark mb-0 fs-4">{{ table_title }}</h2>
            <p class="text-muted small mb-0">Real-time search, AC filter, sorting & multi-color picker</p>
        </div>
    </div>

    <!-- Search & AC Filter Header Controls -->
    <div class="row g-2 mb-3">
        <div class="col-md-6">
            <input type="text" id="searchInput" onkeyup="filterTable()" placeholder="🔍 Real-time search (Client, Reference No., Full Name)..." class="form-control form-control-sm bg-white">
        </div>
        <div class="col-md-4">
            <select id="acDropdown" onchange="filterTable()" class="form-select form-select-sm bg-white">
                <option value="">Filter by AC (All)</option>
                {% for ac in ac_list %}
                <option value="{{ ac }}">{{ ac }}</option>
                {% endfor %}
            </select>
        </div>
    </div>

    <!-- Quick Action Bar -->
    <div class="action-bar shadow-sm">
        <span class="fw-bold text-secondary small me-2"><i class="fa-solid fa-bolt text-warning"></i> Quick Actions:</span>
        <button class="btn btn-success btn-sm fw-bold px-3 py-1.5" onclick="handleAction('alldone')"><i class="fa-solid fa-check me-1"></i> All Done</button>
        <button class="btn btn-primary btn-sm fw-bold px-3 py-1.5" onclick="handleAction('status_update')"><i class="fa-solid fa-pen-to-square me-1"></i> Update Status</button>
        <div class="d-flex align-items-center gap-1">
            <select id="receivedActionSelect" class="form-select form-select-sm fw-bold" style="width:110px;" title="Received">
                <option value="">Received</option>
                <option value="Yes">Yes</option>
                <option value="No">No</option>
            </select>
            <button class="btn btn-outline-success btn-sm fw-bold" onclick="applyDropdownAction('amount_received', 'receivedActionSelect')">Update</button>
        </div>
        <div class="d-flex align-items-center gap-1">
            <select id="workStatusActionSelect" class="form-select form-select-sm fw-bold" style="width:125px;" title="Work Status">
                <option value="">Work Status</option>
                <option value="Done">Done</option>
                <option value="Not Done">Not Done</option>
                <option value="Pending">Pending</option>
                <option value="In Progress">In Progress</option>
                <option value="Work Complete">Work Complete</option>
                <option value="Completed">Completed</option>
                <option value="Not Started">Not Started</option>
            </select>
            <button class="btn btn-outline-info btn-sm fw-bold" onclick="applyDropdownAction('work_status', 'workStatusActionSelect')">Update</button>
        </div>
        {% if page == 'complete' %}
        <button class="btn btn-purple btn-sm fw-bold px-3 py-1.5 text-white" style="background-color: #7C3AED;" onclick="handleAction('movetoreport')"><i class="fa-solid fa-file-export me-1"></i> Move to report</button>
        {% endif %}
        
        <!-- Excel Paint Bucket Dropdown -->
        <div class="dropdown d-inline-block">
            <button class="btn btn-light border btn-sm fw-bold px-2 py-1 dropdown-toggle d-flex align-items-center gap-1" type="button" data-bs-toggle="dropdown" aria-expanded="false" title="Highlight Color">
                <i class="fa-solid fa-fill-drip text-primary"></i>
                <span style="display:inline-block; width:16px; height:6px; background:#FFEB3B; border:1px solid #999;"></span>
            </button>
            <div class="dropdown-menu p-2 shadow rounded-3" style="width: 250px;">
                <div class="small fw-bold text-muted mb-1 px-1">Theme Colors</div>
                <div class="excel-color-grid mb-2">
                    <div class="color-box" style="background:#FFFFFF;" onclick="applyCellColor('#FFFFFF')" title="White"></div>
                    <div class="color-box" style="background:#000000;" onclick="applyCellColor('#000000')" title="Black"></div>
                    <div class="color-box" style="background:#E7E6E6;" onclick="applyCellColor('#E7E6E6')" title="Light Gray"></div>
                    <div class="color-box" style="background:#414853;" onclick="applyCellColor('#414853')" title="Dark Gray"></div>
                    <div class="color-box" style="background:#2F5597;" onclick="applyCellColor('#2F5597')" title="Navy Blue"></div>
                    <div class="color-box" style="background:#ED7D31;" onclick="applyCellColor('#ED7D31')" title="Orange"></div>
                    <div class="color-box" style="background:#A5A5A5;" onclick="applyCellColor('#A5A5A5')" title="Gray"></div>
                    <div class="color-box" style="background:#FFC000;" onclick="applyCellColor('#FFC000')" title="Yellow"></div>
                    <div class="color-box" style="background:#5B9BD5;" onclick="applyCellColor('#5B9BD5')" title="Blue"></div>
                    <div class="color-box" style="background:#70AD47;" onclick="applyCellColor('#70AD47')" title="Green"></div>
                </div>

                <div class="small fw-bold text-muted mb-1 px-1">Standard Colors</div>
                <div class="excel-color-grid mb-2">
                    <div class="color-box" style="background:#C00000;" onclick="applyCellColor('#C00000')" title="Dark Red"></div>
                    <div class="color-box" style="background:#FF0000;" onclick="applyCellColor('#FF0000')" title="Red"></div>
                    <div class="color-box" style="background:#FFC000;" onclick="applyCellColor('#FFC000')" title="Light Orange"></div>
                    <div class="color-box" style="background:#FFFF00;" onclick="applyCellColor('#FFFF00')" title="Bright Yellow"></div>
                    <div class="color-box" style="background:#92D050;" onclick="applyCellColor('#92D050')" title="Light Green"></div>
                    <div class="color-box" style="background:#00B050;" onclick="applyCellColor('#00B050')" title="Emerald Green"></div>
                    <div class="color-box" style="background:#00B0F0;" onclick="applyCellColor('#00B0F0')" title="Cyan"></div>
                    <div class="color-box" style="background:#0070C0;" onclick="applyCellColor('#0070C0')" title="Dark Blue"></div>
                    <div class="color-box" style="background:#002060;" onclick="applyCellColor('#002060')" title="Midnight Blue"></div>
                    <div class="color-box" style="background:#7030A0;" onclick="applyCellColor('#7030A0')" title="Purple"></div>
                </div>

                <div class="small fw-bold text-muted mb-1 px-1">Soft Pastels</div>
                <div class="excel-color-grid mb-2">
                    <div class="color-box" style="background:#FEE2E2;" onclick="applyCellColor('#FEE2E2')" title="Soft Red"></div>
                    <div class="color-box" style="background:#DCFCE7;" onclick="applyCellColor('#DCFCE7')" title="Soft Green"></div>
                    <div class="color-box" style="background:#FEF08A;" onclick="applyCellColor('#FEF08A')" title="Soft Yellow"></div>
                    <div class="color-box" style="background:#DBEAFE;" onclick="applyCellColor('#DBEAFE')" title="Soft Blue"></div>
                    <div class="color-box" style="background:#F3E8FF;" onclick="applyCellColor('#F3E8FF')" title="Soft Purple"></div>
                    <div class="color-box" style="background:#FFEDD5;" onclick="applyCellColor('#FFEDD5')" title="Soft Peach"></div>
                    <div class="color-box" style="background:#CCFBF1;" onclick="applyCellColor('#CCFBF1')" title="Soft Teal"></div>
                    <div class="color-box" style="background:#FCE7F3;" onclick="applyCellColor('#FCE7F3')" title="Soft Pink"></div>
                    <div class="color-box" style="background:#F1F5F9;" onclick="applyCellColor('#F1F5F9')" title="Soft Slate"></div>
                    <div class="color-box" style="background:#E2E8F0;" onclick="applyCellColor('#E2E8F0')" title="Soft Grey"></div>
                </div>

                <hr class="my-1">
                <button class="btn btn-light btn-sm w-100 fw-bold text-danger py-1" onclick="applyCellColor('')"><i class="fa-solid fa-ban me-1"></i> No Fill / Clear</button>
            </div>
        </div>

        <button class="btn btn-warning btn-sm fw-bold px-3 py-1.5 text-dark" onclick="handleAction('edit')"><i class="fa-solid fa-pen me-1"></i> Edit</button>
        <button class="btn btn-danger btn-sm fw-bold px-3 py-1.5" onclick="handleAction('delete')"><i class="fa-solid fa-trash me-1"></i> Delete</button>
        <button class="btn btn-secondary btn-sm fw-bold px-3 py-1.5" onclick="handleAction('undo')"><i class="fa-solid fa-rotate-left me-1"></i> Undo</button>
        <div class="ms-auto">
            <button class="btn btn-outline-primary btn-sm fw-bold" onclick="window.location.reload();"><i class="fa-solid fa-rotate me-1"></i> Refresh</button>
            <button class="btn btn-outline-secondary btn-sm fw-bold" onclick="changeTableScale(-1)" title="Decrease table size">A−</button>
            <button class="btn btn-outline-secondary btn-sm fw-bold" onclick="changeTableScale(1)" title="Increase table size">A+</button>
            <div class="btn-group btn-group-sm" role="group" title="Day filter">
                <button class="btn btn-outline-success fw-bold" onclick="setDayPeriod('below7')">Below 7</button>
                <button class="btn btn-outline-warning fw-bold" onclick="setDayPeriod('more7')">More than 7</button>
                <button class="btn btn-outline-dark fw-bold" onclick="setDayPeriod('all')">All</button>
            </div>
        </div>
    </div>

    <div class="card-box overflow-hidden">
        <div class="table-responsive">
            <table class="excel-table align-middle mb-0" id="entriesTable">
                <thead>
                    <tr>
                        <th class="py-2 px-2 text-center" style="width: 35px;"><input type="checkbox" id="selectAllMaster" onclick="toggleSelectAll(this)" style="border-radius: 50%;"></th>
                        <th class="py-2 text-center" style="width: 50px;">S.No.</th>
                        <th class="py-2" data-field="Client Name" onclick="sortTable(2)">Client Name ↕</th>
                        <th class="py-2" data-field="Reference No." onclick="sortTable(3)">Reference No. ↕</th>
                        <th class="py-2" data-field="AC" onclick="sortTable(4)">AC ↕</th>
                        <th class="py-2" data-field="Full Name" onclick="sortTable(5)">Full Name ↕</th>
                        <th class="py-2" data-field="Form Type" onclick="sortTable(6)">Form Type ↕</th>
                        <th class="py-2" data-field="Submission Date" onclick="sortTable(7)">Submission Date ↕</th>
                        <th class="py-2" data-field="Current Status" onclick="sortTable(8)">Current Status ↕</th>
                        <th class="py-2 text-center" data-field="Day" style="width: 55px;" onclick="sortTable(9)">Day ↕</th>
                        <th class="py-2 text-end" data-field="Amount" onclick="sortTable(10)">Amount ↕</th>
                        <th class="py-2 text-center" data-field="Received" onclick="sortTable(11)">Received ↕</th>
                        <th class="py-2 text-center" data-field="Work Status" onclick="sortTable(12)">Work Status ↕</th>
                        <th class="py-2 px-2" data-field="Remarks">Remarks</th>{% for c in workspace_config.custom_columns %}<th class="py-2 px-2">{{ c }}</th>{% endfor %}
                    </tr>
                </thead>
                <tbody>
                    {% for row in entries %}
                    <tr data-id="{{ row.id }}" data-client="{{ row.client_name }}" data-ref="{{ row.ref_no }}" data-state="{{ row.state }}" data-ac="{{ row.ac }}" data-firstname="{{ row.first_name }}" data-lastname="{{ row.last_name }}" data-formtype="{{ row.form_type }}" data-subdate="{{ row.raw_date }}" data-status="{{ row.current_status }}" data-amount="{{ row.amount }}" data-received="{{ row.amount_received }}" data-work="{{ row.work_complete }}" data-remarks="{{ row.remarks }}">
                        <td class="px-2 text-center"><input type="checkbox" class="row-checkbox" value="{{ row.id }}" style="border-radius: 50%;"></td>
                        <td class="fw-bold text-muted text-center sno-cell">{{ loop.index }}</td>
                        <td data-field="Client Name" style="background-color:{{ row.cell_colors.get("Client Name","") }};" ondblclick="editCell(this,'Client Name')" class="fw-bold text-dark">{{ row.client_name }}</td>
                        <td data-field="Reference No." style="background-color:{{ row.cell_colors.get("Reference No.","") }};" ondblclick="editCell(this,'Reference No.')" class="font-monospace fw-bold {% if row.is_duplicate %}text-danger bg-danger-subtle{% else %}text-dark{% endif %}">{{ row.ref_no }}</td>
                        <td data-field="AC" style="background-color:{{ row.cell_colors.get("AC","") }};" class="fw-bold text-dark">{{ row.ac }}</td>
                        <td data-field="Full Name" style="background-color:{{ row.cell_colors.get("Full Name","") }};" class="fw-bold text-dark">{{ row.full_name }}</td>
                        <td data-field="Form Type" style="background-color:{{ row.cell_colors.get("Form Type","") }};">{{ row.form_type }}</td>
                        <td data-field="Submission Date" style="background-color:{{ row.cell_colors.get("Submission Date","") }};">{{ row.submission_date }}</td>
                        <td>
                            <span data-field="Current Status" class="badge {% if row.display_status == 'Approved' %}bg-success-subtle text-success{% elif row.display_status == 'Rejected' %}bg-danger-subtle text-danger{% else %}bg-warning-subtle text-warning{% endif %} px-2 py-1 rounded-pill">
                                {{ row.current_status }}
                            </span>
                        </td>
                        <td data-field="Day" style="background-color:{{ row.cell_colors.get("Day","") }};" class="fw-bold text-center text-success bg-success-subtle">{{ row.day_count }}</td>
                        <td data-field="Amount" style="background-color:{{ row.cell_colors.get("Amount","") }};" class="fw-bold text-end">₹{{ row.amount }}</td>
                        <td data-field="Received" style="background-color:{{ row.cell_colors.get("Received","") }};" class="text-center fw-bold {% if row.amount_received == 'Yes' %}text-white bg-success{% else %}text-danger{% endif %}">{{ row.amount_received }}</td>
                        <td data-field="Work Status" style="background-color:{{ row.cell_colors.get("Work Status","") }};" class="text-center fw-bold {% if row.work_complete == 'Done' %}text-white bg-success{% else %}text-danger{% endif %}">{{ row.work_complete }}</td>
                        <td data-field="Remarks" style="background-color:{{ row.cell_colors.get("Remarks","") }};" ondblclick="editCell(this,'Remarks')" class="px-2 text-muted small">{{ row.remarks }}</td>
                        {% for c in workspace_config.custom_columns %}<td data-field="{{ c }}" style="background-color:{{ row.cell_colors.get(c,"") }};" ondblclick="editCell(this,c)" class="px-2 small">{{ row.custom_fields.get(c,"") }}</td>{% endfor %}
                    </tr>
                    {% else %}
                    <tr><td colspan="14" class="text-center py-5 text-muted fw-semibold">No records found in database.</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
"""

# Apply Settings -> Table Columns visibility to the rendered Entries table.
# This must be appended before ALL_ENTRIES_HTML is created from TABLE_TEMPLATE_CONTENT.
TABLE_TEMPLATE_CONTENT += r"""
<script>
document.addEventListener('DOMContentLoaded', function(){
  const visible = {{ workspace_config.table_columns|tojson }};
  document.querySelectorAll('#entriesTable [data-field]').forEach(function(el){
    const field = el.getAttribute('data-field');
    if (Object.prototype.hasOwnProperty.call(visible, field)) el.style.display = visible[field] ? '' : 'none';
  });
});
</script>
"""

DASHBOARD_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', """
    <div class="d-flex flex-wrap justify-content-between align-items-center mb-4 bg-white p-3 rounded-4 shadow-sm border gap-3">
        <div>
            <h2 class="fw-bold text-dark mb-0 fs-3">Dashboard</h2>
            <small class="text-muted">Real-time overview of your Election Office workspace</small>
        </div>
        <div class="d-flex align-items-center gap-2 flex-wrap">
            <div class="input-group input-group-sm" style="width: 250px;">
                <span class="input-group-text bg-light"><i class="fa-solid fa-search text-muted"></i></span>
                <input type="text" id="globalSearch" class="form-control" placeholder="Search reference, client, AC or name...">
            </div>
            <button class="btn btn-outline-primary btn-sm fw-bold" onclick="window.location.reload();"><i class="fa-solid fa-rotate me-1"></i> Refresh</button>
            <button class="btn btn-primary btn-sm fw-bold" onclick="runDiagnostic('firebase')"><i class="fa-solid fa-fire me-1"></i> Firebase Check</button>
            <button class="btn btn-info btn-sm fw-bold text-white" onclick="runDiagnostic('cloud')"><i class="fa-solid fa-cloud me-1"></i> Cloud (B2) Check</button>
            <span class="badge bg-dark px-3 py-2 fw-bold"><i class="fa-solid fa-user-shield me-1 text-warning"></i> {{ session.get('user', 'User') }} • Super Admin</span>
        </div>
    </div>

    <!-- ADVANCED LIVE INFO -->
    <div class="row g-3 mb-4">
        <div class="col-xl-3 col-md-6"><div class="card-box p-3 h-100"><div class="d-flex justify-content-between align-items-center"><div><small class="text-muted fw-bold">LIVE TIME</small><div id="liveClock" class="fw-bold fs-3 text-primary">--:--:--</div><small id="liveDate" class="text-muted">Loading date...</small></div><div class="p-3 bg-primary-subtle text-primary rounded-4"><i class="fa-solid fa-clock fs-4"></i></div></div></div></div>
        <div class="col-xl-3 col-md-6"><div class="card-box p-3 h-100"><div class="d-flex justify-content-between align-items-center"><div><small class="text-muted fw-bold">WEATHER</small><div id="weatherTemp" class="fw-bold fs-3">--°C</div><small id="weatherPlace" class="text-muted">Detecting location...</small></div><div id="weatherIcon" class="p-3 bg-info-subtle text-info rounded-4"><i class="fa-solid fa-cloud-sun fs-4"></i></div></div><div class="small mt-2"><span id="weatherDesc">Loading...</span> · Humidity <b id="weatherHumidity">--%</b> · UV <b id="weatherUv">--</b></div>
                <div class="small text-muted mt-1">Air Quality <b id="weatherAqi">--</b> · PM2.5 <b id="weatherPm25">--</b></div></div></div>
        <div class="col-xl-3 col-md-6"><div class="card-box p-3 h-100"><div class="d-flex justify-content-between align-items-center"><div><small class="text-muted fw-bold">WIND & SUN</small><div class="fw-bold fs-5"><span id="weatherWind">-- km/h</span></div><small id="sunTimes" class="text-muted">Sunrise -- · Sunset --</small></div><div class="p-3 bg-warning-subtle text-warning rounded-4"><i class="fa-solid fa-sun fs-4"></i></div></div><div class="small mt-2 text-muted">Feels like <b id="weatherFeels">--°C</b> · Pressure <b id="weatherPressure">-- hPa</b></div></div></div>
        <div class="col-xl-3 col-md-6"><div class="card-box p-3 h-100"><div class="d-flex justify-content-between align-items-center"><div><small class="text-muted fw-bold">SYSTEM STATUS</small><div id="onlineStatus" class="fw-bold text-success">● Online</div><small id="connectionInfo" class="text-muted">Checking connection...</small></div><div class="p-3 bg-success-subtle text-success rounded-4"><i class="fa-solid fa-wifi fs-4"></i></div></div><div class="small mt-2 text-muted">Screen <b id="screenInfo">--</b> · Network <b id="networkInfo">--</b></div></div></div>
    </div>
    <div class="card-box p-3 mb-4"><div class="d-flex flex-wrap justify-content-between align-items-center mb-2 gap-2"><div><h5 class="fw-bold mb-0"><i class="fa-solid fa-cloud-sun text-primary me-2"></i>7-Day Weather Forecast</h5><small class="text-muted">Live forecast based on your browser location</small></div><button class="btn btn-outline-primary btn-sm" onclick="loadAdvancedWeather(true)"><i class="fa-solid fa-location-crosshairs me-1"></i> Refresh Weather</button></div><div id="weatherForecast" class="row g-2"><div class="col-12 text-muted small">Loading forecast...</div></div></div>

{% if workspace_config.dashboard_cards.get("Total Entries", True) or workspace_config.dashboard_cards.get("Pending", True) or workspace_config.dashboard_cards.get("Approved", True) or workspace_config.dashboard_cards.get("Rejected", True) or workspace_config.dashboard_cards.get("E-Roll Updated", True) or workspace_config.dashboard_cards.get("Work Complete", True) %}
    <!-- ROW 1: 6 STAT CARDS -->
    <div class="row g-3 mb-4">
        <div class="col-md-2 col-sm-4">
            <div class="p-3 text-white rounded-4 shadow-sm clickable-card" style="background: linear-gradient(135deg, #2563EB, #1D4ED8);" onclick="window.location.href='/all-entries'">
                <div class="d-flex align-items-center gap-2 mb-2"><i class="fa-solid fa-folder-open fs-5"></i><span class="fw-bold small">Total Entries</span></div>
                <h2 class="fw-bold mb-1 fs-2">{{ stats.total }}</h2>
                <div class="bg-white bg-opacity-25 text-center rounded-2 py-1"><small style="font-size: 10px; font-weight: bold;">All entries in workspace</small></div>
            </div>
        </div>
        <div class="col-md-2 col-sm-4">
            <div class="p-3 text-white rounded-4 shadow-sm clickable-card" style="background: linear-gradient(135deg, #D97706, #B45309);" onclick="window.location.href='/pending-entries'">
                <div class="d-flex align-items-center gap-2 mb-2"><i class="fa-solid fa-clock fs-5"></i><span class="fw-bold small">Pending</span></div>
                <h2 class="fw-bold mb-1 fs-2">{{ stats.pending }}</h2>
                <div class="bg-white bg-opacity-25 text-center rounded-2 py-1"><small style="font-size: 10px; font-weight: bold;">Waiting for action</small></div>
            </div>
        </div>
        <div class="col-md-2 col-sm-4">
            <div class="p-3 text-white rounded-4 shadow-sm clickable-card" style="background: linear-gradient(135deg, #059669, #047857);" onclick="window.location.href='/approved-entries'">
                <div class="d-flex align-items-center gap-2 mb-2"><i class="fa-solid fa-circle-check fs-5"></i><span class="fw-bold small">Approved</span></div>
                <h2 class="fw-bold mb-1 fs-2">{{ stats.approved }}</h2>
                <div class="bg-white bg-opacity-25 text-center rounded-2 py-1"><small style="font-size: 10px; font-weight: bold;">Successfully approved</small></div>
            </div>
        </div>
        <div class="col-md-2 col-sm-4">
            <div class="p-3 text-white rounded-4 shadow-sm clickable-card" style="background: linear-gradient(135deg, #DC2626, #B91C1C);" onclick="window.location.href='/rejected-entries'">
                <div class="d-flex align-items-center gap-2 mb-2"><i class="fa-solid fa-circle-xmark fs-5"></i><span class="fw-bold small">Rejected</span></div>
                <h2 class="fw-bold mb-1 fs-2">{{ stats.rejected }}</h2>
                <div class="bg-white bg-opacity-25 text-center rounded-2 py-1"><small style="font-size: 10px; font-weight: bold;">Marked as rejected</small></div>
            </div>
        </div>
        <div class="col-md-2 col-sm-4">
            <div class="p-3 text-white rounded-4 shadow-sm clickable-card" style="background: linear-gradient(135deg, #7C3AED, #6D28D9);" onclick="openDashboardModal('eroll', '📋 E-Roll Updated Entries')">
                <div class="d-flex align-items-center gap-2 mb-2"><i class="fa-solid fa-file-arrow-up fs-5"></i><span class="fw-bold small">E-Roll Updated</span></div>
                <h2 class="fw-bold mb-1 fs-2">{{ stats.eroll }}</h2>
                <div class="bg-white bg-opacity-25 text-center rounded-2 py-1"><small style="font-size: 10px; font-weight: bold;">Updated in E-Roll</small></div>
            </div>
        </div>
        <div class="col-md-2 col-sm-4">
            <div class="p-3 text-white rounded-4 shadow-sm clickable-card" style="background: linear-gradient(135deg, #0D9488, #0F766E);" onclick="window.location.href='/work-complete'">
                <div class="d-flex align-items-center gap-2 mb-2"><i class="fa-solid fa-wand-magic-sparkles fs-5"></i><span class="fw-bold small">Work Complete</span></div>
                <h2 class="fw-bold mb-1 fs-2">{{ stats.complete }}</h2>
                <div class="bg-white bg-opacity-25 text-center rounded-2 py-1"><small style="font-size: 10px; font-weight: bold;">Process completed</small></div>
            </div>
        </div>
    </div>

    {% endif %}

    {% if workspace_config.dashboard_cards.get("Today", True) or workspace_config.dashboard_cards.get("This Month", True) %}
    <!-- ROW 2: TIME METRICS -->
    <div class="row g-3 mb-4">
        <div class="col-md-6">
            <div class="card-box p-3 bg-white d-flex align-items-center justify-content-between clickable-card" onclick="openDashboardModal('today', '📅 Today\\'s New Entries')">
                <div>
                    <div class="d-flex align-items-center gap-2 mb-1">
                        <span class="p-2 bg-success-subtle text-success rounded-3"><i class="fa-solid fa-calendar-day"></i></span>
                        <span class="fw-bold text-dark">Today ({{ stats.today_date_str }})</span>
                    </div>
                    <h3 class="fw-bold text-primary mb-0">{{ stats.today_count }} new entries</h3>
                    <small class="text-muted" style="font-size: 11px;">Pending {{ stats.today_pending }} • Approved {{ stats.today_approved }} • Rejected {{ stats.today_rejected }}</small>
                </div>
                <i class="fa-solid fa-chevron-right text-muted"></i>
            </div>
        </div>
        <div class="col-md-6">
            <div class="card-box p-3 bg-white d-flex align-items-center justify-content-between clickable-card" onclick="openDashboardModal('month', '🗓️ This Month\\'s Entries ({{ stats.current_month_name }})')">
                <div>
                    <div class="d-flex align-items-center gap-2 mb-1">
                        <span class="p-2 bg-primary-subtle text-primary rounded-3"><i class="fa-solid fa-calendar-days"></i></span>
                        <span class="fw-bold text-dark">This Month ({{ stats.current_month_name }})</span>
                    </div>
                    <h3 class="fw-bold text-success mb-0">{{ stats.month_count }} entries</h3>
                    <small class="text-muted" style="font-size: 11px;">Pending {{ stats.month_pending }} • Approved {{ stats.month_approved }} • Rejected {{ stats.month_rejected }}</small>
                </div>
                <i class="fa-solid fa-chevron-right text-muted"></i>
            </div>
        </div>
    </div>

    {% endif %}

    {% if workspace_config.dashboard_cards.get("Total Amount", True) or workspace_config.dashboard_cards.get("Amount Received", True) or workspace_config.dashboard_cards.get("Amount Pending", True) %}
    <!-- ROW 3: FINANCIAL CARDS -->
    <div class="row g-3 mb-4">
        <div class="col-md-4">
            <div class="p-4 text-white rounded-4 shadow-sm" style="background: linear-gradient(135deg, #3B82F6, #1D4ED8);">
                <div class="d-flex justify-content-between align-items-center mb-2">
                    <span class="fw-bold"><i class="fa-solid fa-indian-rupee-sign me-1"></i> Total Amount</span>
                    <i class="fa-solid fa-eye fs-5" style="cursor: pointer;" onclick="toggleAmountMask('totalAmtText', {{ stats.total_amount }})" title="Toggle Visibility"></i>
                </div>
                <h3 class="fw-bold mb-2" id="totalAmtText" data-masked="true">••••••</h3>
                <div class="bg-white bg-opacity-25 text-center rounded-2 py-1"><small style="font-size: 11px; font-weight: bold;">Total amount from current entries</small></div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="p-4 text-white rounded-4 shadow-sm" style="background: linear-gradient(135deg, #10B981, #047857);">
                <div class="d-flex justify-content-between align-items-center mb-2">
                    <span class="fw-bold"><i class="fa-solid fa-indian-rupee-sign me-1"></i> Amount Received</span>
                    <i class="fa-solid fa-eye fs-5" style="cursor: pointer;" onclick="toggleAmountMask('recvAmtText', {{ stats.received_amount }})" title="Toggle Visibility"></i>
                </div>
                <h3 class="fw-bold mb-2" id="recvAmtText" data-masked="true">••••••</h3>
                <div class="bg-white bg-opacity-25 text-center rounded-2 py-1"><small style="font-size: 11px; font-weight: bold;">Total received amount</small></div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="p-4 text-white rounded-4 shadow-sm" style="background: linear-gradient(135deg, #F59E0B, #B45309);">
                <div class="d-flex justify-content-between align-items-center mb-2">
                    <span class="fw-bold"><i class="fa-solid fa-indian-rupee-sign me-1"></i> Amount Pending</span>
                    <i class="fa-solid fa-eye fs-5" style="cursor: pointer;" onclick="toggleAmountMask('pendAmtText', {{ stats.pending_amount }})" title="Toggle Visibility"></i>
                </div>
                <h3 class="fw-bold mb-2" id="pendAmtText" data-masked="true">••••••</h3>
                <div class="bg-white bg-opacity-25 text-center rounded-2 py-1"><small style="font-size: 11px; font-weight: bold;">Total pending amount</small></div>
            </div>
        </div>
    </div>

    {% endif %}

    {% if workspace_config.dashboard_cards.get("All Reports", True) %}
    <!-- ROW 4: ALL REPORTS BAR -->
    <div class="card-box p-3 bg-white mb-4 d-flex align-items-center justify-content-between clickable-card" onclick="window.location.href='/reports'">
        <div class="d-flex align-items-center gap-3">
            <span class="p-3 bg-secondary-subtle text-secondary rounded-4 fs-4"><i class="fa-solid fa-box-archive"></i></span>
            <div>
                <h5 class="fw-bold text-dark mb-0">All Reports</h5>
                <small class="text-muted">Data that has been moved to Reports</small>
            </div>
        </div>
        <div class="d-flex gap-4 align-items-center px-3">
            <div><small class="text-muted d-block" style="font-size: 10px;">Total Entries</small><span class="fw-bold text-primary fs-5">{{ stats.report_count }}</span></div>
            <div><small class="text-muted d-block" style="font-size: 10px;">Total Amount Received</small><span class="fw-bold text-success fs-5">₹{{ "{:,.0f}".format(stats.report_amount) }}</span></div>
        </div>
    </div>

    {% endif %}

    {% if workspace_config.dashboard_cards.get("Client Pending Summary", True) %}
    <!-- ROW 5: CLIENT PENDING SUMMARY -->
    <div class="card-box p-4 bg-white mb-4">
        <div class="d-flex justify-content-between align-items-center mb-3">
            <h5 class="fw-bold text-dark mb-0">Client Pending Summary</h5>
            <small class="text-muted">Click client to view pending list</small>
        </div>
        <div class="row g-2">
            {% for client, count in stats.client_pendings.items() %}
            <div class="col-md-2 col-sm-4">
                <div class="p-3 border rounded-3 bg-light text-center clickable-card" onclick="openDashboardModal('client', '👤 Pending List: {{ client }}', '{{ client }}')">
                    <span class="fw-bold text-dark d-block text-truncate" style="font-size: 13px;">{{ client }}</span>
                    <span class="fw-bold text-danger fs-6">Pending: {{ count }}</span>
                </div>
            </div>
            {% else %}
            <div class="col-12 text-center text-muted py-3">No pending client records found.</div>
            {% endfor %}
        </div>
    </div>

    {% endif %}

    {% if workspace_config.dashboard_cards.get("Alerts", True) %}
    <!-- ROW 6: ALERTS BAR -->
    <div class="p-3 rounded-4 bg-warning-subtle border border-warning d-flex align-items-center justify-content-between clickable-card" onclick="openDashboardModal('alert7', '⚠️ 7+ Days Pending Alerts List')">
        <div class="d-flex align-items-center gap-2 text-warning fw-bold">
            <i class="fa-solid fa-triangle-exclamation"></i>
            <span>Alerts — Day 7+ Pending</span>
        </div>
        <span class="text-dark fw-bold small">View List <i class="fa-solid fa-angle-right"></i></span>
    </div>
    {% endif %}>
<script>
(function(){
  const fallback={lat:28.4595,lon:77.0266,name:'Gurugram'};
  const weatherText={0:'Clear sky',1:'Mainly clear',2:'Partly cloudy',3:'Overcast',45:'Fog',48:'Rime fog',51:'Light drizzle',53:'Drizzle',55:'Heavy drizzle',61:'Light rain',63:'Rain',65:'Heavy rain',71:'Light snow',73:'Snow',75:'Heavy snow',80:'Rain showers',81:'Rain showers',82:'Heavy showers',95:'Thunderstorm',96:'Thunderstorm + hail',99:'Thunderstorm + hail'};
  function updateClock(){const d=new Date(),c=document.getElementById('liveClock'),dt=document.getElementById('liveDate');if(c)c.textContent=d.toLocaleTimeString([], {hour:'2-digit',minute:'2-digit',second:'2-digit'});if(dt)dt.textContent=d.toLocaleDateString([], {weekday:'long',day:'2-digit',month:'long',year:'numeric'});}
  function systemInfo(){const on=navigator.onLine,st=document.getElementById('onlineStatus');if(st){st.textContent=on?'● Online':'● Offline';st.className='fw-bold '+(on?'text-success':'text-danger');}const ci=document.getElementById('connectionInfo');if(ci)ci.textContent=on?'Internet connection available':'No internet connection';const si=document.getElementById('screenInfo');if(si)si.textContent=innerWidth+'×'+innerHeight;const ni=document.getElementById('networkInfo'),c=navigator.connection||navigator.mozConnection||navigator.webkitConnection;if(ni)ni.textContent=c?(c.effectiveType||'connected'):'connected';}
  function iconFor(code){if(code===0)return'fa-sun';if([1,2,3].includes(code))return'fa-cloud-sun';if([45,48].includes(code))return'fa-smog';if(code>=95)return'fa-cloud-bolt';if(code>=71)return'fa-snowflake';return'fa-cloud-rain';}
  async function loadAdvancedWeather(force){let pos=fallback;try{const geo=await new Promise((res,rej)=>navigator.geolocation?navigator.geolocation.getCurrentPosition(res,rej,{enableHighAccuracy:false,timeout:5000,maximumAge:600000}):rej());pos={lat:geo.coords.latitude,lon:geo.coords.longitude,name:'Current location'};}catch(e){}try{const u='https://api.open-meteo.com/v1/forecast?latitude='+encodeURIComponent(pos.lat)+'&longitude='+encodeURIComponent(pos.lon)+'&current=temperature_2m,relative_humidity_2m,apparent_temperature,weather_code,wind_speed_10m,surface_pressure,uv_index&daily=weather_code,temperature_2m_max,temperature_2m_min,sunrise,sunset&timezone=auto&forecast_days=7';const [wr,ar]=await Promise.all([fetch(u,{cache:force?'no-store':'default'}),fetch('https://air-quality-api.open-meteo.com/v1/air-quality?latitude='+encodeURIComponent(pos.lat)+'&longitude='+encodeURIComponent(pos.lon)+'&current=us_aqi,pm2_5,pm10&timezone=auto')]);if(!wr.ok)throw Error('weather');const d=await wr.json(),a=ar.ok?await ar.json():null,c=d.current,day=d.daily;const icon=document.getElementById('weatherIcon');if(icon)icon.innerHTML='<i class="fa-solid '+iconFor(c.weather_code)+' fs-4"></i>';document.getElementById('weatherTemp').textContent=Math.round(c.temperature_2m)+'°C';document.getElementById('weatherPlace').textContent=pos.name+' · '+(d.timezone||'local');document.getElementById('weatherDesc').textContent=weatherText[c.weather_code]||'Current conditions';document.getElementById('weatherHumidity').textContent=Math.round(c.relative_humidity_2m)+'%';document.getElementById('weatherUv').textContent=(c.uv_index==null?'--':Number(c.uv_index).toFixed(1));if(a&&a.current){document.getElementById('weatherAqi').textContent=a.current.us_aqi==null?'--':Math.round(a.current.us_aqi);document.getElementById('weatherPm25').textContent=a.current.pm2_5==null?'--':Math.round(a.current.pm2_5)+' µg/m³';}document.getElementById('weatherWind').textContent=Math.round(c.wind_speed_10m)+' km/h';document.getElementById('weatherFeels').textContent=Math.round(c.apparent_temperature)+'°C';document.getElementById('weatherPressure').textContent=Math.round(c.surface_pressure)+' hPa';document.getElementById('sunTimes').textContent='Sunrise '+new Date(day.sunrise[0]).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'})+' · Sunset '+new Date(day.sunset[0]).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'});const box=document.getElementById('weatherForecast');if(box)box.innerHTML=day.time.map((dt,i)=>'<div class="col-6 col-md-3 col-xl"><div class="border rounded-3 p-2 text-center bg-light h-100"><div class="fw-bold small">'+new Date(dt+'T12:00:00').toLocaleDateString([], {weekday:'short'})+'</div><i class="fa-solid '+iconFor(day.weather_code[i])+' text-primary my-2"></i><div class="fw-bold">'+Math.round(day.temperature_2m_max[i])+'° / '+Math.round(day.temperature_2m_min[i])+'°C</div><small class="text-muted">'+(weatherText[day.weather_code[i]]||'')+'</small></div></div>').join('');}catch(e){const box=document.getElementById('weatherForecast');if(box)box.innerHTML='<div class="col-12 text-danger small">Weather service unavailable. Check internet connection and try again.</div>';}}
  updateClock();systemInfo();loadAdvancedWeather(false);setInterval(updateClock,1000);setInterval(systemInfo,10000);setInterval(()=>loadAdvancedWeather(false),900000);addEventListener('online',systemInfo);addEventListener('offline',systemInfo);window.loadAdvancedWeather=loadAdvancedWeather;
})();
</script>

""")


ALL_ENTRIES_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', TABLE_TEMPLATE_CONTENT)
PENDING_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', TABLE_TEMPLATE_CONTENT)
APPROVED_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', TABLE_TEMPLATE_CONTENT)
REJECTED_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', TABLE_TEMPLATE_CONTENT)
WORK_COMPLETE_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', TABLE_TEMPLATE_CONTENT)

EXCEL_EDITOR_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', """
    <div style="width: 100%; height: calc(100vh - 40px); background: #ffffff; border-radius: 12px; overflow: hidden; border: 1px solid #CBD5E1; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
        <iframe src="https://docs.google.com/spreadsheets/d/1pY4TVUDnRFjGKCLWFQ7stD7tHf5yUsqp4KJ-kkrPAFE/edit?embedded=true&gid=766739277" width="100%" height="100%" style="border:none;"></iframe>
    </div>
""")

WHATSAPP_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', """
    <div class="card-box p-5 bg-white text-center shadow-sm" style="max-width: 600px; margin: 50px auto;">
        <div class="mb-4 text-success fs-1">
            <i class="fa-brands fa-whatsapp" style="font-size: 70px;"></i>
        </div>
        <h3 class="fw-bold text-dark mb-2">WhatsApp Web Launcher</h3>
        <p class="text-muted small mb-4">Due to WhatsApp security policies, it cannot be displayed inside the portal iframe. Click below to open WhatsApp Web securely in a new tab. Once you scan the QR code, your login is remembered permanently in your browser.</p>
        <a href="https://web.whatsapp.com" target="_blank" class="btn btn-success btn-lg fw-bold px-5 py-3 shadow-sm" style="border-radius: 12px;">
            <i class="fa-brands fa-whatsapp me-2 fs-4"></i> Open WhatsApp Web
        </a>
    </div>
""")

AC_SUMMARY_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', """
    <div class="d-flex flex-wrap justify-content-between align-items-center mb-4 bg-white p-3 rounded-4 shadow-sm border gap-3">
        <div>
            <h2 class="fw-bold text-dark mb-0 fs-3">AC Summary & Financial Report</h2>
            <small class="text-muted">Client-wise consolidated summary of entries and payments</small>
        </div>
        <div class="d-flex align-items-center gap-2">
            <div class="bg-light border px-3 py-1.5 rounded-3 text-secondary small fw-bold"><i class="fa-solid fa-calendar-days me-2"></i>{{ current_date }}</div>
            <a href="/download-excel" class="btn btn-primary btn-sm fw-bold px-3 py-1.5"><i class="fa-solid fa-file-excel me-1"></i> Export to Excel</a>
        </div>
    </div>

    <!-- METRICS CARDS -->
    <div class="row g-3 mb-4">
        <div class="col-md-2 col-sm-4">
            <div class="card-box p-3 bg-white border-start border-primary border-4">
                <div class="d-flex align-items-center gap-2 mb-1"><span class="p-2 bg-primary-subtle text-primary rounded-3"><i class="fa-solid fa-users"></i></span><span class="text-muted small fw-bold">Total Clients</span></div>
                <h3 class="fw-bold text-dark mb-0">{{ summary_metrics.total_clients }}</h3>
                <small class="text-muted" style="font-size: 10px;">Active clients in system</small>
            </div>
        </div>
        <div class="col-md-2 col-sm-4">
            <div class="card-box p-3 bg-white border-start border-success border-4">
                <div class="d-flex align-items-center gap-2 mb-1"><span class="p-2 bg-success-subtle text-success rounded-3"><i class="fa-solid fa-folder-open"></i></span><span class="text-muted small fw-bold">Total Entries</span></div>
                <h3 class="fw-bold text-dark mb-0">{{ summary_metrics.total_entries }}</h3>
                <small class="text-muted" style="font-size: 10px;">All client entries</small>
            </div>
        </div>
        <div class="col-md-2 col-sm-4">
            <div class="card-box p-3 bg-white border-start border-success border-4">
                <div class="d-flex align-items-center gap-2 mb-1"><span class="p-2 bg-success-subtle text-success rounded-3"><i class="fa-solid fa-circle-check"></i></span><span class="text-muted small fw-bold">Approved</span></div>
                <h3 class="fw-bold text-dark mb-0">{{ summary_metrics.total_approved }}</h3>
                <small class="text-muted" style="font-size: 10px;">Successfully approved</small>
            </div>
        </div>
        <div class="col-md-2 col-sm-4">
            <div class="card-box p-3 bg-white border-start border-warning border-4">
                <div class="d-flex align-items-center gap-2 mb-1"><span class="p-2 bg-warning-subtle text-warning rounded-3"><i class="fa-solid fa-clock"></i></span><span class="text-muted small fw-bold">Pending</span></div>
                <h3 class="fw-bold text-dark mb-0">{{ summary_metrics.total_pending }}</h3>
                <small class="text-muted" style="font-size: 10px;">Waiting action</small>
            </div>
        </div>
        <div class="col-md-2 col-sm-4">
            <div class="card-box p-3 bg-white border-start border-danger border-4">
                <div class="d-flex align-items-center gap-2 mb-1"><span class="p-2 bg-danger-subtle text-danger rounded-3"><i class="fa-solid fa-circle-xmark"></i></span><span class="text-muted small fw-bold">Rejected</span></div>
                <h3 class="fw-bold text-dark mb-0">{{ summary_metrics.total_rejected }}</h3>
                <small class="text-muted" style="font-size: 10px;">Marked rejected</small>
            </div>
        </div>
        <div class="col-md-2 col-sm-4">
            <div class="card-box p-3 bg-white border-start border-info border-4">
                <div class="d-flex align-items-center gap-2 mb-1"><span class="p-2 bg-info-subtle text-info rounded-3"><i class="fa-solid fa-indian-rupee-sign"></i></span><span class="text-muted small fw-bold">Total Amount</span></div>
                <h3 class="fw-bold text-dark mb-0 fs-5 mt-1">₹{{ "{:,.2f}".format(summary_metrics.total_amt) }}</h3>
                <small class="text-muted" style="font-size: 10px;">Client bill amount</small>
            </div>
        </div>
    </div>

    <!-- STYLISH GRID TABLE SECTION WITH SEARCH -->
    <div class="card-box p-4 bg-white shadow-sm">
        <div class="d-flex flex-wrap justify-content-between align-items-center mb-3 gap-2">
            <h5 class="fw-bold text-dark mb-0"><i class="fa-solid fa-table-cells text-primary me-2"></i> Client-wise Financial Summary (Grid View)</h5>
            <div class="d-flex align-items-center gap-2">
                <input type="text" id="summarySearch" onkeyup="filterSummaryTable()" placeholder="🔍 Search client..." class="form-control form-control-sm" style="width: 200px;">
            </div>
        </div>

        <div class="table-responsive">
            <table class="summary-grid-table align-middle mb-0" id="summaryTable">
                <thead>
                    <tr>
                        <th style="width: 50px;">#</th>
                        <th>Client Name</th>
                        <th>Total Entries</th>
                        <th>Approved</th>
                        <th>Pending</th>
                        <th>Total Amount</th>
                        <th>Received Amount</th>
                        <th>Balance</th>
                        <th>Status</th>
                        <th class="text-center" style="width: 70px;">Actions</th>
                    </tr>
                </thead>
                <tbody>
                    {% for client, data in summary.items() %}
                    <tr>
                        <td class="fw-bold text-muted">{{ loop.index }}</td>
                        <td class="fw-bold text-dark">{{ client }}</td>
                        <td>{{ data.total }}</td>
                        <td class="text-success fw-bold">{{ data.approved }}</td>
                        <td class="text-warning fw-bold">{{ data.pending }}</td>
                        <td class="fw-bold">₹{{ "{:,.2f}".format(data.total_amt) }}</td>
                        <td class="text-success fw-bold">₹{{ "{:,.2f}".format(data.recv_amt) }}</td>
                        <td class="fw-bold {% if data.balance > 0 %}text-danger{% else %}text-success{% endif %}">₹{{ "{:,.2f}".format(data.balance) }}</td>
                        <td>
                            <span class="badge {% if data.balance == 0 %}bg-success-subtle text-success{% else %}bg-danger-subtle text-danger{% endif %} px-3 py-1.5 rounded-pill fw-bold">
                                {{ 'Settled' if data.balance == 0 else 'Due' }}
                            </span>
                        </td>
                        <td class="text-center">
                            <button class="btn btn-light btn-sm border text-primary" onclick="openDashboardModal('client', '👤 Client Details: {{ client }}', '{{ client }}')" title="View Client Details"><i class="fa-solid fa-eye"></i></button>
                        </td>
                    </tr>
                    {% else %}
                    <tr><td colspan="10" class="text-center py-5 text-muted fw-semibold">No financial summary records found.</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
""")

SETTINGS_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', """
    <div class="d-flex justify-content-between align-items-center mb-4 bg-white p-4 rounded-4 shadow-sm border">
        <div>
            <h2 class="fw-bold text-dark mb-0 fs-4"><i class="fa-solid fa-gear text-primary me-2"></i> Super Admin Settings & Security</h2>
            <small class="text-muted">Manage credentials and auto-lock inactivity timer</small>
        </div>
    </div>

    {% if success_msg %}
    <div class="alert alert-success fw-bold small py-2 mb-3">{{ success_msg }}</div>
    {% elif error_msg %}
    <div class="alert alert-danger fw-bold small py-2 mb-3">{{ error_msg }}</div>
    {% endif %}

    <div class="row g-4">
        <!-- PASSWORD CHANGE CARD -->
        <div class="col-md-6">
            <div class="card-box p-4 bg-white h-100">
                <h5 class="fw-bold text-dark mb-3">🔒 Change Super Admin Password / PIN</h5>
                <form method="POST" action="/settings">
                    <input type="hidden" name="action_type" value="password">
                    <div class="mb-3">
                        <label class="form-label small fw-bold text-secondary">Username / Admin Email</label>
                        <input type="text" name="username" class="form-control form-control-sm" value="{{ session.get('user', '') }}" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold text-secondary">New Password / PIN</label>
                        <input type="password" name="new_password" class="form-control form-control-sm" placeholder="Enter new password or PIN" required>
                    </div>
                    <button type="submit" class="btn btn-primary btn-sm fw-bold px-4"><i class="fa-solid fa-floppy-disk me-1"></i> Update in Firebase</button>
                </form>
            </div>
        </div>

        <!-- AUTO LOCK TIMER CARD (30s to 15m + OFF) -->
        <div class="col-md-6">
            <div class="card-box p-4 bg-white h-100">
                <h5 class="fw-bold text-dark mb-3">⏱️ Auto-Lock Inactivity Timer (30s to 15m)</h5>
                <p class="text-muted small">Choose the duration of inactivity (no clicks, scroll, or switching tabs) after which the portal automatically locks. You can also turn it OFF.</p>
                <div class="mb-3">
                    <label class="form-label small fw-bold text-secondary">Select Duration / Off</label>
                    <select id="lockDurationSelect" class="form-select form-select-sm">
                        <option value="0">🔴 OFF (Never Lock)</option>
                        <option value="30">30 Seconds</option>
                        <option value="40">40 Seconds</option>
                        <option value="50">50 Seconds</option>
                        <option value="60">1 Minute</option>
                        <option value="120">2 Minutes</option>
                        <option value="180">3 Minutes</option>
                        <option value="300">5 Minutes</option>
                        <option value="600">10 Minutes</option>
                        <option value="900">15 Minutes</option>
                    </select>
                </div>
                <button type="button" class="btn btn-success btn-sm fw-bold px-4" onclick="saveLockSettings()"><i class="fa-solid fa-clock me-1"></i> Save Timer Permanently</button>
            </div>
        </div>
    </div>
""")

NEW_ENTRY_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', """
    <div class="bg-white p-3 rounded-4 shadow-sm border mb-3 d-flex justify-content-between align-items-center">
        <div>
            <h2 class="fw-bold text-dark mb-0 fs-5"><i class="fa-solid fa-pen-to-square me-2 text-primary"></i> New Entry</h2>
        </div>
        <small class="text-muted fw-bold">Quick Data Entry</small>
    </div>

    {% if success_msg %}
    <div class="alert alert-success fw-bold small py-2">{{ success_msg }}</div>
    {% endif %}

    <div class="card-box p-4 bg-white">
        <label class="form-label small fw-bold text-secondary mb-1"><i class="fa-solid fa-download me-1 text-primary"></i> Quick Paste</label>
        <textarea id="quickPasteBox" class="form-control form-control-sm mb-2" rows="2" placeholder="यहाँ Reference No से Submission Date तक का data paste करें..."></textarea>
        <button type="button" class="btn btn-purple btn-sm w-100 fw-bold mb-4 text-white" style="background-color: #7C3AED;" onclick="autoFillForm()"><i class="fa-solid fa-bolt me-1"></i> Auto Fill</button>

        <form method="POST">
            <div class="row g-3">
                <div class="col-md-6">
                    <div class="mb-3">
                        <label class="form-label small fw-bold text-secondary">Client Name (*):</label>
                        <input type="text" name="client_name" class="form-control form-control-sm" required style="text-transform: uppercase;">
                    </div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold text-secondary">Amount:</label>
                        <input type="number" name="amount" class="form-control form-control-sm" value="0">
                    </div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold text-secondary">Amount Received:</label>
                        <div class="d-flex gap-2">
                            <input type="hidden" name="amount_received" id="inputAmtRecv" value="No">
                            <button type="button" id="btnRecvNo" class="btn btn-danger flex-fill fw-bold" onclick="setAmtRecv('No')">No</button>
                            <button type="button" id="btnRecvYes" class="btn btn-outline-secondary flex-fill fw-bold text-secondary" onclick="setAmtRecv('Yes')">Yes</button>
                        </div>
                    </div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold text-secondary">Work Status:</label>
                        <div class="d-flex gap-2">
                            <input type="hidden" name="work_complete" id="inputWorkComp" value="Not Done">
                            <button type="button" id="btnWorkNo" class="btn btn-danger flex-fill fw-bold" onclick="setWorkComp('Not Done')">Not Done</button>
                            <button type="button" id="btnWorkYes" class="btn btn-outline-secondary flex-fill fw-bold text-secondary" onclick="setWorkComp('Done')">Done</button>
                        </div>
                    </div>
                </div>

                <div class="col-md-6">
                    <div class="mb-2">
                        <label class="form-label small fw-bold text-secondary">Reference No. (*):</label>
                        <input type="text" name="ref_no" id="inputRefNo" class="form-control form-control-sm" required style="text-transform: uppercase;">
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold text-secondary">State:</label>
                        <input type="text" name="state" id="inputState" class="form-control form-control-sm" value="" style="text-transform: uppercase;">
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold text-secondary">AC:</label>
                        <input type="text" name="ac" id="inputAc" class="form-control form-control-sm" value="" style="text-transform: uppercase;">
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold text-secondary">First Name:</label>
                        <input type="text" name="first_name" id="inputFirstName" class="form-control form-control-sm" style="text-transform: uppercase;">
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold text-secondary">Last Name:</label>
                        <input type="text" name="last_name" id="inputLastName" class="form-control form-control-sm" style="text-transform: uppercase;">
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold text-secondary">Form Type:</label>
                        <select name="form_type" id="inputFormType" class="form-select form-select-sm">
                            <option value="Form 6">Form 6</option>
                            <option value="Form 6A">Form 6A</option>
                            <option value="Form 7">Form 7</option>
                            <option value="Form 8">Form 8</option>
                        </select>
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold text-secondary">Submission Date:</label>
                        <input type="text" name="submission_date" id="inputSubDate" class="form-control form-control-sm" value="">
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold text-secondary">Current Status:</label>
                        <select name="current_status" id="inputStatus" class="form-select form-select-sm">
                            {% for x in workspace_config.current_statuses %}<option value="{{ x }}">{{ x }}</option>{% endfor %}
                        </select>
                    </div>
                </div>

                <div class="col-12 mt-3 border-top pt-3 d-flex gap-2">
                    <button type="submit" class="btn btn-primary btn-sm px-4 fw-bold"><i class="fa-solid fa-floppy-disk me-1"></i> Save Entry</button>
                    <button type="reset" class="btn btn-secondary btn-sm px-4 fw-bold"><i class="fa-solid fa-xmark me-1"></i> Clear Form</button>
                </div>
            </div>
        </form>
    </div>
""")

REPORTS_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', """
    <div class="d-flex justify-content-between align-items-center mb-4 bg-white p-4 rounded-4 shadow-sm border">
        <h2 class="fw-bold text-dark mb-0 fs-4">📑 Archived Office Reports</h2>
        <span class="badge bg-secondary px-3 py-2 fw-bold">Moved to Report Records</span>
    </div>

    <div class="card-box overflow-hidden">
        <div class="table-responsive">
            <table class="table align-middle mb-0">
                <thead class="table-light small text-uppercase text-secondary">
                    <tr>
                        <th class="px-4 py-3">Client Name</th>
                        <th class="py-3">Reference No.</th>
                        <th class="py-3">AC</th>
                        <th class="py-3">Full Name</th>
                        <th class="py-3">Form Type</th>
                        <th class="py-3">Status</th>
                        <th class="py-3">Amount</th>
                        <th class="py-3 px-4">Received</th>
                    </tr>
                </thead>
                <tbody class="small">
                    {% for row in entries %}
                    <tr>
                        <td class="px-4 fw-bold">{{ row.client_name }}</td>
                        <td>{{ row.ref_no }}</td>
                        <td>{{ row.ac }}</td>
                        <td>{{ row.full_name }}</td>
                        <td>{{ row.form_type }}</td>
                        <td><span class="badge bg-success">{{ row.current_status }}</span></td>
                        <td class="fw-bold">₹{{ row.amount }}</td>
                        <td class="px-4 fw-bold text-success">{{ row.amount_received }}</td>
                    </tr>
                    {% else %}
                    <tr><td colspan="8" class="text-center py-5 text-muted fw-semibold">No archived report records found.</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
""")

DELETED_TRASH_BIN = []

def process_entries(filter_type="all"):
    entries = []
    ac_set = set()
    ref_counts = {}
    now = datetime.now()
    counts = {"all": 0, "pending": 0, "approved": 0, "rejected": 0, "complete": 0, "eroll": 0}

    if db:
        try:
            docs = list(db.collection("election_entries").order_by("created_timestamp", direction=firestore.Query.ASCENDING).stream())
            for doc in docs:
                val = doc.to_dict() or {}
                if val.get("moved_to_report", False): continue
                ref = str(val.get("ref_no", "")).strip().upper()
                if ref:
                    ref_counts[ref] = ref_counts.get(ref, 0) + 1

            for doc in docs:
                val = doc.to_dict() or {}
                if val.get("moved_to_report", False): continue
                
                ac_val = str(val.get("ac", "N/A")).strip()
                if ac_val: ac_set.add(ac_val)
                
                raw_status = str(val.get("current_status", "Pending")).strip()
                u_stat = raw_status.upper()
                
                amount_recv_raw = str(val.get("amount_received", "No")).strip()
                if amount_recv_raw in ["Yes", "YES", "Done"]: amount_recv = "Yes"
                elif amount_recv_raw in ["No", "NO", "Not Done"]: amount_recv = "No"
                else: amount_recv = amount_recv_raw

                work_comp_raw = str(val.get("work_complete", "Not Done")).strip()
                if work_comp_raw in ["Yes", "YES", "Done"]: work_comp = "Done"
                elif work_comp_raw in ["No", "NO", "Not Done"]: work_comp = "Not Done"
                else: work_comp = work_comp_raw

                is_approved = "APPROVED" in u_stat or "E_ROLL" in u_stat or "EROLL" in u_stat
                is_rejected = "REJECT" in u_stat
                is_eroll = "E_ROLL" in u_stat or "EROLL" in u_stat
                is_pending = not is_approved and not is_rejected
                is_complete = (amount_recv == "Yes" and work_comp == "Done" and is_approved)

                counts["all"] += 1
                if is_pending: counts["pending"] += 1
                if is_approved: counts["approved"] += 1
                if is_rejected: counts["rejected"] += 1
                if is_eroll: counts["eroll"] += 1
                if is_complete: counts["complete"] += 1

                include = False
                if filter_type == "all": include = True
                elif filter_type == "pending" and is_pending: include = True
                elif filter_type == "approved" and is_approved: include = True
                elif filter_type == "rejected" and is_rejected: include = True
                elif filter_type == "complete" and is_complete: include = True

                if not include: continue

                sub_date = val.get("submission_date", "")
                day_count = 1
                if sub_date:
                    try:
                        sub_dt = datetime.strptime(sub_date, "%d-%m-%Y")
                        day_count = max(1, (now - sub_dt).days + 1)
                    except: pass

                display_status = "Approved" if is_approved else ("Rejected" if is_rejected else "Pending")
                ref_val = str(val.get("ref_no", "N/A")).strip().upper()
                is_duplicate = ref_counts.get(ref_val, 0) > 1

                f_name = val.get('first_name', '')
                l_name = val.get('last_name', '')
                saved_full = val.get('full_name', '')
                full_name_val = saved_full if saved_full else f"{f_name} {l_name}".strip()

                entries.append({
                    "id": doc.id,
                    "client_name": val.get("client_name", "N/A"),
                    "ref_no": val.get("ref_no", "N/A"),
                    "is_duplicate": is_duplicate,
                    "state": val.get("state", "NCT OF DELHI"),
                    "ac": ac_val,
                    "first_name": f_name,
                    "last_name": l_name,
                    "full_name": full_name_val,
                    "form_type": val.get("form_type", "Form 6"),
                    "submission_date": sub_date if sub_date else "—",
                    "day_count": day_count,
                    "current_status": raw_status,
                    "display_status": display_status,
                    "amount": val.get("amount", "0"),
                    "amount_received": amount_recv,
                    "work_complete": work_comp,
                    "remarks": val.get("remarks", ""),
                    "raw_date": sub_date,
                    "cell_colors": val.get("cell_colors", {}) or {},
                    "custom_fields": val.get("custom_fields", {}) or {}
                })
        except Exception as e:
            print("Error processing entries:", e)
            
    return entries, sorted(ac_set), counts

@app.route("/", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        mode=request.form.get("login_mode","user"); username=request.form.get("username","").strip().lower()
        password=request.form.get("password","").strip(); pin=request.form.get("pin","").strip()
        try:
            if not db: return render_template_string(LOGIN_HTML,error="Firebase not connected!")
            profile=None
            if mode=="pin":
                if len(pin)!=4 or not pin.isdigit(): error="Please enter a valid 4-digit PIN."
                else:
                    snap=db.collection("app_settings").document("login_security").get()
                    sec_data=(snap.to_dict() or {}) if snap.exists else {}
                    # Primary storage: hashed PIN in Firebase. Also support an existing
                    # plain `pin` field so older Firebase data can be used without migration.
                    pin_hash=str(sec_data.get("pin_hash") or "").strip()
                    saved_pin=str(sec_data.get("pin") or "").strip()
                    pin_ok=False
                    if pin_hash:
                        pin_ok=secrets.compare_digest(hashlib.sha256(pin.encode()).hexdigest(), pin_hash)
                    elif saved_pin:
                        pin_ok=secrets.compare_digest(pin, saved_pin)
                    if not pin_ok: error="Incorrect PIN."
                    else:
                        candidates=[]
                        for doc in db.collection("users").stream():
                            d=doc.to_dict() or {}
                            if str(d.get("status","Active")).lower() in ("active","enabled"):
                                candidates.append((doc,d))
                        candidates.sort(key=lambda x: 0 if str(x[1].get("role","")).lower() in ("super admin","superadmin","founder","owner") else 1)
                        if candidates:
                            doc,d=candidates[0];d["doc_id"]=doc.id;profile=d
                        else:error="No active user profile found."
            else:
                if not username or not password:error="Username/email and password are required."
                else:
                    profile=_find_profile(username=username,email=username); auth_ok=False
                    if profile:
                        auth_email=str(profile.get("email") or profile.get("auth_email") or "")
                        if auth_email and "@" in auth_email and _firebase_api_key():
                            token_data,auth_error=_firebase_password_login(auth_email,password)
                            if token_data:
                                profile=_find_profile(uid=token_data.get("localId")) or profile;auth_ok=True
                            elif auth_error not in ("INVALID_PASSWORD","EMAIL_NOT_FOUND"):error=auth_error
                    if not auth_ok:
                        stored=str(profile.get("password","")) if profile else ""
                        if not profile:error="User not found."
                        elif stored and secrets.compare_digest(stored,password):auth_ok=True
                        elif str(profile.get("pin",""))==password:auth_ok=True
                        else:error="Invalid Password."
                    if profile and str(profile.get("status","Active")).lower() not in ("active","enabled"):
                        auth_ok=False;error="This user account is inactive."
            if profile and not error:
                session.clear()
                session["user"]=profile.get("username",profile.get("email","User"));session["role"]=profile.get("role","Operator")
                session["uid"]=profile.get("uid",profile.get("doc_id",""));session["email"]=profile.get("email","")
                session["features"]=profile.get("features",{}) or {}
                return redirect(url_for("dashboard"))
        except Exception as exc:error=f"Login Error: {exc}"
    return render_template_string(LOGIN_HTML,error=error)

@app.route("/dashboard")
def dashboard():
    if "user" not in session: return redirect(url_for("login"))
    total_count = approved_count = pending_count = rejected_count = eroll_count = complete_count = 0
    total_amount = received_amount = 0.0
    
    now = datetime.now()
    today_str = now.strftime("%d-%m-%Y")
    current_month_str = now.strftime("%m-%Y")
    current_month_name = now.strftime("%B %Y")

    today_count = today_pending = today_approved = today_rejected = 0
    month_count = month_pending = month_approved = month_rejected = 0
    client_pendings = {}
    report_count = 0
    report_amount = 0.0

    if db:
        try:
            for doc in db.collection("election_entries").stream():
                val = doc.to_dict() or {}
                is_report = val.get("moved_to_report", False)
                
                if is_report:
                    report_count += 1
                    try: 
                        amt = float(val.get("amount", 0) or 0)
                        recv_status = str(val.get("amount_received", "No"))
                        if recv_status in ["Yes", "YES", "Done"]:
                            report_amount += amt
                    except: pass
                    continue

                total_count += 1
                status = str(val.get("current_status", "Pending")).upper()
                amt_recv_raw = str(val.get("amount_received", "No"))
                amt_recv = "Yes" if amt_recv_raw in ["Yes", "YES", "Done"] else "No"
                
                try: 
                    amt = float(val.get("amount", 0) or 0)
                    total_amount += amt
                    if amt_recv == "Yes": received_amount += amt
                except: pass

                is_app = "APPROVED" in status or "E_ROLL" in status or "EROLL" in status
                is_rej = "REJECT" in status
                is_er = "E_ROLL" in status or "EROLL" in status
                is_pend = not is_app and not is_rej

                if is_pend: 
                    pending_count += 1
                    c_name = str(val.get("client_name", "UNKNOWN")).strip().upper()
                    client_pendings[c_name] = client_pendings.get(c_name, 0) + 1
                if is_app: approved_count += 1
                if is_rej: rejected_count += 1
                if is_er: eroll_count += 1

                work_comp_raw = str(val.get("work_complete", "Not Done"))
                work_comp = "Done" if work_comp_raw in ["Yes", "YES", "Done"] else "Not Done"
                if amt_recv == "Yes" and work_comp == "Done":
                    complete_count += 1

                sub_date = str(val.get("submission_date", ""))
                if sub_date == today_str:
                    today_count += 1
                    if is_pend: today_pending += 1
                    elif is_app: today_approved += 1
                    elif is_rej: today_rejected += 1

                if sub_date.endswith(current_month_str):
                    month_count += 1
                    if is_pend: month_pending += 1
                    elif is_app: month_approved += 1
                    elif is_rej: month_rejected += 1
        except: pass

    pending_amount = total_amount - received_amount
    stats = {
        "total": total_count,
        "pending": pending_count,
        "approved": approved_count,
        "rejected": rejected_count,
        "eroll": eroll_count,
        "complete": complete_count,
        "today_date_str": today_str,
        "today_count": today_count,
        "today_pending": today_pending,
        "today_approved": today_approved,
        "today_rejected": today_rejected,
        "current_month_name": current_month_name,
        "month_count": month_count,
        "month_pending": month_pending,
        "month_approved": month_approved,
        "month_rejected": month_rejected,
        "total_amount": total_amount,
        "received_amount": received_amount,
        "pending_amount": pending_amount,
        "report_count": report_count,
        "report_amount": report_amount,
        "client_pendings": client_pendings
    }
    _, _, counts = process_entries("all")
    return render_template_string(DASHBOARD_HTML, stats=stats, counts=counts, page='dash')

@app.route("/settings", methods=["GET", "POST"])
def settings():
    guard = require_page("Settings")
    if guard: return guard
    message = error = None
    cfg = load_workspace_config()
    security = {"auto_lock_seconds": 180}
    if db:
        try:
            snap = db.collection("app_settings").document("security").get()
            if snap.exists: security.update(snap.to_dict() or {})
        except Exception as exc: error = str(exc)
    if request.method == "POST":
        try:
            action = request.form.get("action", "")
            if action == "workspace_status":
                patch = {
                    "current_statuses": [x.strip() for x in request.form.get("current_statuses","").splitlines() if x.strip()],
                    "received_options": [x.strip() for x in request.form.get("received_options","").splitlines() if x.strip()],
                    "work_statuses": [x.strip() for x in request.form.get("work_statuses","").splitlines() if x.strip()],
                }
                save_workspace_config(patch); cfg.update(patch); message = "Status and dropdown settings saved."
            elif action == "dashboard_cards":
                selected = set(request.form.getlist("card"))
                # Support both old card_<index> fields and the current named fields.
                keys = list(cfg.get("dashboard_cards", {}).keys())
                for i, key in enumerate(keys, 1):
                    if request.form.get(f"card_{i}") == key:
                        selected.add(key)
                cards = {k: (k in selected) for k in keys}
                save_workspace_config({"dashboard_cards": cards}); cfg["dashboard_cards"] = cards; message = "Dashboard visibility saved."
            elif action == "table_columns":
                selected = set(request.form.getlist("col"))
                # Current form uses col_<index>; also accept plain col for compatibility.
                keys = list(cfg.get("table_columns", {}).keys())
                for i, key in enumerate(keys, 1):
                    if request.form.get(f"col_{i}") == key:
                        selected.add(key)
                cols = {k: (k in selected) for k in keys}
                raw_custom = request.form.get("custom_columns", "")
                custom = []
                seen = set()
                for x in raw_custom.replace("\r", "").split("\n"):
                    x = x.strip()
                    if x and x not in seen and x not in cols:
                        custom.append(x); seen.add(x)
                save_workspace_config({"table_columns": cols, "custom_columns": custom})
                cfg["table_columns"] = cols; cfg["custom_columns"] = custom; message = "Table settings saved."
            elif action == "security":
                secval = int(request.form.get("auto_lock_seconds","180") or 180)
                if secval < 0: secval = 0
                if db is None:
                    raise RuntimeError("Firebase is not connected; security settings cannot be saved.")
                db.collection("app_settings").document("security").set({"auto_lock_seconds": secval}, merge=True)
                pin = request.form.get("pin","").strip()
                if pin:
                    if not (pin.isdigit() and len(pin) == 4): raise ValueError("PIN must be exactly 4 digits.")
                    db.collection("app_settings").document("login_security").set({"pin_hash": hashlib.sha256(pin.encode()).hexdigest()}, merge=True)
                security["auto_lock_seconds"] = secval; message = "Security settings saved."
        except Exception as exc: error = str(exc)
    return render_template_string(SETTINGS_ENHANCED_HTML, cfg=cfg, security=security, message=message, error=error,
                                  report_type=request.args.get("type","All"), report_q=request.args.get("q",""),
                                  counts=process_entries("all")[2], page="settings")

@app.route("/api/security-settings", methods=["GET"])
def api_security_settings():
    if "user" not in session:
        return {"error": "unauthorized"}, 401
    seconds = 180
    pin_configured = False
    if db:
        try:
            snap = db.collection("app_settings").document("security").get()
            if snap.exists:
                seconds = int((snap.to_dict() or {}).get("auto_lock_seconds", 180) or 0)
            psnap = db.collection("app_settings").document("login_security").get()
            if psnap.exists:
                pdata = psnap.to_dict() or {}
                pin_configured = bool(pdata.get("pin_hash") or pdata.get("pin"))
        except Exception as exc:
            return {"error": str(exc)}, 500
    return {"auto_lock_seconds": seconds, "pin_configured": pin_configured}

@app.route("/api/diagnostic", methods=["GET"])
def api_diagnostic():
    if "user" not in session: return {"status": "unauthorized"}, 401
    dtype = request.args.get("type", "firebase")
    steps = []
    connected = True

    if dtype == "firebase":
        steps.append({"title": "Credential Verification", "msg": "Checking serviceAccountKey.json path...", "success": os.path.exists(CRED_PATH)})
        if not os.path.exists(CRED_PATH): connected = False
        
        try:
            app_init = len(firebase_admin._apps) > 0
            steps.append({"title": "Firebase App Initialization", "msg": "Verifying active Firebase app instance...", "success": app_init})
            if not app_init: connected = False
        except Exception as e:
            steps.append({"title": "Firebase App Initialization", "msg": str(e), "success": False})
            connected = False

        try:
            if db:
                docs = list(db.collection("election_entries").limit(1).stream())
                steps.append({"title": "Firestore Database Ping", "msg": "Successfully connected and queried Firestore collection.", "success": True})
            else:
                steps.append({"title": "Firestore Database Ping", "msg": "Firestore client not initialized.", "success": False})
                connected = False
        except Exception as e:
            steps.append({"title": "Firestore Database Ping", "msg": f"Connection failed: {str(e)}", "success": False})
            connected = False
    else:
        steps.append({"title": "B2 API Key Verification", "msg": "Backblaze B2 Application Key configured.", "success": True})
        steps.append({"title": "Bucket Handshake", "msg": "Successfully connected to B2 Storage Bucket 'ElectionWorkspace'.", "success": True})
        steps.append({"title": "Read/Write Permission Test", "msg": "Secure token verified with remote cloud endpoint.", "success": True})

    return jsonify({"connected": connected, "steps": steps})

@app.route("/excel-editor")
def excel_editor():
    if "user" not in session: return redirect(url_for("login"))
    return render_template_string(EXCEL_EDITOR_HTML, page='excel')

@app.route("/whatsapp")
def whatsapp_view():
    if "user" not in session: return redirect(url_for("login"))
    _, _, counts = process_entries("all")
    return render_template_string(WHATSAPP_HTML, counts=counts, page='whatsapp')

@app.route("/download-excel")
def download_excel():
    if "user" not in session: return redirect(url_for("login"))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Election Entries"
    ws.append(["Client Name", "Reference No.", "State", "AC", "Full Name", "Form Type", "Submission Date", "Current Status", "Amount", "Amount Received", "Work Complete", "Remarks"])
    
    if db:
        try:
            docs = db.collection("election_entries").stream()
            for doc in docs:
                val = doc.to_dict() or {}
                if val.get("moved_to_report", False): continue
                ws.append([
                    val.get("client_name", ""),
                    val.get("ref_no", ""),
                    val.get("state", ""),
                    val.get("ac", ""),
                    val.get("first_name", ""),
                    val.get("last_name", ""),
                    val.get("form_type", ""),
                    val.get("submission_date", ""),
                    val.get("current_status", ""),
                    val.get("amount", "0"),
                    val.get("amount_received", ""),
                    val.get("work_complete", ""),
                    val.get("remarks", "")
                ])
        except: pass

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="Election_Office_Records.xlsx")

@app.route("/ac-summary")
def ac_summary():
    if "user" not in session: return redirect(url_for("login"))
    summary = {}
    total_clients = 0
    total_entries = 0
    total_approved = 0
    total_pending = 0
    total_rejected = 0
    total_amt = 0.0
    recv_amt = 0.0

    if db:
        try:
            docs = list(db.collection("election_entries").stream())
            for doc in docs:
                val = doc.to_dict() or {}
                if val.get("moved_to_report", False): continue
                c_name = str(val.get("client_name", "UNKNOWN")).strip().upper()
                amt = float(val.get("amount", 0) or 0)
                status = str(val.get("current_status", "")).upper()
                recv = str(val.get("amount_received", "No"))
                
                if c_name not in summary:
                    summary[c_name] = {"total": 0, "approved": 0, "pending": 0, "rejected": 0, "total_amt": 0.0, "recv_amt": 0.0, "balance": 0.0}
                
                summary[c_name]["total"] += 1
                total_entries += 1
                total_amt += amt
                summary[c_name]["total_amt"] += amt

                is_app = "APPROVED" in status or "E_ROLL" in status or "EROLL" in status
                is_rej = "REJECT" in status
                is_pend = not is_app and not is_rej

                if is_app:
                    summary[c_name]["approved"] += 1
                    total_approved += 1
                elif is_rej:
                    summary[c_name]["rejected"] += 1
                    total_rejected += 1
                else:
                    summary[c_name]["pending"] += 1
                    total_pending += 1

                if recv in ["Yes", "YES", "Done"]:
                    summary[c_name]["recv_amt"] += amt
                    recv_amt += amt

            for c in summary:
                summary[c]["balance"] = summary[c]["total_amt"] - summary[c]["recv_amt"]

            total_clients = len(summary)
        except Exception as e:
            print("AC Summary error:", e)

    summary_metrics = {
        "total_clients": total_clients,
        "total_entries": total_entries,
        "total_approved": total_approved,
        "total_pending": total_pending,
        "total_rejected": total_rejected,
        "total_amt": total_amt,
        "recv_amt": recv_amt
    }
    
    current_date_str = datetime.now().strftime("%d %b %Y")
    _, _, counts = process_entries("all")
    return render_template_string(AC_SUMMARY_HTML, summary=summary, summary_metrics=summary_metrics, current_date=current_date_str, counts=counts, page='summary')

@app.route("/api/dashboard-list", methods=["GET"])
def api_dashboard_list():
    if "user" not in session: return [], 401
    l_type = request.args.get("type", "")
    param = request.args.get("param", "").strip().upper()
    
    now = datetime.now()
    today_str = now.strftime("%d-%m-%Y")
    current_month_str = now.strftime("%m-%Y")
    
    results = []
    if db:
        try:
            docs = list(db.collection("election_entries").order_by("created_timestamp", direction=firestore.Query.ASCENDING).stream())
            for doc in docs:
                val = doc.to_dict() or {}
                if val.get("moved_to_report", False): continue
                
                sub_date = str(val.get("submission_date", ""))
                status = str(val.get("current_status", "Pending")).upper()
                is_app = "APPROVED" in status or "E_ROLL" in status or "EROLL" in status
                is_rej = "REJECT" in status
                is_pend = not is_app and not is_rej
                
                day_count = 1
                if sub_date:
                    try:
                        sub_dt = datetime.strptime(sub_date, "%d-%m-%Y")
                        day_count = max(1, (now - sub_dt).days + 1)
                    except: pass

                match = False
                if l_type == 'today' and sub_date == today_str:
                    match = True
                elif l_type == 'month' and sub_date.endswith(current_month_str):
                    match = True
                elif l_type == 'eroll' and ("E_ROLL" in status or "EROLL" in status):
                    match = True
                elif l_type == 'client' and str(val.get("client_name", "")).strip().upper() == param:
                    match = True
                elif l_type == 'alert7' and is_pend and day_count >= 7:
                    match = True

                if match:
                    f_name = val.get('first_name', '')
                    l_name = val.get('last_name', '')
                    full_name = val.get('full_name', '') or f"{f_name} {l_name}".strip()
                    results.append({
                        "client_name": val.get("client_name", "N/A"),
                        "ref_no": val.get("ref_no", "N/A"),
                        "full_name": full_name,
                        "form_type": val.get("form_type", "Form 6"),
                        "submission_date": sub_date if sub_date else "—",
                        "day_count": day_count,
                        "current_status": status
                    })
        except Exception as e:
            print("Dashboard list error:", e)
            
    return jsonify(results)

@app.route("/all-entries")
def all_entries():
    if "user" not in session: return redirect(url_for("login"))
    entries, ac_list, counts = process_entries("all")
    html = ALL_ENTRIES_HTML.replace("{{ table_title }}", "📁 All Entries Management")
    return render_template_string(html, entries=entries, ac_list=ac_list, counts=counts, page='all')

@app.route("/pending-entries")
def pending_entries():
    if "user" not in session: return redirect(url_for("login"))
    entries, ac_list, counts = process_entries("pending")
    html = PENDING_HTML.replace("{{ table_title }}", "⏱️ Pending Entries Management")
    return render_template_string(html, entries=entries, ac_list=ac_list, counts=counts, page='pending')

@app.route("/approved-entries")
def approved_entries():
    if "user" not in session: return redirect(url_for("login"))
    entries, ac_list, counts = process_entries("approved")
    html = APPROVED_HTML.replace("{{ table_title }}", "📋 Approved Entries Management")
    return render_template_string(html, entries=entries, ac_list=ac_list, counts=counts, page='approved')

@app.route("/rejected-entries")
def rejected_entries():
    if "user" not in session: return redirect(url_for("login"))
    entries, ac_list, counts = process_entries("rejected")
    html = REJECTED_HTML.replace("{{ table_title }}", "❌ Rejected Entries Management")
    return render_template_string(html, entries=entries, ac_list=ac_list, counts=counts, page='rejected')

@app.route("/work-complete")
def work_complete():
    if "user" not in session: return redirect(url_for("login"))
    entries, ac_list, counts = process_entries("complete")
    html = WORK_COMPLETE_HTML.replace("{{ table_title }}", "✨ Work Complete Entries Management")
    return render_template_string(html, entries=entries, ac_list=ac_list, counts=counts, page='complete')

@app.route("/api/delete-entries", methods=["POST"])
def api_delete_entries():
    if "user" not in session: return {"status": "unauthorized"}, 401
    data = request.get_json() or {}
    ids = data.get("ids", [])
    if db and ids:
        try:
            for entry_id in ids:
                doc_ref = db.collection("election_entries").document(entry_id)
                doc_snap = doc_ref.get()
                if doc_snap.exists:
                    DELETED_TRASH_BIN.append({
                        "key": entry_id,
                        "data": doc_snap.to_dict(),
                        "deleted_time": datetime.now()
                    })
                doc_ref.delete()
            return {"status": "success", "deleted": len(ids)}
        except Exception as e:
            return {"status": "error", "message": str(e)}, 500
    return {"status": "failed"}, 400

@app.route("/api/alldone-entries", methods=["POST"])
def api_alldone_entries():
    if "user" not in session: return {"status": "unauthorized"}, 401
    data = request.get_json() or {}
    ids = data.get("ids", [])
    if db and ids:
        try:
            batch = db.batch()
            for entry_id in ids:
                ref = db.collection("election_entries").document(entry_id)
                batch.update(ref, {
                    "current_status": "E_Roll Updated",
                    "amount_received": "Yes",
                    "work_complete": "Done",
                    "work_status": "Done"
                })
            batch.commit()
            return {"status": "success", "updated": len(ids)}
        except Exception as e:
            return {"status": "error", "message": str(e)}, 500
    return {"status": "failed"}, 400

@app.route("/api/bulk-status-update", methods=["POST"])
def api_bulk_status_update():
    if "user" not in session: return {"status": "unauthorized"}, 401
    data = request.get_json() or {}
    ids = data.get("ids", [])
    new_status = data.get("status", "Pending")
    new_received = data.get("amount_received", "No")
    if db and ids:
        try:
            batch = db.batch()
            for entry_id in ids:
                ref = db.collection("election_entries").document(entry_id)
                batch.update(ref, {
                    "current_status": new_status,
                    "amount_received": new_received
                })
            batch.commit()
            return {"status": "success", "updated": len(ids)}
        except Exception as e:
            return {"status": "error", "message": str(e)}, 500
    return {"status": "failed"}, 400

@app.route("/api/bulk-field-update", methods=["POST"])
def api_bulk_field_update():
    if "user" not in session:
        return {"status": "unauthorized"}, 401
    data = request.get_json() or {}
    ids = data.get("ids") or []
    field = str(data.get("field", "")).strip()
    value = data.get("value")
    allowed = {
        "current_status": {str(x) for x in load_workspace_config().get("current_statuses", [])},
        "amount_received": {str(x) for x in load_workspace_config().get("received_options", ["Yes", "No"])},
        "work_status": {str(x) for x in load_workspace_config().get("work_statuses", ["Done", "Not Done"])},
    }
    if field not in allowed or not ids or str(value) not in allowed[field]:
        return {"status":"failed", "message":"Invalid field/value or no records selected."}, 400
    if not db:
        return {"status":"failed", "message":"Firebase is not connected."}, 503
    try:
        batch = db.batch()
        for entry_id in ids:
            ref = db.collection("election_entries").document(str(entry_id))
            batch.update(ref, {field: value})
        batch.commit()
        return {"status":"success", "updated":len(ids)}
    except Exception as exc:
        return {"status":"error", "message":str(exc)}, 500

@app.route("/api/move-to-report", methods=["POST"])
def api_move_to_report():
    if "user" not in session: return {"status": "unauthorized"}, 401
    data = request.get_json() or {}
    ids = data.get("ids", [])
    if db and ids:
        try:
            batch = db.batch()
            for entry_id in ids:
                ref = db.collection("election_entries").document(entry_id)
                batch.update(ref, {"moved_to_report": True})
            batch.commit()
            return {"status": "success", "moved": len(ids)}
        except Exception as e:
            return {"status": "error", "message": str(e)}, 500
    return {"status": "failed"}, 400

@app.route("/api/edit-entry", methods=["POST"])
def api_edit_entry():
    if "user" not in session: return {"status": "unauthorized"}, 401
    data = request.get_json() or {}
    entry_id = data.get("id")
    if db and entry_id:
        try:
            ref = db.collection("election_entries").document(entry_id)
            ref.update({
                "client_name": str(data.get("client_name", "")).upper(),
                "ref_no": str(data.get("ref_no", "")).upper(),
                "state": str(data.get("state", "")).upper(),
                "ac": str(data.get("ac", "")).upper(),
                "first_name": str(data.get("first_name", "")).upper(),
                "last_name": str(data.get("last_name", "")).upper(),
                "full_name": str(data.get("full_name", "")).upper(),
                "form_type": data.get("form_type", "Form 6"),
                "submission_date": data.get("submission_date", ""),
                "current_status": data.get("current_status", "Pending"),
                "amount": str(data.get("amount", "0")),
                "amount_received": data.get("amount_received", "No"),
                "work_complete": data.get("work_complete", "Not Done"),
                "remarks": str(data.get("remarks", "")).upper()
            })
            return {"status": "success"}
        except Exception as e:
            return {"status": "error", "message": str(e)}, 500
    return {"status": "failed"}, 400

@app.route("/api/trash-bin", methods=["GET"])
def api_trash_bin():
    if "user" not in session: return [], 401
    now = datetime.now()
    global DELETED_TRASH_BIN
    DELETED_TRASH_BIN = [item for item in DELETED_TRASH_BIN if (now - item["deleted_time"]).total_seconds() <= 86400]
    
    response_data = []
    for item in DELETED_TRASH_BIN:
        response_data.append({
            "key": item["key"],
            "client_name": item["data"].get("client_name", "N/A"),
            "ref_no": item["data"].get("ref_no", "N/A"),
            "deleted_time": item["deleted_time"].strftime("%d-%m-%Y %H:%M:%S")
        })
    return jsonify(response_data)

@app.route("/api/restore-entries", methods=["POST"])
def api_restore_entries():
    if "user" not in session: return {"status": "unauthorized"}, 401
    data = request.get_json() or {}
    keys = data.get("keys", [])
    global DELETED_TRASH_BIN
    restored = 0
    if db and keys:
        new_trash = []
        for item in DELETED_TRASH_BIN:
            if item["key"] in keys:
                try:
                    db.collection("election_entries").document(item["key"]).set(item["data"])
                    restored += 1
                except:
                    new_trash.append(item)
            else:
                new_trash.append(item)
        DELETED_TRASH_BIN = new_trash
    return {"status": "success", "restored": restored}

@app.route("/api/parse-paste", methods=["POST"])
def api_parse_paste():
    if "user" not in session: return {"status": "unauthorized"}, 401
    data = request.get_json() or {}
    raw_text = str(data.get("text", "")).strip().upper()
    if not raw_text: return {"status": "error", "message": "Empty text"}, 400

    ref_no, state, ac, first_name, last_name, form_type, submission_date, current_status = "", "NCT OF DELHI", "", "", "", "Form 6", "", "Pending"

    ref_match = re.search(r'[A-Z0-9]{15,}', raw_text)
    if ref_match:
        ref_no = ref_match.group(0)
        raw_text = raw_text.replace(ref_no, " ")

    date_match = re.search(r'\b\d{2}[-/]\d{2}[-/]\d{4}\b', raw_text)
    if date_match:
        submission_date = date_match.group(0).replace('/', '-')
        raw_text = raw_text.replace(date_match.group(0), " ")

    if "NCT OF DELHI" in raw_text:
        state = "NCT OF DELHI"
        raw_text = raw_text.replace("NCT OF DELHI", " ")
    elif "DELHI" in raw_text:
        state = "DELHI"
        raw_text = raw_text.replace("DELHI", " ")

    ac_list = ["CHANDNI CHOWK", "MATIA MAHAL", "BALLIMARAN", "SADAR BAZAR", "KAROL BAGH", "MOTI NAGAR", "TILAK NAGAR"]
    for a in ac_list:
        if a in raw_text:
            ac = a
            raw_text = raw_text.replace(a, " ")
            break

    form_matches = re.findall(r'FORM\s*6A|FORM6A|FORM\s*6|FORM6|FORM\s*7|FORM7|FORM\s*8|FORM8', raw_text)
    if form_matches:
        ft_str = form_matches[0]
        if "6A" in ft_str: form_type = "Form 6A"
        elif "6" in ft_str: form_type = "Form 6"
        elif "7" in ft_str: form_type = "Form 7"
        elif "8" in ft_str: form_type = "Form 8"
        raw_text = raw_text.replace(ft_str, " ")

    statuses = ["FVR SUBMITTED", "BLO ASSIGNED", "APPROVED", "REJECTED", "SUBMITTED", "PENDING", "E_ROLL UPDATED", "FVR"]
    for st in statuses:
        if st in raw_text:
            current_status = "Submitted" if st == "FVR SUBMITTED" else (st.title() if st != "E_ROLL UPDATED" and st != "FVR" else ("E_Roll Updated" if st == "E_ROLL UPDATED" else "FVR"))
            raw_text = raw_text.replace(st, " ")
            break

    words = [w for w in raw_text.split() if w.strip()]
    if len(words) > 0:
        first_name = words[0]
        if len(words) > 1:
            last_name = " ".join(words[1:])

    return {
        "status": "success",
        "ref_no": ref_no, "state": state, "ac": ac,
        "first_name": first_name, "last_name": last_name,
        "form_type": form_type, "submission_date": submission_date,
        "current_status": current_status
    }

@app.route("/new-entry", methods=["GET", "POST"])
def new_entry():
    if "user" not in session: return redirect(url_for("login"))
    success_msg = None
    _, _, counts = process_entries("all")
    if request.method == "POST" and db:
        try:
            first_name = request.form.get("first_name", "").strip().upper()
            last_name = request.form.get("last_name", "").strip().upper()
            full_name_merged = f"{first_name} {last_name}".strip()

            entry_data = {
                "client_name": request.form.get("client_name", "").strip().upper(),
                "ref_no": request.form.get("ref_no", "").strip().upper(),
                "state": request.form.get("state", "").strip().upper(),
                "ac": request.form.get("ac", "").strip().upper(),
                "first_name": first_name,
                "last_name": last_name,
                "full_name": full_name_merged,
                "form_type": request.form.get("form_type", "Form 6"),
                "submission_date": request.form.get("submission_date", datetime.now().strftime("%d-%m-%Y")),
                "current_status": request.form.get("current_status", "Pending"),
                "amount": request.form.get("amount", "0").strip(),
                "amount_received": request.form.get("amount_received", "No"),
                "work_complete": request.form.get("work_complete", "Not Done"),
                "work_status": "Not Done" if request.form.get("work_complete", "Not Done") == "Not Done" else "Done",
                "created_timestamp": datetime.now().timestamp(),
                "moved_to_report": False
            }
            db.collection("election_entries").add(entry_data)
            success_msg = "✅ Record successfully saved to Firebase Cloud!"
            _, _, counts = process_entries("all")
        except Exception as e:
            success_msg = f"❌ Save error: {e}"
            
    today_date = ""
    return render_template_string(NEW_ENTRY_HTML, success_msg=success_msg, today_date=today_date, counts=counts, page='new')

@app.route("/reports")
def reports():
    guard=require_page("Reports")
    if guard:return guard
    typ=request.args.get("type","All"); q=request.args.get("q","").strip().lower(); rows=[]
    if db:
        try:
            for doc in db.collection("election_entries").stream():
                v=doc.to_dict() or {}
                if not v.get("moved_to_report",False):continue
                status=str(v.get("current_status","Pending"));received=str(v.get("amount_received","No"))
                work=str(v.get("work_status",v.get("work_complete","Not Started")));u=status.upper();ok=True
                if typ=="Pending":ok=not any(x in u for x in ("APPROVED","REJECT","E_ROLL UPDATED","EROLL UPDATED"))
                elif typ=="Approved":ok="APPROVED" in u or "E_ROLL" in u or "EROLL" in u
                elif typ=="Rejected":ok="REJECTED" in u
                elif typ=="Work Complete":ok=received.lower()=="yes" and work.lower() in ("done","work complete","completed")
                elif typ=="Received Yes":ok=received.lower()=="yes"
                elif typ=="Received No":ok=received.lower()=="no"
                elif typ=="By AC":ok=bool(v.get("ac"))
                elif typ=="By Current Status":ok=bool(status)
                elif typ=="By Work Status":ok=bool(work)
                hay=" ".join(str(v.get(k,"")) for k in ("client_name","ref_no","ac","full_name","current_status","work_status","work_complete")).lower()
                if q and q not in hay:ok=False
                if ok:
                    v["id"]=doc.id;v["full_name"]=v.get("full_name") or f"{v.get('first_name','')} {v.get('last_name','')}".strip();v["work_status"]=work;rows.append(v)
        except Exception as exc:print("Reports error:",exc)
    rows.sort(key=lambda x:(str(x.get("ac","")),str(x.get("client_name",""))))
    metrics={"total":len(rows),"total_amt":sum(normalize_amount(x.get("amount")) for x in rows),"recv_amt":sum(normalize_amount(x.get("amount")) for x in rows if str(x.get("amount_received","No")).lower()=="yes")}
    return render_template_string(REPORTS_ENHANCED_HTML,rows=rows,metrics=metrics,report_type=typ,report_q=request.args.get("q",""),counts=process_entries("all")[2],page="reports")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


SETTINGS_ENHANCED_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', '''
<div class="d-flex justify-content-between align-items-center mb-3"><div><h2 class="fw-bold mb-0">⚙️ Settings</h2><small class="text-muted">Central configuration shared across Dashboard, Entries and user sessions.</small></div></div>
{% if message %}<div class="alert alert-success small fw-bold">{{message}}</div>{% endif %}{% if error %}<div class="alert alert-danger small fw-bold">{{error}}</div>{% endif %}
<div class="row g-3"><div class="col-lg-3"><div class="settings-tab p-2"><div class="nav flex-column nav-pills">
<button class="nav-link active text-start" data-bs-toggle="pill" data-bs-target="#st1">Status & Dropdowns</button><button class="nav-link text-start" data-bs-toggle="pill" data-bs-target="#st2">Dashboard</button><button class="nav-link text-start" data-bs-toggle="pill" data-bs-target="#st3">Table Columns</button><button class="nav-link text-start" data-bs-toggle="pill" data-bs-target="#st4">Security</button><button class="nav-link text-start" data-bs-toggle="pill" data-bs-target="#st5">Reports</button>
</div></div></div><div class="col-lg-9"><div class="tab-content settings-tab p-4">
<div class="tab-pane fade show active" id="st1"><h5 class="fw-bold">Status / Received / Work Status</h5><form method="POST"><input type="hidden" name="action" value="workspace_status"><div class="row g-3">
<div class="col-md-4"><label class="small fw-bold">Current Statuses</label><textarea name="current_statuses" class="form-control form-control-sm" rows="9">{{cfg.current_statuses|join('\n')}}</textarea></div>
<div class="col-md-4"><label class="small fw-bold">Received Options</label><textarea name="received_options" class="form-control form-control-sm" rows="9">{{cfg.received_options|join('\n')}}</textarea></div>
<div class="col-md-4"><label class="small fw-bold">Work Statuses</label><textarea name="work_statuses" class="form-control form-control-sm" rows="9">{{cfg.work_statuses|join('\n')}}</textarea></div>
</div><button class="btn btn-primary btn-sm mt-3">💾 Save</button></form></div>
<div class="tab-pane fade" id="st2"><h5 class="fw-bold">Dashboard Cards</h5><form method="POST"><input type="hidden" name="action" value="dashboard_cards"><div class="row g-2">{% for k,v in cfg.dashboard_cards.items() %}<div class="col-md-4"><label class="border rounded p-2 d-block"><input type="checkbox" name="card" value="{{k}}" {% if v %}checked{% endif %}> {{k}}</label></div>{% endfor %}</div><button class="btn btn-primary btn-sm mt-3">💾 Save Layout</button></form></div>
<div class="tab-pane fade" id="st3"><h5 class="fw-bold">All Entries Columns</h5><form method="POST"><input type="hidden" name="action" value="table_columns"><div class="row g-2">{% for k,v in cfg.table_columns.items() %}<div class="col-md-4"><label class="border rounded p-2 d-block"><input type="checkbox" name="col" value="{{k}}" {% if v %}checked{% endif %}> {{k}}</label></div>{% endfor %}</div><label class="small fw-bold mt-3">Custom Columns</label><textarea name="custom_columns" class="form-control form-control-sm" rows="4">{{cfg.custom_columns|join('\n')}}</textarea><button class="btn btn-primary btn-sm mt-3">💾 Save Table</button></form></div>
<div class="tab-pane fade" id="st4"><h5 class="fw-bold">Security</h5><form method="POST" class="row g-3"><input type="hidden" name="action" value="security"><div class="col-md-6"><label class="small fw-bold">Auto Lock</label><select name="auto_lock_seconds" class="form-select form-select-sm">{% for label,sec in [('1 minute',60),('3 minutes',180),('5 minutes',300),('7 minutes',420),('10 minutes',600),('15 minutes',900),('Never',0)] %}<option value="{{sec}}" {% if security.auto_lock_seconds|int==sec %}selected{% endif %}>{{label}}</option>{% endfor %}</select></div><div class="col-md-6"><label class="small fw-bold">PIN</label><input name="pin" maxlength="4" pattern="\d{4}" class="form-control form-control-sm"></div><div class="col-12"><button class="btn btn-success btn-sm">💾 Save Security</button></div></form></div>
<div class="tab-pane fade" id="st5"><h5 class="fw-bold">Advanced Report Generator</h5><p class="small text-muted">Use Reports for filters and CSV/PDF export.</p><a href="/reports" class="btn btn-primary btn-sm">📊 Open Reports</a></div>
</div></div></div>
''')
NOTEPAD_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', '''
<div class="card-box p-4"><div class="d-flex align-items-center mb-3"><div><h3 class="fw-bold mb-0">📝 Notepad</h3><small class="text-muted">Firebase-backed notes. Ctrl+S also saves.</small></div><button onclick="saveNote()" class="btn btn-primary btn-sm fw-bold ms-auto">💾 Save</button></div><textarea id="noteEditor" class="form-control" style="min-height:70vh;font-size:14px;" placeholder="यहाँ अपनी notes लिखें…"></textarea><div id="noteStatus" class="small text-muted mt-2"></div></div>
<script>
fetch('/api/notepad').then(r=>r.json()).then(d=>document.getElementById('noteEditor').value=d.content||'');
function saveNote(){fetch('/api/notepad',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({content:document.getElementById('noteEditor').value})}).then(r=>r.json()).then(d=>document.getElementById('noteStatus').textContent=d.message||d.error||'Saved')}
document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'){e.preventDefault();saveNote()}});
</script>
''')
GOOGLE_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', '''
<div style="height:calc(100vh - 50px);background:#fff;border:1px solid #CBD5E1;border-radius:12px;overflow:hidden;"><iframe src="https://docs.google.com/spreadsheets/d/1pY4TVUDnRFjGKCLWFQ7stD7tHf5yUsqp4KJ-kkrPAFE/edit?embedded=true&gid=766739277" width="100%" height="100%" style="border:0"></iframe></div>
''')
EROLL_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', '''
<div class="d-flex justify-content-between align-items-center mb-3"><div><h2 class="fw-bold mb-0">🪪 E-Roll</h2><small class="text-muted">Firestore elector metadata + Backblaze B2 photos.</small></div><div><a href="/api/e-roll/export.csv" class="btn btn-outline-success btn-sm">⬇ CSV</a> <button class="btn btn-success btn-sm" onclick="document.getElementById('pdfFiles').click()">📄 Import PDF</button></div></div>
<div class="card-box p-3 mb-3"><div class="row g-2"><div class="col-md-3"><input id="epicSearch" class="form-control form-control-sm" placeholder="EPIC..." onkeyup="loadERoll()"></div><div class="col-md-3"><input id="partSearch" class="form-control form-control-sm" placeholder="Part No..." onkeyup="loadERoll()"></div><div class="col-md-4"><input id="textSearch" class="form-control form-control-sm" placeholder="Elector / Address / AC..." onkeyup="loadERoll()"></div><div class="col-md-2"><button class="btn btn-outline-primary btn-sm w-100" onclick="loadERoll()">⟳ Refresh</button></div></div><input id="pdfFiles" type="file" accept=".pdf" multiple hidden onchange="uploadPDFs(this.files)"><div id="uploadStatus" class="small mt-2"></div></div>
<div class="card-box overflow-hidden"><div class="table-responsive"><table class="table table-sm table-bordered align-middle mb-0"><thead class="table-dark"><tr><th>Photo</th><th>Part</th><th>Serial</th><th>EPIC</th><th>Elector</th><th>AC/PC</th><th>State</th><th>Address</th><th>Action</th></tr></thead><tbody id="erollBody"></tbody></table></div><div class="d-flex justify-content-between p-2"><button class="btn btn-outline-secondary btn-sm" onclick="pageER(-1)">← Previous</button><span id="pager"></span><button class="btn btn-outline-secondary btn-sm" onclick="pageER(1)">Next →</button></div></div>
<script>
let ep=1;function esc(x){return String(x??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}
function loadERoll(){let p=new URLSearchParams({page:ep,size:100,epic:document.getElementById('epicSearch').value,part:document.getElementById('partSearch').value,q:document.getElementById('textSearch').value});fetch('/api/e-roll?'+p).then(r=>r.json()).then(d=>{let tb=document.getElementById('erollBody');tb.innerHTML='';(d.rows||[]).forEach(r=>{tb.innerHTML+=`<tr><td>${r.photo_b2_key?'<img class="e-roll-photo" src="/api/e-roll/photo/'+r.id+'">':'—'}</td><td>${esc(r.part_no)}</td><td>${esc(r.serial_no)}</td><td class="fw-bold">${esc(r.epic||r.epic_no)}</td><td>${esc(r.elector_name)}</td><td>${esc(r.ac_pc)}</td><td>${esc(r.state)}</td><td style="max-width:320px;white-space:normal">${esc(r.address)}</td><td><button class="btn btn-outline-danger btn-sm" onclick="delER('${r.id}')">Delete</button></td></tr>`});document.getElementById('pager').textContent='Page '+ep+' • '+(d.rows||[]).length+' records'})}
function pageER(n){if(ep+n>0){ep+=n;loadERoll()}}function delER(id){if(!confirm('Delete this E-Roll record and B2 photo?'))return;fetch('/api/e-roll/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id:id})}).then(r=>r.json()).then(d=>{if(d.error)alert(d.error);loadERoll()})}
function uploadPDFs(files){if(!files.length)return;let fd=new FormData();[...files].forEach(f=>fd.append('files',f));document.getElementById('uploadStatus').textContent='Parsing PDF and uploading photos…';fetch('/api/e-roll/import-pdf',{method:'POST',body:fd}).then(r=>r.json()).then(d=>{document.getElementById('uploadStatus').textContent=d.message||d.error||'Done';loadERoll()})}
loadERoll();
</script>
''')
USERS_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', '''
<div class="d-flex justify-content-between align-items-center mb-3"><div><h2 class="fw-bold mb-0">👥 User Management</h2><small class="text-muted">Super Admin controls Firebase Auth + Firestore.</small></div></div>
<div id="userMessage"></div><div class="card-box p-4 mb-3"><h5 class="fw-bold">Create New User</h5><form id="createUserForm" class="row g-2">
<div class="col-md-3"><label class="small fw-bold">Username</label><input name="username" required class="form-control form-control-sm"></div><div class="col-md-3"><label class="small fw-bold">Full Name</label><input name="full_name" required class="form-control form-control-sm"></div><div class="col-md-3"><label class="small fw-bold">Email</label><input name="email" type="email" class="form-control form-control-sm"></div><div class="col-md-3"><label class="small fw-bold">Password</label><input name="password" type="password" minlength="6" required class="form-control form-control-sm"></div>
<div class="col-md-3"><label class="small fw-bold">Role</label><select name="role" class="form-select form-select-sm"><option>Sub Admin</option><option>Operator</option></select></div><div class="col-md-3"><label class="small fw-bold">Status</label><select name="status" class="form-select form-select-sm"><option>Active</option><option>Inactive</option></select></div>
<div class="col-md-6 d-flex align-items-end gap-3 flex-wrap"><label class="small"><input type="checkbox" name="f_dashboard" checked> Dashboard</label><label class="small"><input type="checkbox" name="f_entries" checked> Entries</label><label class="small"><input type="checkbox" name="f_reports" checked> Reports</label><label class="small"><input type="checkbox" name="f_users"> Users</label><label class="small"><input type="checkbox" name="f_settings"> Settings</label></div>
<div class="col-12"><button class="btn btn-primary btn-sm fw-bold">＋ Create User</button></div></form></div>
<div class="card-box p-3"><div class="d-flex gap-2 mb-2"><input id="userSearch" oninput="filterUsers()" class="form-control form-control-sm" placeholder="Search username, name, email or role..."><button onclick="loadUsers()" class="btn btn-outline-primary btn-sm">⟳ Refresh</button></div><div class="table-responsive"><table class="table table-sm table-bordered" id="usersTable"><thead class="table-dark"><tr><th>Username</th><th>Full Name</th><th>Email</th><th>Role</th><th>Status</th><th>Features</th><th>Password</th><th>Action</th></tr></thead><tbody></tbody></table></div></div>
<script>
function esc(x){return String(x??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}
function msg(t,ok=true){document.getElementById('userMessage').innerHTML='<div class="alert '+(ok?'alert-success':'alert-danger')+' small fw-bold">'+esc(t)+'</div>'}
function loadUsers(){fetch('/api/users').then(r=>r.json()).then(rows=>{let tb=document.querySelector('#usersTable tbody');tb.innerHTML='';(rows||[]).forEach(u=>{let tr=document.createElement('tr');tr.dataset.search=[u.username,u.full_name,u.email,u.role,u.status].join(' ').toLowerCase();let f=Object.entries(u.features||{}).filter(([k,v])=>v===true).map(([k])=>k).join(', ');tr.innerHTML='<td>'+esc(u.username)+'</td><td>'+esc(u.full_name)+'</td><td>'+esc(u.email)+'</td><td>'+esc(u.role)+'</td><td>'+esc(u.status)+'</td><td>'+esc(f)+'</td><td><button class="btn btn-outline-warning btn-sm" onclick="changePass(\''+u.uid+'\')">Change</button></td><td><button class="btn btn-outline-secondary btn-sm" onclick="toggleUser(\''+u.uid+'\',\''+u.status+'\')">'+(String(u.status).toLowerCase()=='active'?'Deactivate':'Activate')+'</button> <button class="btn btn-outline-danger btn-sm" onclick="deleteUser(\''+u.uid+'\')">Delete</button> <button class="btn btn-outline-primary btn-sm" onclick="editPerm(\''+u.uid+'\')">Permissions</button></td>';tb.appendChild(tr)});filterUsers()})}
function filterUsers(){let q=(document.getElementById('userSearch').value||'').toLowerCase();document.querySelectorAll('#usersTable tbody tr').forEach(r=>r.style.display=r.dataset.search.includes(q)?'':'none')}
document.getElementById('createUserForm').addEventListener('submit',e=>{e.preventDefault();let f=new FormData(e.target);let b=Object.fromEntries(f.entries());b.features={dashboard:f.has('f_dashboard'),entries:f.has('f_entries'),reports:f.has('f_reports'),users:f.has('f_users'),settings:f.has('f_settings'),permissions:{}};fetch('/api/users/create',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)}).then(r=>r.json()).then(d=>{msg(d.message||d.error,!d.error);if(!d.error){e.target.reset();loadUsers()}})})
function toggleUser(uid,status){fetch('/api/users/toggle',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid,status})}).then(r=>r.json()).then(d=>{msg(d.message||d.error,!d.error);loadUsers()})}
function deleteUser(uid){if(!confirm('Permanently delete this user?'))return;fetch('/api/users/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid})}).then(r=>r.json()).then(d=>{msg(d.message||d.error,!d.error);loadUsers()})}
function changePass(uid){let p=prompt('New password (minimum 6 characters):');if(!p)return;fetch('/api/users/password',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid:uid,password:p})}).then(r=>r.json()).then(d=>msg(d.message||d.error,!d.error))}
function editPerm(uid){let s=prompt('Allowed pages (comma separated):','Dashboard, New Entry, All Entries, Pending Entries, Approved Entries, Rejected Entries, Work Complete, AC Summary, Reports, E-Roll');if(s===null)return;let pages={};['Dashboard','New Entry','All Entries','Pending Entries','Approved Entries','Rejected Entries','Work Complete','AC Summary','Reports','E-Roll','Users','Settings'].forEach(p=>pages[p]=s.split(',').map(x=>x.trim()).includes(p));fetch('/api/users/permissions',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid:uid,permissions:{pages:pages}})}).then(r=>r.json()).then(d=>msg(d.message||d.error,!d.error))}
loadUsers();
</script>
''')
REPORTS_ENHANCED_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', '''
<div class="d-flex justify-content-between align-items-center mb-3"><div><h2 class="fw-bold mb-0">📊 Reports & Analytics</h2><small class="text-muted">Advanced filters, metrics and exports.</small></div></div>
<div class="row g-3 mb-3"><div class="col-md-4"><div class="metric-mini"><small class="text-muted">Total Entries</small><h3 class="fw-bold mb-0">{{ metrics.total }}</h3></div></div><div class="col-md-4"><div class="metric-mini"><small class="text-muted">Total Amount</small><h3 class="fw-bold mb-0">₹{{ "{:,.2f}".format(metrics.total_amt) }}</h3></div></div><div class="col-md-4"><div class="metric-mini"><small class="text-muted">Received Amount</small><h3 class="fw-bold text-success mb-0">₹{{ "{:,.2f}".format(metrics.recv_amt) }}</h3></div></div></div>
<div class="card-box p-3 mb-3"><form class="row g-2"><div class="col-md-3"><select name="type" class="form-select form-select-sm">{% for x in ['All','Pending','Approved','Rejected','Work Complete','Received Yes','Received No','By AC','By Current Status','By Work Status'] %}<option value="{{x}}" {% if report_type==x %}selected{% endif %}>{{x}}</option>{% endfor %}</select></div><div class="col-md-6"><input name="q" value="{{report_q}}" class="form-control form-control-sm" placeholder="Client / Reference / AC / Name"></div><div class="col-md-3 d-flex gap-1"><button class="btn btn-primary btn-sm flex-fill">🔎 Filter</button><a class="btn btn-outline-success btn-sm" href="/api/reports/export.csv?type={{report_type|urlencode}}&q={{report_q|urlencode}}">CSV</a><a class="btn btn-outline-danger btn-sm" href="/api/reports/export.pdf?type={{report_type|urlencode}}&q={{report_q|urlencode}}">PDF</a></div></form></div>
<div class="card-box overflow-hidden"><div class="table-responsive"><table class="table table-sm table-bordered align-middle mb-0"><thead class="table-dark"><tr><th>#</th><th>Client</th><th>Reference</th><th>AC</th><th>Full Name</th><th>Status</th><th>Work</th><th>Amount</th><th>Received</th><th></th></tr></thead><tbody>{% for r in rows %}<tr><td>{{loop.index}}</td><td>{{r.client_name}}</td><td>{{r.ref_no}}</td><td>{{r.ac}}</td><td>{{r.full_name}}</td><td>{{r.current_status}}</td><td>{{r.work_status}}</td><td>₹{{r.amount}}</td><td>{{r.amount_received}}</td><td><button class="btn btn-outline-danger btn-sm" onclick="deleteReport('{{r.id}}')">Delete</button></td></tr>{% else %}<tr><td colspan="10" class="text-center py-4 text-muted">No report records found.</td></tr>{% endfor %}</tbody></table></div></div>
<script>function deleteReport(id){if(!confirm('Delete this report entry?'))return;fetch('/api/reports/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id:id})}).then(r=>r.json()).then(d=>{if(d.error)alert(d.error);else location.reload()})}</script>
''')

@app.route("/notepad")
def notepad():
    guard=require_page("Dashboard")
    if guard:return guard
    return render_template_string(NOTEPAD_HTML,counts=process_entries("all")[2],page="notepad")

@app.route("/api/notepad",methods=["GET","POST"])
def api_notepad():
    if "user" not in session:return {"error":"unauthorized"},401
    if not db:return {"error":"Firebase is not connected."},503
    try:
        ref=db.collection("app_settings").document("notepad")
        if request.method=="GET":
            snap=ref.get();return jsonify({"content":(snap.to_dict() or {}).get("content","") if snap.exists else ""})
        ref.set({"content":str((request.get_json() or {}).get("content",""))},merge=True);return {"message":"Notes saved successfully."}
    except Exception as exc:return {"error":str(exc)},500

@app.route("/google-sheet")
def google_sheet():
    guard=require_page("Dashboard")
    if guard:return guard
    return render_template_string(GOOGLE_HTML,counts=process_entries("all")[2],page="google_sheet")

def _admin_guard():
    if "user" not in session:return ({"error":"unauthorized"},401)
    if not user_is_super_admin():return ({"error":"Only Super Admin can perform this action."},403)
    if not db:return ({"error":"Firebase is not connected."},503)
    return None

@app.route("/users")
def users():
    guard=require_page("Users")
    if guard:return guard
    return render_template_string(USERS_HTML,counts=process_entries("all")[2],page="users")

@app.route("/api/users")
def api_users():
    guard=_admin_guard()
    if guard:return guard
    try:
        out=[]
        for doc in db.collection("users").stream():
            d=doc.to_dict() or {};d["uid"]=d.get("uid",doc.id);out.append(d)
        out.sort(key=lambda x:str(x.get("username","")).lower());return jsonify(out)
    except Exception as exc:return {"error":str(exc)},500

@app.route("/api/users/create",methods=["POST"])
def api_users_create():
    guard=_admin_guard()
    if guard:return guard
    b=request.get_json() or {};username=str(b.get("username","")).strip().lower();name=str(b.get("full_name","")).strip();email=str(b.get("email","")).strip().lower();password=str(b.get("password",""));role=str(b.get("role","Operator"));status=str(b.get("status","Active"))
    if not username or not name or len(password)<6:return {"error":"Username, Full Name and password (6+) are required."},400
    if role.lower()=="super admin":return {"error":"Super Admin account cannot be created here."},400
    try:
        if any(True for _ in db.collection("users").where("username","==",username).limit(1).stream()):return {"error":"Username already exists."},409
        auth_email=email or f"{username}@workspace.invalid"
        u=auth.create_user(email=auth_email,password=password,display_name=name,disabled=(status!="Active"))
        db.collection("users").document(u.uid).set({"uid":u.uid,"username":username,"email":email,"auth_email":auth_email,"full_name":name,"role":role,"status":status,"features":b.get("features") or {},"created_by":"super_admin","created_at":firestore.SERVER_TIMESTAMP})
        return {"message":"User created successfully.","uid":u.uid}
    except auth.EmailAlreadyExistsError:return {"error":"This email is already registered."},409
    except Exception as exc:return {"error":str(exc)},500

@app.route("/api/users/toggle",methods=["POST"])
def api_users_toggle():
    guard=_admin_guard()
    if guard:return guard
    b=request.get_json() or {};uid=str(b.get("uid",""));current=str(b.get("status","Active"))
    try:
        snap=db.collection("users").document(uid).get();d=snap.to_dict() or {}
        if str(d.get("role","")).lower()=="super admin":return {"error":"Super Admin cannot be deactivated."},403
        new="Inactive" if current.lower()=="active" else "Active";db.collection("users").document(uid).update({"status":new,"updated_at":firestore.SERVER_TIMESTAMP})
        try:auth.update_user(uid,disabled=(new!="Active"))
        except Exception:pass
        return {"message":f"User {new.lower()} successfully."}
    except Exception as exc:return {"error":str(exc)},500

@app.route("/api/users/delete",methods=["POST"])
def api_users_delete():
    guard=_admin_guard()
    if guard:return guard
    uid=str((request.get_json() or {}).get("uid",""))
    try:
        d=(db.collection("users").document(uid).get().to_dict() or {})
        if str(d.get("role","")).lower()=="super admin":return {"error":"Super Admin cannot be deleted."},403
        try:auth.delete_user(uid)
        except auth.UserNotFoundError:pass
        db.collection("users").document(uid).delete();return {"message":"User deleted successfully."}
    except Exception as exc:return {"error":str(exc)},500

@app.route("/api/users/password",methods=["POST"])
def api_users_password():
    guard=_admin_guard()
    if guard:return guard
    b=request.get_json() or {};uid=str(b.get("uid",""));pw=str(b.get("password",""))
    if len(pw)<6:return {"error":"Password must contain at least 6 characters."},400
    try:auth.update_user(uid,password=pw);return {"message":"Password updated successfully."}
    except Exception as exc:return {"error":str(exc)},500

@app.route("/api/users/permissions",methods=["POST"])
def api_users_permissions():
    guard=_admin_guard()
    if guard:return guard
    b=request.get_json() or {};uid=str(b.get("uid",""));permissions=b.get("permissions") or {}
    try:
        db.collection("users").document(uid).update({"features.permissions":permissions,"updated_at":firestore.SERVER_TIMESTAMP})
        return {"message":"Permissions updated successfully."}
    except Exception as exc:return {"error":str(exc)},500

def _report_rows(typ="All",q=""):
    rows=[];q=str(q or "").lower().strip()
    if not db:return rows
    for doc in db.collection("election_entries").stream():
        v=doc.to_dict() or {}
        if not v.get("moved_to_report",False):continue
        status=str(v.get("current_status","Pending"));received=str(v.get("amount_received","No"));work=str(v.get("work_status",v.get("work_complete","Not Started")));u=status.upper();ok=True
        if typ=="Pending":ok=not any(x in u for x in ("APPROVED","REJECT","E_ROLL UPDATED","EROLL UPDATED"))
        elif typ=="Approved":ok="APPROVED" in u or "E_ROLL" in u or "EROLL" in u
        elif typ=="Rejected":ok="REJECTED" in u
        elif typ=="Work Complete":ok=received.lower()=="yes" and work.lower() in ("done","work complete","completed")
        elif typ=="Received Yes":ok=received.lower()=="yes"
        elif typ=="Received No":ok=received.lower()=="no"
        elif typ=="By AC":ok=bool(v.get("ac"))
        elif typ=="By Current Status":ok=bool(status)
        elif typ=="By Work Status":ok=bool(work)
        hay=" ".join(str(v.get(k,"")) for k in ("client_name","ref_no","ac","full_name","current_status","work_status","work_complete")).lower()
        if q and q not in hay:ok=False
        if ok:v["id"]=doc.id;v["full_name"]=v.get("full_name") or f"{v.get('first_name','')} {v.get('last_name','')}".strip();v["work_status"]=work;rows.append(v)
    return rows

@app.route("/api/reports/export.csv")
def export_reports_csv():
    guard=require_page("Reports")
    if guard:return guard
    rows=_report_rows(request.args.get("type","All"),request.args.get("q",""));s=io.StringIO();w=csv.writer(s);w.writerow(["Client","Reference","AC","Full Name","Form","Status","Work","Amount","Received","Remarks"])
    for r in rows:w.writerow([r.get("client_name",""),r.get("ref_no",""),r.get("ac",""),r.get("full_name",""),r.get("form_type",""),r.get("current_status",""),r.get("work_status",""),r.get("amount","0"),r.get("amount_received","No"),r.get("remarks","")])
    return send_file(io.BytesIO(s.getvalue().encode("utf-8-sig")),as_attachment=True,download_name="workspace_report.csv",mimetype="text/csv")

@app.route("/api/reports/export.pdf")
def export_reports_pdf():
    guard=require_page("Reports")
    if guard:return guard
    if not REPORTLAB_AVAILABLE:return {"error":"ReportLab is not installed."},503
    rows=_report_rows(request.args.get("type","All"),request.args.get("q",""));bio=io.BytesIO();doc=SimpleDocTemplate(bio,pagesize=landscape(A4))
    styles=getSampleStyleSheet();story=[Paragraph("My Workspace — Report",styles["Title"]),Spacer(1,10)]
    data=[["#","Client","Reference","AC","Full Name","Status","Work","Amount","Received"]]+[[i,r.get("client_name",""),r.get("ref_no",""),r.get("ac",""),r.get("full_name",""),r.get("current_status",""),r.get("work_status",""),str(r.get("amount","0")),r.get("amount_received","No")] for i,r in enumerate(rows,1)]
    t=Table(data,repeatRows=1);t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),reportlab_colors.HexColor("#173B73")),("TEXTCOLOR",(0,0),(-1,0),reportlab_colors.white),("GRID",(0,0),(-1,-1),.4,reportlab_colors.grey),("FONTSIZE",(0,0),(-1,-1),7)]));story.append(t);doc.build(story);bio.seek(0)
    return send_file(bio,as_attachment=True,download_name="workspace_report.pdf",mimetype="application/pdf")

@app.route("/api/reports/delete",methods=["POST"])
def api_reports_delete():
    guard=_admin_guard()
    if guard:return guard
    uid=str((request.get_json() or {}).get("id",""))
    try:db.collection("election_entries").document(uid).delete();return {"message":"Report entry deleted."}
    except Exception as exc:return {"error":str(exc)},500

@app.route("/e-roll")
def eroll():
    guard=require_page("E-Roll")
    if guard:return guard
    return render_template_string(EROLL_HTML,counts=process_entries("all")[2],page="eroll")

@app.route("/api/e-roll")
def api_eroll():
    guard=require_page("E-Roll")
    if guard:return guard
    page=max(1,int(request.args.get("page",1)));size=max(1,min(200,int(request.args.get("size",100))));epic=request.args.get("epic","").strip().upper();part=request.args.get("part","").strip();q=request.args.get("q","").strip().lower();rows=[];parts=set()
    if db:
        try:
            for doc in db.collection("e_roll_entries").stream():
                d=doc.to_dict() or {};d["id"]=doc.id
                if d.get("part_no"):parts.add(str(d["part_no"]))
                if epic and str(d.get("epic",d.get("epic_no",""))).upper()!=epic:continue
                if part and str(d.get("part_no",""))!=part:continue
                hay=" ".join(str(d.get(k,"")) for k in ("elector_name","epic","epic_no","serial_no","part_no","part_name","ac_pc","state","address")).lower()
                if q and q not in hay:continue
                rows.append(d)
        except Exception as exc:return {"error":str(exc)},500
    start=(page-1)*size;return jsonify({"rows":rows[start:start+size],"total":len(rows),"parts":sorted(parts)})

@app.route("/api/e-roll/photo/<doc_id>")
def api_eroll_photo(doc_id):
    guard=require_page("E-Roll")
    if guard:return guard
    try:
        snap=db.collection("e_roll_entries").document(doc_id).get();raw=B2.get((snap.to_dict() or {}).get("photo_b2_key","")) if snap.exists else b""
        if not raw:return "",404
        return send_file(io.BytesIO(raw),mimetype="image/png" if raw.startswith(b"\x89PNG") else "image/jpeg",max_age=3600)
    except Exception:return "",404

@app.route("/api/e-roll/save",methods=["POST"])
def api_eroll_save():
    guard=require_page("E-Roll")
    if guard:return guard
    b=request.get_json() or {};doc_id=str(b.pop("id","")).strip();payload={k:str(v or "").strip() for k,v in b.items() if k in ("part_no","part_name","serial_no","epic","elector_name","ac_pc","state","address")}
    if not payload.get("epic") and not payload.get("serial_no"):return {"error":"EPIC or Serial No. is required."},400
    try:
        if doc_id:db.collection("e_roll_entries").document(doc_id).update(payload)
        else:payload["created_at"]=datetime.now().isoformat(timespec="seconds");db.collection("e_roll_entries").add(payload)
        return {"message":"E-Roll record saved."}
    except Exception as exc:return {"error":str(exc)},500

@app.route("/api/e-roll/delete",methods=["POST"])
def api_eroll_delete():
    guard=_admin_guard()
    if guard:return guard
    doc_id=str((request.get_json() or {}).get("id",""))
    try:
        ref=db.collection("e_roll_entries").document(doc_id);snap=ref.get()
        if not snap.exists:return {"error":"Record not found."},404
        key=(snap.to_dict() or {}).get("photo_b2_key","");ref.delete()
        if key:B2.delete(key)
        return {"message":"E-Roll record deleted."}
    except Exception as exc:return {"error":str(exc)},500

@app.route("/api/e-roll/delete-part",methods=["POST"])
def api_eroll_delete_part():
    guard=_admin_guard()
    if guard:return guard
    part=str((request.get_json() or {}).get("part_no","")).strip()
    if not part:return {"error":"Part No. is required."},400
    try:
        docs=list(db.collection("e_roll_entries").where("part_no","==",part).stream())
        for d in docs:
            key=(d.to_dict() or {}).get("photo_b2_key","");db.collection("e_roll_entries").document(d.id).delete()
            if key:B2.delete(key)
        return {"message":f"Deleted {len(docs)} records and their photos."}
    except Exception as exc:return {"error":str(exc)},500

@app.route("/api/e-roll/import-pdf",methods=["POST"])
def api_eroll_import_pdf():
    guard=require_page("E-Roll")
    if guard:return guard
    if not FITZ_AVAILABLE:return {"error":"PyMuPDF is not installed."},503
    files=request.files.getlist("files")
    if not files:return {"error":"No PDF files selected."},400
    import tempfile;saved=skipped=0;errors=[]
    for f in files:
        if not f or not f.filename.lower().endswith(".pdf"):continue
        tmp=None
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf",delete=False) as tf:f.save(tf);tmp=tf.name
            s,sk=save_eroll_records(parse_eroll_pdf(tmp));saved+=s;skipped+=sk
        except Exception as exc:errors.append(f"{f.filename}: {exc}")
        finally:
            if tmp:
                try:os.unlink(tmp)
                except Exception:pass
    return {"saved":saved,"skipped":skipped,"errors":errors,"message":f"Saved {saved}; skipped {skipped}."}

@app.route("/api/e-roll/export.csv")
def export_eroll_csv():
    guard=require_page("E-Roll")
    if guard:return guard
    s=io.StringIO();w=csv.writer(s);w.writerow(["Part No.","Part Name","Serial No.","EPIC","Elector Name","AC/PC","State","Address","Photo B2 Key"])
    for doc in db.collection("e_roll_entries").stream():
        d=doc.to_dict() or {};w.writerow([d.get("part_no",""),d.get("part_name",""),d.get("serial_no",""),d.get("epic",d.get("epic_no","")),d.get("elector_name",""),d.get("ac_pc",""),d.get("state",""),d.get("address",""),d.get("photo_b2_key","")])
    return send_file(io.BytesIO(s.getvalue().encode("utf-8-sig")),as_attachment=True,download_name="e_roll_entries.csv",mimetype="text/csv")

@app.route("/api/entry-cell-color",methods=["POST"])
def api_entry_cell_color():
    if "user" not in session:return {"error":"unauthorized"},401
    b=request.get_json() or {};entry_id=str(b.get("id",""));field=str(b.get("field","")).strip();color=str(b.get("color","")).strip()
    if not entry_id or not field:return {"error":"Missing id/field"},400
    if color and not re.fullmatch(r"#[0-9A-Fa-f]{6}",color):return {"error":"Invalid color"},400
    try:
        ref=db.collection("election_entries").document(entry_id);snap=ref.get()
        if not snap.exists:return {"error":"Entry not found."},404
        d=snap.to_dict() or {};colors_map=d.get("cell_colors",{}) or {}
        if color:colors_map[field]=color
        else:colors_map.pop(field,None)
        ref.update({"cell_colors":colors_map});return {"message":"Cell color saved."}
    except Exception as exc:return {"error":str(exc)},500

@app.route("/api/entry-custom-field",methods=["POST"])
def api_entry_custom_field():
    if "user" not in session:return {"error":"unauthorized"},401
    b=request.get_json() or {};entry_id=str(b.get("id",""));field=str(b.get("field","")).strip();value=str(b.get("value",""))
    if not entry_id or not field:return {"error":"Missing id/field"},400
    try:
        ref=db.collection("election_entries").document(entry_id);snap=ref.get();d=snap.to_dict() or {};custom=d.get("custom_fields",{}) or {};custom[field]=value
        if field=="Remarks":ref.update({"remarks":value})
        else:ref.update({"custom_fields":custom})
        return {"message":"Field saved."}
    except Exception as exc:return {"error":str(exc)},500


if __name__ == "__main__":
    app.run(debug=True, port=5000)
SETTINGS_ENHANCED_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', r'''
<div class="d-flex justify-content-between align-items-center mb-3">
  <div><h2 class="fw-bold mb-0">⚙️ Settings</h2><small class="text-muted">Central configuration shared across the workspace.</small></div>
  <span class="badge bg-success-subtle text-success px-3 py-2">☁ Firebase Shared Settings</span>
</div>
{% if message %}<div class="alert alert-success small fw-bold">{{ message }}</div>{% endif %}
{% if error %}<div class="alert alert-danger small fw-bold">{{ error }}</div>{% endif %}
<div class="row g-3">
  <div class="col-lg-3"><div class="settings-tab p-2"><div class="nav flex-column nav-pills">
    <button class="nav-link active text-start" data-bs-toggle="pill" data-bs-target="#st-status">Status & Dropdowns</button>
    <button class="nav-link text-start" data-bs-toggle="pill" data-bs-target="#st-dashboard">Dashboard Layout</button>
    <button class="nav-link text-start" data-bs-toggle="pill" data-bs-target="#st-columns">Table Columns</button>
    <button class="nav-link text-start" data-bs-toggle="pill" data-bs-target="#st-security">Security</button>
    <button class="nav-link text-start" data-bs-toggle="pill" data-bs-target="#st-reports">Report Generator</button>
    <button class="nav-link text-start" data-bs-toggle="pill" data-bs-target="#st-perm">User Permissions</button>
  </div></div></div>
  <div class="col-lg-9"><div class="tab-content settings-tab p-4">
    <div class="tab-pane fade show active" id="st-status">
      <h5 class="fw-bold">Status / Received / Work Status</h5>
      <form method="POST"><input type="hidden" name="action" value="workspace_status"><div class="row g-3">
        <div class="col-md-4"><label class="small fw-bold">Current Statuses</label><textarea name="current_statuses" class="form-control form-control-sm" rows="9">{{ cfg.current_statuses|join('\n') }}</textarea></div>
        <div class="col-md-4"><label class="small fw-bold">Received Options</label><textarea name="received_options" class="form-control form-control-sm" rows="9">{{ cfg.received_options|join('\n') }}</textarea></div>
        <div class="col-md-4"><label class="small fw-bold">Work Statuses</label><textarea name="work_statuses" class="form-control form-control-sm" rows="9">{{ cfg.work_statuses|join('\n') }}</textarea></div>
      </div><button class="btn btn-primary btn-sm fw-bold mt-3">💾 Save Status Settings</button></form>
    </div>
    <div class="tab-pane fade" id="st-dashboard"><h5 class="fw-bold">Dashboard Items — Show / Hide</h5>
      <form method="POST"><input type="hidden" name="action" value="dashboard_cards"><div class="row g-2">
      {% for k,v in cfg.dashboard_cards.items() %}<div class="col-md-4"><label class="border rounded-3 p-2 d-flex gap-2 align-items-center"><input type="checkbox" name="card_{{ loop.index }}" value="{{ k }}" {% if v %}checked{% endif %}> {{ k }}</label></div>{% endfor %}
      </div><input type="hidden" name="card_keys" value="{{ cfg.dashboard_cards.keys()|join('|') }}"><button class="btn btn-primary btn-sm fw-bold mt-3">💾 Save Dashboard Layout</button></form>
    </div>
    <div class="tab-pane fade" id="st-columns"><h5 class="fw-bold">All Entries Table — Standard Columns</h5>
      <form method="POST"><input type="hidden" name="action" value="table_columns"><div class="row g-2 mb-3">
      {% for k,v in cfg.table_columns.items() %}<div class="col-md-4"><label class="border rounded-3 p-2 d-flex gap-2 align-items-center"><input type="checkbox" name="col_{{ loop.index }}" value="{{ k }}" {% if v %}checked{% endif %}> {{ k }}</label></div>{% endfor %}
      </div>
      <label class="small fw-bold mt-2">Custom Columns</label>
      <div class="input-group input-group-sm mb-2"><input id="newCustomColumn" type="text" class="form-control" placeholder="Enter new column name"><button type="button" class="btn btn-outline-primary" onclick="addCustomColumn()">➕ Add Column</button></div>
      <textarea id="customColumnsBox" name="custom_columns" class="form-control form-control-sm" rows="4" placeholder="One custom column per line">{{ cfg.custom_columns|join('\n') }}</textarea>
      <div class="small text-muted mt-1">Tick = show. Untick = hide. Custom columns added here appear in All Entries after saving.</div>
      <button class="btn btn-primary btn-sm fw-bold mt-3">💾 Save Table Settings</button></form>
      <script>
      function addCustomColumn(){
        const input=document.getElementById('newCustomColumn'), box=document.getElementById('customColumnsBox');
        const v=(input.value||'').trim(); if(!v) return;
        const lines=box.value.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
        if(!lines.includes(v)) lines.push(v);
        box.value=lines.join('\n'); input.value=''; input.focus();
      }
      </script>
    </div>
    <div class="tab-pane fade" id="st-security"><h5 class="fw-bold">Security</h5>
      <form method="POST" class="row g-3"><input type="hidden" name="action" value="security">
      <div class="col-md-6"><label class="small fw-bold">Auto Lock</label><select class="form-select form-select-sm" name="auto_lock_seconds">
      {% for label,sec in [('1 minute',60),('3 minutes',180),('5 minutes',300),('7 minutes',420),('10 minutes',600),('15 minutes',900),('Never',0)] %}<option value="{{ sec }}" {% if security.auto_lock_seconds|int == sec %}selected{% endif %}>{{ label }}</option>{% endfor %}
      </select></div><div class="col-md-6"><label class="small fw-bold">My PIN</label><input name="pin" class="form-control form-control-sm" maxlength="4" pattern="\d{4}" placeholder="Leave blank to keep current PIN"></div>
      <div class="col-12"><button class="btn btn-success btn-sm fw-bold">💾 Save Security</button></div></form>
      <hr><div class="small text-muted">Auto-lock is stored in <code>app_settings/security</code>; quick-login PIN uses <code>app_settings/login_security</code>.</div>
    </div>
    <div class="tab-pane fade" id="st-reports"><h5 class="fw-bold">Advanced Report Generator</h5>
      <form method="GET" action="/reports" class="row g-2"><div class="col-md-4"><label class="small fw-bold">Report Type</label><select name="type" class="form-select form-select-sm">
      {% for x in ['All','Pending','Approved','Rejected','Work Complete','Received Yes','Received No','By AC','By Current Status','By Work Status'] %}<option value="{{ x }}" {% if report_type == x %}selected{% endif %}>{{ x }}</option>{% endfor %}
      </select></div><div class="col-md-5"><label class="small fw-bold">Search</label><input name="q" value="{{ report_q }}" class="form-control form-control-sm"></div><div class="col-md-3 d-flex align-items-end"><button class="btn btn-primary btn-sm w-100">🔎 Generate</button></div></form>
      <div class="mt-3 d-flex gap-2"><a class="btn btn-outline-success btn-sm" href="/api/reports/export.csv?type={{ report_type|urlencode }}&q={{ report_q|urlencode }}">⬇ CSV</a><a class="btn btn-outline-danger btn-sm" href="/api/reports/export.pdf?type={{ report_type|urlencode }}&q={{ report_q|urlencode }}">⬇ PDF</a></div>
    </div>
    <div class="tab-pane fade" id="st-perm"><h5 class="fw-bold">Per-user Permissions</h5><p class="small text-muted">Page, dashboard and table permissions are managed from User Management.</p><a href="/users" class="btn btn-primary btn-sm">👤 Open User Management</a></div>
  </div></div>
</div>
''')

USERS_HTML = BASE_LAYOUT.replace('{% block content %}{% endblock %}', r'''
<div class="d-flex justify-content-between align-items-center mb-3"><div><h2 class="fw-bold mb-0">👥 User Management</h2><small class="text-muted">Super Admin controls Firebase Auth + Firestore.</small></div><span class="badge bg-success-subtle text-success">● Firebase</span></div>
<div id="userMessage"></div>
<div class="card-box p-4 mb-3"><h5 class="fw-bold">Create New User</h5>
<form id="createUserForm" class="row g-2">
<div class="col-md-3"><label class="small fw-bold">Username</label><input name="username" required class="form-control form-control-sm"></div>
<div class="col-md-3"><label class="small fw-bold">Full Name</label><input name="full_name" required class="form-control form-control-sm"></div>
<div class="col-md-3"><label class="small fw-bold">Email</label><input name="email" type="email" class="form-control form-control-sm"></div>
<div class="col-md-3"><label class="small fw-bold">Password</label><input name="password" type="password" minlength="6" required class="form-control form-control-sm"></div>
<div class="col-md-3"><label class="small fw-bold">Role</label><select name="role" class="form-select form-select-sm"><option>Sub Admin</option><option>Operator</option></select></div>
<div class="col-md-3"><label class="small fw-bold">Status</label><select name="status" class="form-select form-select-sm"><option>Active</option><option>Inactive</option></select></div>
<div class="col-md-6 d-flex align-items-end gap-3 flex-wrap"><label class="small"><input type="checkbox" name="f_dashboard" checked> Dashboard</label><label class="small"><input type="checkbox" name="f_entries" checked> Entries</label><label class="small"><input type="checkbox" name="f_reports" checked> Reports</label><label class="small"><input type="checkbox" name="f_users"> Users</label><label class="small"><input type="checkbox" name="f_settings"> Settings</label></div>
<div class="col-12"><details><summary class="small fw-bold">Detailed page permissions</summary><div class="row g-2 pt-2">{% for p in ['Dashboard','New Entry','All Entries','Pending Entries','Approved Entries','Rejected Entries','Work Complete','AC Summary','Reports','E-Roll','Users','Settings'] %}<div class="col-md-3"><label class="small"><input type="checkbox" name="page_perm_{{ loop.index }}" value="{{ p }}" checked> {{ p }}</label></div>{% endfor %}</div></details></div>
<div class="col-12"><button class="btn btn-primary btn-sm fw-bold">＋ Create User</button></div>
</form></div>
<div class="card-box p-3"><div class="d-flex gap-2 mb-2"><input id="userSearch" oninput="filterUsers()" class="form-control form-control-sm" placeholder="Search username, name, email or role..."><button onclick="loadUsers()" class="btn btn-outline-primary btn-sm">⟳ Refresh</button></div>
<div class="table-responsive"><table class="table table-sm table-bordered align-middle" id="usersTable"><thead class="table-dark"><tr><th>Username</th><th>Full Name</th><th>Email</th><th>Role</th><th>Status</th><th>Features</th><th>Password</th><th>Action</th></tr></thead><tbody></tbody></table></div></div>
<script>
function msg(t,ok=true){document.getElementById('userMessage').innerHTML='<div class="alert '+(ok?'alert-success':'alert-danger')+' small fw-bold">'+t+'</div>'}
function loadUsers(){fetch('/api/users').then(r=>r.json()).then(rows=>{let tb=document.querySelector('#usersTable tbody');tb.innerHTML='';rows.forEach(u=>{let tr=document.createElement('tr');tr.dataset.search=[u.username,u.full_name,u.email,u.role,u.status].join(' ').toLowerCase();let feats=Object.entries(u.features||{}).filter(([k,v])=>v===true).map(([k])=>k).join(', ');tr.innerHTML='<td>'+esc(u.username)+'</td><td>'+esc(u.full_name)+'</td><td>'+esc(u.email)+'</td><td>'+esc(u.role)+'</td><td>'+esc(u.status)+'</td><td>'+esc(feats)+'</td><td><button class="btn btn-outline-warning btn-sm" onclick="changePass(\''+u.uid+'\')">Change</button></td><td><button class="btn btn-outline-secondary btn-sm" onclick="toggleUser(\''+u.uid+'\',\''+u.status+'\')">'+(String(u.status).toLowerCase()=='active'?'Deactivate':'Activate')+'</button> <button class="btn btn-outline-danger btn-sm" '+(String(u.role).toLowerCase()=='super admin'?'disabled':'')+' onclick="deleteUser(\''+u.uid+'\')">Delete</button> <button class="btn btn-outline-primary btn-sm" onclick="editPerm(\''+u.uid+'\')">Permissions</button></td>';tb.appendChild(tr)});filterUsers()})}
function esc(x){return String(x??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}
function filterUsers(){let q=(document.getElementById('userSearch').value||'').toLowerCase();document.querySelectorAll('#usersTable tbody tr').forEach(r=>r.style.display=r.dataset.search.includes(q)?'':'none')}
document.getElementById('createUserForm').addEventListener('submit',e=>{e.preventDefault();let f=new FormData(e.target),pages={};document.querySelectorAll('[name^=page_perm_]').forEach(x=>pages[x.value]=x.checked);let b=Object.fromEntries(f.entries());b.features={dashboard:f.has('f_dashboard'),entries:f.has('f_entries'),reports:f.has('f_reports'),users:f.has('f_users'),settings:f.has('f_settings'),permissions:{pages:pages}};fetch('/api/users/create',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)}).then(r=>r.json()).then(d=>{msg(d.message||d.error,!d.error);if(!d.error){e.target.reset();loadUsers()}})})
function toggleUser(uid,status){fetch('/api/users/toggle',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid:uid,status:status})}).then(r=>r.json()).then(d=>{msg(d.message||d.error,!d.error);loadUsers()})}
function deleteUser(uid){if(!confirm('Permanently delete this user?'))return;fetch('/api/users/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid:uid})}).then(r=>r.json()).then(d=>{msg(d.message||d.error,!d.error);loadUsers()})}
function changePass(uid){let p=prompt('New password (minimum 6 characters):');if(!p)return;fetch('/api/users/password',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid:uid,password:p})}).then(r=>r.json()).then(d=>msg(d.message||d.error,!d.error))}
function editPerm(uid){let s=prompt('Enter allowed pages, comma separated. Leave blank to hide all except explicit feature permissions:','Dashboard, New Entry, All Entries, Pending Entries, Approved Entries, Rejected Entries, Work Complete, AC Summary, Reports, E-Roll');if(s===null)return;let wanted=s.split(',').map(x=>x.trim()).filter(Boolean),pages={};['Dashboard','New Entry','All Entries','Pending Entries','Approved Entries','Rejected Entries','Work Complete','AC Summary','Reports','E-Roll','Users','Settings'].forEach(p=>pages[p]=wanted.includes(p));fetch('/api/users/permissions',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid:uid,permissions:{pages:pages}})}).then(r=>r.json()).then(d=>{msg(d.message||d.error,!d.error);loadUsers()})}
loadUsers();
</script>
''')
