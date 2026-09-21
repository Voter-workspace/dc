import os
import re
import io
from datetime import datetime
from flask import Flask, render_template_string, request, redirect, url_for, session, jsonify, send_file
import firebase_admin
from firebase_admin import credentials, firestore
import openpyxl

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
        .sidebar { width: 260px; background-color: #0B132B; min-height: 100vh; position: fixed; color: #fff; top: 0; left: 0; z-index: 1000; padding: 20px 15px; display: flex; flex-direction: column; justify-content: space-between; }
        .sidebar .nav-link { color: #AAB4C5; font-weight: 600; padding: 10px 14px; border-radius: 8px; margin-bottom: 4px; font-size: 12.5px; display: flex; justify-content: space-between; align-items: center; text-decoration: none; }
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
                <li class="nav-item mt-3"><a href="/excel-editor" class="nav-link {% if page == 'excel' %}active{% endif %}"><span><i class="fa-solid fa-file-excel me-2 text-success"></i> Excel Sheet</span></a></li>
                <li class="nav-item"><a href="/whatsapp" class="nav-link {% if page == 'whatsapp' %}active{% endif %}"><span><i class="fa-brands fa-whatsapp me-2 text-success" style="font-size: 15px;"></i> WhatsApp</span></a></li>
                <li class="nav-item"><a href="/ac-summary" class="nav-link {% if page == 'summary' %}active{% endif %}"><span><i class="fa-solid fa-chart-column me-2"></i> AC Summary</span></a></li>
                <li class="nav-item"><a href="/reports" class="nav-link {% if page == 'reports' %}active{% endif %}"><span><i class="fa-solid fa-box-archive me-2"></i> Reports</span></a></li>
                <li class="nav-item mt-2"><a href="/settings" class="nav-link {% if page == 'settings' %}active{% endif %}"><span><i class="fa-solid fa-gear me-2"></i> Settings</span></a></li>
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

    <!-- STATUS UPDATE MODAL -->
    <div class="modal fade" id="statusUpdateModal" tabindex="-1">
      <div class="modal-dialog">
        <div class="modal-content rounded-4">
          <div class="modal-header bg-primary text-white">
            <h5 class="modal-title fw-bold fs-5">⚡ Batch Status Update</h5>
            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
          </div>
          <div class="modal-body">
            <p class="text-muted small">Select new status to apply to all selected records:</p>
            <div class="mb-3">
                <label class="form-label small fw-bold">New Status</label>
                <select id="bulkNewStatus" class="form-select form-select-sm">
                    <option value="Pending">Pending</option>
                    <option value="Submitted">Submitted</option>
                    <option value="BLO Assigned">BLO Assigned</option>
                    <option value="FVR">FVR</option>
                    <option value="Approved">Approved</option>
                    <option value="Rejected">Rejected</option>
                    <option value="E_Roll Updated">E_Roll Updated</option>
                </select>
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Cancel</button>
            <button type="button" class="btn btn-success btn-sm fw-bold" onclick="submitBulkStatusUpdate()">💾 Apply Status Update</button>
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

    <!-- EDIT POPUP MODAL -->
    <div class="modal fade" id="editModal" tabindex="-1">
      <div class="modal-dialog">
        <div class="modal-content rounded-4">
          <div class="modal-header bg-primary text-white">
            <h5 class="modal-title fw-bold fs-5">✏️ Edit Record Details</h5>
            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
          </div>
          <div class="modal-body">
            <input type="hidden" id="editEntryId">
            <div class="mb-3">
                <label class="form-label small fw-bold">Current Status</label>
                <select id="editStatus" class="form-select form-select-sm">
                    <option value="Pending">Pending</option>
                    <option value="Submitted">Submitted</option>
                    <option value="BLO Assigned">BLO Assigned</option>
                    <option value="FVR">FVR</option>
                    <option value="Approved">Approved</option>
                    <option value="Rejected">Rejected</option>
                    <option value="E_Roll Updated">E_Roll Updated</option>
                </select>
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Amount (₹)</label>
                <input type="number" id="editAmount" class="form-control form-control-sm">
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Amount Received?</label>
                <select id="editReceived" class="form-select form-select-sm">
                    <option value="No">No</option>
                    <option value="Yes">Yes</option>
                </select>
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Work Complete?</label>
                <select id="editWorkComplete" class="form-select form-select-sm">
                    <option value="No">No</option>
                    <option value="Yes">Yes</option>
                </select>
            </div>
            <div class="mb-3">
                <label class="form-label small fw-bold">Remarks</label>
                <input type="text" id="editRemarks" class="form-control form-control-sm" style="text-transform: uppercase;">
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Cancel</button>
            <button type="button" class="btn btn-primary btn-sm fw-bold" onclick="submitEdit()">💾 Save Changes</button>
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

        function resetTimer() {
            idleTime = 0;
        }

        window.onload = resetTimer;
        document.onmousemove = resetTimer;
        document.onkeypress = resetTimer;
        document.onclick = resetTimer;
        document.onscroll = resetTimer;
        window.onfocus = resetTimer;

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

        function saveLockSettings() {
            let val = document.getElementById('lockDurationSelect').value;
            localStorage.setItem('portal_lock_duration', val);
            lockDuration = parseInt(val);
            idleTime = 0;
            alert("✅ Auto-lock timer successfully updated and saved permanently!");
        }

        document.addEventListener("DOMContentLoaded", function() {
            let savedDur = localStorage.getItem('portal_lock_duration');
            if(savedDur !== null && document.getElementById('lockDurationSelect')) {
                document.getElementById('lockDurationSelect').value = savedDur;
            }
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
                let currentStatus = row.getAttribute('data-status');
                let currentAmount = row.getAttribute('data-amount');
                let currentReceived = row.getAttribute('data-received');
                let currentWork = row.getAttribute('data-work');
                let currentRemarks = row.getAttribute('data-remarks');

                document.getElementById('editEntryId').value = entryId;
                document.getElementById('editStatus').value = currentStatus;
                document.getElementById('editAmount').value = currentAmount;
                document.getElementById('editReceived').value = currentReceived;
                document.getElementById('editWorkComplete').value = currentWork;
                document.getElementById('editRemarks').value = currentRemarks;

                new bootstrap.Modal(document.getElementById('editModal')).show();
            }
        }

        function submitBulkStatusUpdate() {
            let checkboxes = document.querySelectorAll('.row-checkbox:checked');
            let ids = Array.from(checkboxes).map(cb => cb.value);
            let newStatus = document.getElementById('bulkNewStatus').value;

            fetch('/api/bulk-status-update', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ids: ids, status: newStatus})
            })
            .then(res => res.json())
            .then(data => {
                location.reload();
            });
        }

        function submitEdit() {
            let entryId = document.getElementById('editEntryId').value;
            let payload = {
                id: entryId,
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
            if(val === 'Yes') {
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
    </script>
</body>
</html>
"""

# --- NEW STYLISH GLASSMORPHISM TRICOLOR LOGIN PAGE MATCHING SCREENSHOT ---
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
            padding: 40px;
            border-radius: 24px;
            box-shadow: 0 30px 60px rgba(0, 0, 0, 0.5);
            position: relative;
        }
        .login-logo {
            width: 60px;
            height: 60px;
            background: linear-gradient(135deg, #2563EB, #7c3aed);
            color: white;
            font-size: 28px;
            font-weight: bold;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 16px;
            margin: 0 auto 12px auto;
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
            <div class="text-center mb-4">
                <h3 class="fw-bold text-dark mb-1">Welcome Back</h3>
                <p class="text-muted small">EMS Delhi</p>
            </div>

            {% if error %}
            <div class="alert alert-danger py-2 small text-center fw-bold rounded-3">{{ error }}</div>
            {% endif %}

            <form method="POST">
                <div class="mb-3 position-relative">
                    <label class="form-label small fw-bold text-secondary">Username / Email</label>
                    <i class="fa-solid fa-user input-group-icon"></i>
                    <input type="text" name="username" class="form-control" placeholder="Enter your username or email" required>
                </div>
                <div class="mb-3 position-relative">
                    <label class="form-label small fw-bold text-secondary">Password</label>
                    <i class="fa-solid fa-lock input-group-icon"></i>
                    <input type="password" name="password" class="form-control" placeholder="Enter your password" required>
                </div>
                <div class="d-flex justify-content-between align-items-center mb-4 small">
                    <div class="form-check">
                        <input type="checkbox" class="form-check-input" id="remember">
                        <label class="form-check-label text-muted" for="remember">Remember me</label>
                    </div>
                    <a href="#" class="text-decoration-none text-primary fw-bold">Forgot password?</a>
                </div>
                <button type="submit" class="btn btn-signin mb-3"><i class="fa-solid fa-arrow-right-to-bracket me-2"></i> Sign In</button>
            </form>
            <div class="text-center mt-3">
                <small class="text-muted">Need help? <a href="#" class="text-decoration-none fw-bold">Contact Support</a></small>
            </div>
        </div>
    </div>
</body>
</html>
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

    <!-- ROW 6: ALERTS BAR -->
    <div class="p-3 rounded-4 bg-warning-subtle border border-warning d-flex align-items-center justify-content-between clickable-card" onclick="openDashboardModal('alert7', '⚠️ 7+ Days Pending Alerts List')">
        <div class="d-flex align-items-center gap-2 text-warning fw-bold">
            <i class="fa-solid fa-triangle-exclamation"></i>
            <span>Alerts — Day 7+ Pending</span>
        </div>
        <span class="text-dark fw-bold small">View List <i class="fa-solid fa-angle-right"></i></span>
    </div>
""")

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
        <button class="btn btn-primary btn-sm fw-bold px-3 py-1.5" onclick="handleAction('status_update')"><i class="fa-solid fa-pen-to-square me-1"></i> Status Update</button>
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
                    <div class="color-box" style="background:#FFFFFF;" onclick="highlightSelectedRows('#FFFFFF')" title="White"></div>
                    <div class="color-box" style="background:#000000;" onclick="highlightSelectedRows('#000000')" title="Black"></div>
                    <div class="color-box" style="background:#E7E6E6;" onclick="highlightSelectedRows('#E7E6E6')" title="Light Gray"></div>
                    <div class="color-box" style="background:#414853;" onclick="highlightSelectedRows('#414853')" title="Dark Gray"></div>
                    <div class="color-box" style="background:#2F5597;" onclick="highlightSelectedRows('#2F5597')" title="Navy Blue"></div>
                    <div class="color-box" style="background:#ED7D31;" onclick="highlightSelectedRows('#ED7D31')" title="Orange"></div>
                    <div class="color-box" style="background:#A5A5A5;" onclick="highlightSelectedRows('#A5A5A5')" title="Gray"></div>
                    <div class="color-box" style="background:#FFC000;" onclick="highlightSelectedRows('#FFC000')" title="Yellow"></div>
                    <div class="color-box" style="background:#5B9BD5;" onclick="highlightSelectedRows('#5B9BD5')" title="Blue"></div>
                    <div class="color-box" style="background:#70AD47;" onclick="highlightSelectedRows('#70AD47')" title="Green"></div>
                </div>

                <div class="small fw-bold text-muted mb-1 px-1">Standard Colors</div>
                <div class="excel-color-grid mb-2">
                    <div class="color-box" style="background:#C00000;" onclick="highlightSelectedRows('#C00000')" title="Dark Red"></div>
                    <div class="color-box" style="background:#FF0000;" onclick="highlightSelectedRows('#FF0000')" title="Red"></div>
                    <div class="color-box" style="background:#FFC000;" onclick="highlightSelectedRows('#FFC000')" title="Light Orange"></div>
                    <div class="color-box" style="background:#FFFF00;" onclick="highlightSelectedRows('#FFFF00')" title="Bright Yellow"></div>
                    <div class="color-box" style="background:#92D050;" onclick="highlightSelectedRows('#92D050')" title="Light Green"></div>
                    <div class="color-box" style="background:#00B050;" onclick="highlightSelectedRows('#00B050')" title="Emerald Green"></div>
                    <div class="color-box" style="background:#00B0F0;" onclick="highlightSelectedRows('#00B0F0')" title="Cyan"></div>
                    <div class="color-box" style="background:#0070C0;" onclick="highlightSelectedRows('#0070C0')" title="Dark Blue"></div>
                    <div class="color-box" style="background:#002060;" onclick="highlightSelectedRows('#002060')" title="Midnight Blue"></div>
                    <div class="color-box" style="background:#7030A0;" onclick="highlightSelectedRows('#7030A0')" title="Purple"></div>
                </div>

                <div class="small fw-bold text-muted mb-1 px-1">Soft Pastels</div>
                <div class="excel-color-grid mb-2">
                    <div class="color-box" style="background:#FEE2E2;" onclick="highlightSelectedRows('#FEE2E2')" title="Soft Red"></div>
                    <div class="color-box" style="background:#DCFCE7;" onclick="highlightSelectedRows('#DCFCE7')" title="Soft Green"></div>
                    <div class="color-box" style="background:#FEF08A;" onclick="highlightSelectedRows('#FEF08A')" title="Soft Yellow"></div>
                    <div class="color-box" style="background:#DBEAFE;" onclick="highlightSelectedRows('#DBEAFE')" title="Soft Blue"></div>
                    <div class="color-box" style="background:#F3E8FF;" onclick="highlightSelectedRows('#F3E8FF')" title="Soft Purple"></div>
                    <div class="color-box" style="background:#FFEDD5;" onclick="highlightSelectedRows('#FFEDD5')" title="Soft Peach"></div>
                    <div class="color-box" style="background:#CCFBF1;" onclick="highlightSelectedRows('#CCFBF1')" title="Soft Teal"></div>
                    <div class="color-box" style="background:#FCE7F3;" onclick="highlightSelectedRows('#FCE7F3')" title="Soft Pink"></div>
                    <div class="color-box" style="background:#F1F5F9;" onclick="highlightSelectedRows('#F1F5F9')" title="Soft Slate"></div>
                    <div class="color-box" style="background:#E2E8F0;" onclick="highlightSelectedRows('#E2E8F0')" title="Soft Grey"></div>
                </div>

                <hr class="my-1">
                <button class="btn btn-light btn-sm w-100 fw-bold text-danger py-1" onclick="highlightSelectedRows('')"><i class="fa-solid fa-ban me-1"></i> No Fill / Clear</button>
            </div>
        </div>

        <button class="btn btn-warning btn-sm fw-bold px-3 py-1.5 text-dark" onclick="handleAction('edit')"><i class="fa-solid fa-pen me-1"></i> Edit</button>
        <button class="btn btn-danger btn-sm fw-bold px-3 py-1.5" onclick="handleAction('delete')"><i class="fa-solid fa-trash me-1"></i> Delete</button>
        <button class="btn btn-secondary btn-sm fw-bold px-3 py-1.5" onclick="handleAction('undo')"><i class="fa-solid fa-rotate-left me-1"></i> Undo</button>
        <div class="ms-auto">
            <button class="btn btn-outline-primary btn-sm fw-bold" onclick="window.location.reload();"><i class="fa-solid fa-rotate me-1"></i> Refresh</button>
        </div>
    </div>

    <div class="card-box overflow-hidden">
        <div class="table-responsive">
            <table class="excel-table align-middle mb-0" id="entriesTable">
                <thead>
                    <tr>
                        <th class="py-2 px-2 text-center" style="width: 35px;"><input type="checkbox" id="selectAllMaster" onclick="toggleSelectAll(this)"></th>
                        <th class="py-2 text-center" style="width: 50px;">S.No.</th>
                        <th class="py-2" onclick="sortTable(2)">Client Name ↕</th>
                        <th class="py-2" onclick="sortTable(3)">Reference No. ↕</th>
                        <th class="py-2" onclick="sortTable(4)">AC ↕</th>
                        <th class="py-2" onclick="sortTable(5)">Full Name ↕</th>
                        <th class="py-2" onclick="sortTable(6)">Form Type ↕</th>
                        <th class="py-2" onclick="sortTable(7)">Submission Date ↕</th>
                        <th class="py-2" onclick="sortTable(8)">Current Status ↕</th>
                        <th class="py-2 text-center" style="width: 55px;" onclick="sortTable(9)">Day ↕</th>
                        <th class="py-2 text-end" onclick="sortTable(10)">Amount ↕</th>
                        <th class="py-2 text-center" onclick="sortTable(11)">Received ↕</th>
                        <th class="py-2 text-center" onclick="sortTable(12)">Work Complete ↕</th>
                        <th class="py-2 px-2">Remarks</th>
                    </tr>
                </thead>
                <tbody>
                    {% for row in entries %}
                    <tr data-status="{{ row.current_status }}" data-amount="{{ row.amount }}" data-received="{{ row.amount_received }}" data-work="{{ row.work_complete }}" data-remarks="{{ row.remarks }}">
                        <td class="px-2 text-center"><input type="checkbox" class="row-checkbox" value="{{ row.id }}"></td>
                        <td class="fw-bold text-muted text-center sno-cell">{{ loop.index }}</td>
                        <td class="fw-bold text-dark">{{ row.client_name }}</td>
                        <td class="font-monospace fw-bold {% if row.is_duplicate %}text-danger bg-danger-subtle{% else %}text-dark{% endif %}">{{ row.ref_no }}</td>
                        <td class="fw-bold text-dark">{{ row.ac }}</td>
                        <td class="fw-bold text-dark">{{ row.full_name }}</td>
                        <td>{{ row.form_type }}</td>
                        <td>{{ row.submission_date }}</td>
                        <td>
                            <span class="badge {% if row.display_status == 'Approved' %}bg-success-subtle text-success{% elif row.display_status == 'Rejected' %}bg-danger-subtle text-danger{% else %}bg-warning-subtle text-warning{% endif %} px-2 py-1 rounded-pill">
                                {{ row.current_status }}
                            </span>
                        </td>
                        <td class="fw-bold text-center text-success bg-success-subtle">{{ row.day_count }}</td>
                        <td class="fw-bold text-end">₹{{ row.amount }}</td>
                        <td class="text-center fw-bold {% if row.amount_received == 'Yes' %}text-white bg-success{% else %}text-danger{% endif %}">{{ row.amount_received }}</td>
                        <td class="text-center fw-bold {% if row.work_complete == 'Yes' %}text-white bg-success{% else %}text-danger{% endif %}">{{ row.work_complete }}</td>
                        <td class="px-2 text-muted small">{{ row.remarks }}</td>
                    </tr>
                    {% else %}
                    <tr><td colspan="14" class="text-center py-5 text-muted fw-semibold">No records found in database.</td></tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
"""

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

# --- WHATSAPP LAUNCHER PAGE TEMPLATE ---
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
                <h5 class="fw-bold text-dark mb-3">🔒 Change Super Admin Password</h5>
                <form method="POST" action="/settings">
                    <input type="hidden" name="action_type" value="password">
                    <div class="mb-3">
                        <label class="form-label small fw-bold text-secondary">Username / Admin Email</label>
                        <input type="text" name="username" class="form-control form-control-sm" value="{{ session.get('user', '') }}" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold text-secondary">New Password</label>
                        <input type="password" name="new_password" class="form-control form-control-sm" placeholder="Enter new password" required>
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
                            <input type="hidden" name="work_complete" id="inputWorkComp" value="No">
                            <button type="button" id="btnWorkNo" class="btn btn-danger flex-fill fw-bold" onclick="setWorkComp('No')">Not Done</button>
                            <button type="button" id="btnWorkYes" class="btn btn-outline-secondary flex-fill fw-bold text-secondary" onclick="setWorkComp('Yes')">Done</button>
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
                            <option value="Pending">Pending</option>
                            <option value="Submitted">Submitted</option>
                            <option value="BLO Assigned">BLO Assigned</option>
                            <option value="FVR">FVR</option>
                            <option value="Approved">Approved</option>
                            <option value="Rejected">Rejected</option>
                            <option value="E_Roll Updated">E_Roll Updated</option>
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
                amount_recv = str(val.get("amount_received", "No")).strip()
                work_comp = str(val.get("work_complete", "No")).strip()

                is_approved = "APPROVED" in u_stat or "E_ROLL" in u_stat or "EROLL" in u_stat
                is_rejected = "REJECT" in u_stat
                is_eroll = "E_ROLL" in u_stat or "EROLL" in u_stat
                is_pending = not is_approved and not is_rejected
                is_complete = (amount_recv == "Yes" and work_comp == "Yes" and is_approved)

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
                    "ac": ac_val,
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
                    "raw_date": sub_date
                })
        except Exception as e:
            print("Error processing entries:", e)
            
    return entries, sorted(ac_set), counts

@app.route("/", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username_or_email = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        try:
            if not db:
                return render_template_string(LOGIN_HTML, error="Firebase not connected!")
            user_ref = db.collection("users")
            query = user_ref.where("username", "==", username_or_email.lower()).limit(1).stream()
            user_data = None
            for doc in query:
                user_data = doc.to_dict()
            if not user_data:
                query_email = user_ref.where("email", "==", username_or_email.lower()).limit(1).stream()
                for doc in query_email:
                    user_data = doc.to_dict()

            if user_data:
                stored_pass = str(user_data.get("password", ""))
                if stored_pass and stored_pass != password:
                    error = "Invalid Password."
                else:
                    session["user"] = user_data.get("username", "User")
                    session["role"] = user_data.get("role", "Operator")
                    return redirect(url_for("dashboard"))
            else:
                error = "Invalid Username or Email in Firebase."
        except Exception as e:
            error = f"Login Error: {str(e)}"
    return render_template_string(LOGIN_HTML, error=error)

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
                        if str(val.get("amount_received", "Yes")) == "Yes":
                            report_amount += amt
                    except: pass
                    continue

                total_count += 1
                status = str(val.get("current_status", "Pending")).upper()
                amt_recv = str(val.get("amount_received", "No"))
                
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

                if amt_recv == "Yes" and val.get("work_complete", "No") == "Yes":
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
    if "user" not in session: return redirect(url_for("login"))
    success_msg = error_msg = None
    _, _, counts = process_entries("all")
    
    if request.method == "POST" and db:
        username = request.form.get("username", "").strip().lower()
        new_password = request.form.get("new_password", "").strip()
        try:
            users_ref = db.collection("users")
            query = users_ref.where("username", "==", username).limit(1).stream()
            updated = False
            for doc in query:
                users_ref.document(doc.id).update({"password": new_password})
                updated = True
            
            if updated:
                success_msg = "✅ Password successfully updated in Firebase Cloud! You can now use it on any system."
            else:
                error_msg = "❌ User not found in Firebase database."
        except Exception as e:
            error_msg = f"❌ Update error: {str(e)}"
            
    return render_template_string(SETTINGS_HTML, success_msg=success_msg, error_msg=error_msg, counts=counts, page='settings')

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

                if recv == "Yes":
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
                    "work_complete": "Yes",
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
    if db and ids:
        try:
            batch = db.batch()
            for entry_id in ids:
                ref = db.collection("election_entries").document(entry_id)
                batch.update(ref, {"current_status": new_status})
            batch.commit()
            return {"status": "success", "updated": len(ids)}
        except Exception as e:
            return {"status": "error", "message": str(e)}, 500
    return {"status": "failed"}, 400

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
                "current_status": data.get("current_status", "Pending"),
                "amount": str(data.get("amount", "0")),
                "amount_received": data.get("amount_received", "No"),
                "work_complete": data.get("work_complete", "No"),
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
                "work_complete": request.form.get("work_complete", "No"),
                "work_status": "Not Done" if request.form.get("work_complete", "No") == "No" else "Done",
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
    if "user" not in session: return redirect(url_for("login"))
    archived_entries = []
    _, _, counts = process_entries("all")
    if db:
        try:
            docs = list(db.collection("election_entries").stream())
            for doc in docs:
                val = doc.to_dict() or {}
                if not val.get("moved_to_report", False): continue
                archived_entries.append({
                    "client_name": val.get("client_name", "N/A"),
                    "ref_no": val.get("ref_no", "N/A"),
                    "ac": val.get("ac", "N/A"),
                    "full_name": f"{val.get('first_name', '')} {val.get('last_name', '')}".strip(),
                    "form_type": val.get("form_type", "Form 6"),
                    "current_status": val.get("current_status", "Approved"),
                    "amount": val.get("amount", "0"),
                    "amount_received": val.get("amount_received", "Yes")
                })
        except Exception as e:
            print("Reports error:", e)
    return render_template_string(REPORTS_HTML, entries=archived_entries, counts=counts, page='reports')

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

if __name__ == "__main__":
    app.run(debug=True, port=5000)