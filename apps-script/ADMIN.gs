// ═══════════════════════════════════════════════════════════════════════════════
// PHYSICS FOUNDATION — ADMIN DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════
//
// HOW TO DEPLOY (separate from student portal)
// ─────────────────────────────────────────────
// 1. In Apps Script editor: Deploy → New deployment
// 2. Type: Web app
// 3. Description: Admin Dashboard
// 4. Execute as: Me (misra.ravikant@gmail.com)   ← IMPORTANT
// 5. Who has access: Anyone with a Google Account
// 6. Deploy → copy the URL
// 7. Paste that URL into physics-portal/admin.html (ADMIN_APP_URL)
//
// Only misra.ravikant@gmail.com can use it — everyone else is blocked.
// ═══════════════════════════════════════════════════════════════════════════════

var ADMIN_EMAIL_ADDR = 'misra.ravikant@gmail.com';

function doGet(e) {
  // Block anyone who is not the admin
  var caller = Session.getActiveUser().getEmail();
  if (caller !== ADMIN_EMAIL_ADDR) {
    return HtmlService.createHtmlOutput(
      '<html><head><style>body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;background:#0f172a;color:white;}</style></head>' +
      '<body><div style="text-align:center;"><h2>Access Denied</h2><p>This page is only accessible to the portal administrator.</p></div></body></html>'
    ).setTitle('Admin — Access Denied');
  }

  var action = e && e.parameter && e.parameter.action ? e.parameter.action : 'dashboard';

  // Handle POST-like actions via URL params
  if (action === 'unlock') {
    return _adminActionUnlock(e);
  } else if (action === 'addStudent') {
    return _adminActionAddStudent(e);
  }

  return _adminDashboard();
}

// ─── MAIN DASHBOARD ───────────────────────────────────────────────────────────
function _adminDashboard() {
  var cfg      = _cfg();
  var ss       = SpreadsheetApp.openById(cfg.SHEET_ID);
  var students = _students(ss);
  var units    = _units(ss).filter(function(u){ return u.LessonURL !== ''; });
  var emailLog = _getEmailLog(ss);

  // ── Progress grid ────────────────────────────────────────────────────────────
  var statusIcon = { complete:'✅', available:'📖', in_progress:'✏️', awaiting_review:'🟡', corrections:'❌', locked:'🔒' };
  var statusBg   = { complete:'#dcfce7', available:'#dbeafe', in_progress:'#dbeafe', awaiting_review:'#fef9c3', corrections:'#fee2e2', locked:'#f1f5f9' };

  var gridHeaders = '<th class="sticky-col">Student</th>' +
    units.map(function(u){ return '<th class="unit-th" title="'+u.UnitName+'">'+u.UnitID+'</th>'; }).join('');

  var gridRows = students.map(function(student) {
    var progress = _progress(ss, student.StudentID);
    var progMap  = {};
    progress.forEach(function(p){ progMap[p.UnitID] = p; });

    var cells = units.map(function(u) {
      var p      = progMap[u.UnitID] || { Status:'locked' };
      var status = p.Status || 'locked';
      var icon   = statusIcon[status] || '🔒';
      var bg     = statusBg[status]   || '#f1f5f9';
      var canUnlock = (status === 'locked' || status === 'corrections');
      var unlockBtn = canUnlock
        ? '<a href="?action=unlock&sid='+student.StudentID+'&uid='+u.UnitID+'" class="unlock-btn" onclick="return confirm(\'Unlock '+u.UnitName+' for '+student.StudentName+'?\')">unlock</a>'
        : '';
      return '<td style="background:'+bg+';text-align:center;" title="'+student.StudentName+' — '+u.UnitName+': '+status+'">'+icon+'<br>'+unlockBtn+'</td>';
    }).join('');

    return '<tr><td class="sticky-col"><strong>'+student.StudentName+'</strong></td>'+cells+'</tr>';
  }).join('');

  // ── Pending reviews ───────────────────────────────────────────────────────────
  var pendingRows = '';
  students.forEach(function(student) {
    var progress = _progress(ss, student.StudentID);
    progress.filter(function(p){ return p.Status === 'awaiting_review'; }).forEach(function(p) {
      var u = _unit(ss, p.UnitID);
      pendingRows += '<tr>'+
        '<td>'+student.StudentName+'</td>'+
        '<td>'+p.UnitID+'</td>'+
        '<td>'+(u?u.UnitName:'')+'</td>'+
        '<td>'+_formatDate(p.HomeworkSubmittedAt)+'</td>'+
        '<td>'+(p.HomeworkDriveURL ? '<a href="'+p.HomeworkDriveURL+'" target="_blank">View</a>' : '—')+'</td>'+
        '<td><a href="?action=unlock&sid='+student.StudentID+'&uid='+p.UnitID+'" class="btn-sm btn-green" onclick="return confirm(\'Mark complete and unlock next?\')">✅ Approve & Unlock Next</a></td>'+
        '</tr>';
    });
  });
  if (!pendingRows) pendingRows = '<tr><td colspan="6" style="text-align:center;color:#94a3b8;padding:20px;">No pending submissions</td></tr>';

  // ── Homework completion summary ───────────────────────────────────────────────
  var hwRows = '';
  var hwCount = 0;
  students.forEach(function(student) {
    var progress = _progress(ss, student.StudentID);
    var submitted = progress.filter(function(p) {
      return p.HomeworkSubmittedAt && p.HomeworkSubmittedAt !== '';
    });
    submitted.forEach(function(p) {
      var u = _unit(ss, p.UnitID);
      var statusBadge = {
        complete:        '<span class="hw-badge hw-done">✅ Complete</span>',
        awaiting_review: '<span class="hw-badge hw-pending">🟡 Awaiting Review</span>',
        corrections:     '<span class="hw-badge hw-corr">❌ Corrections</span>',
        available:       '<span class="hw-badge hw-sub">📤 Submitted</span>',
      }[p.Status] || '<span class="hw-badge">'+p.Status+'</span>';

      var parentDecision = p.ParentDecision
        ? (p.ParentDecision === 'approved' ? '✅ Approved' : '❌ Corrections Requested')
        : '⏳ Pending';

      hwRows += '<tr>'+
        '<td><strong>'+student.StudentName+'</strong></td>'+
        '<td style="font-family:monospace;font-size:.8rem;">'+p.UnitID+'</td>'+
        '<td>'+(u ? u.UnitName : '')+'</td>'+
        '<td>'+_formatDate(p.HomeworkSubmittedAt)+'</td>'+
        '<td>'+statusBadge+'</td>'+
        '<td>'+parentDecision+'</td>'+
        '<td>'+(p.ParentComments ? '<span title="'+p.ParentComments+'" style="cursor:help;color:#64748b;">💬 '+p.ParentComments.substring(0,40)+(p.ParentComments.length>40?'…':'')+'</span>' : '—')+'</td>'+
        '<td>'+(p.HomeworkDriveURL ? '<a href="'+p.HomeworkDriveURL+'" target="_blank" class="btn-sm" style="background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;">View</a>' : '—')+'</td>'+
        '</tr>';
      hwCount++;
    });
  });
  if (!hwRows) hwRows = '<tr><td colspan="8" style="text-align:center;color:#94a3b8;padding:20px;">No homework submitted yet</td></tr>';
    return '<tr><td>'+_formatDate(r[0])+'</td><td>'+r[1]+'</td><td>'+r[2]+'</td><td>'+r[3]+'</td><td>'+r[4]+'</td><td class="'+(r[6]==='sent'?'sent':'error')+'">'+r[6]+'</td></tr>';
  }).join('');

  // ── Email log ─────────────────────────────────────────────────────────────────
  var logRows = emailLog.slice(0,20).map(function(r){
    return '<tr><td>'+_formatDate(r[0])+'</td><td>'+r[1]+'</td><td>'+r[2]+'</td><td>'+r[3]+'</td><td>'+r[4]+'</td><td class="'+(r[6]==='sent'?'sent':'error')+'">'+r[6]+'</td></tr>';
  }).join('');
  var totalComplete = 0, totalPending = 0, totalCorrections = 0;
  students.forEach(function(s) {
    _progress(ss, s.StudentID).forEach(function(p) {
      if (p.Status==='complete') totalComplete++;
      else if (p.Status==='awaiting_review') totalPending++;
      else if (p.Status==='corrections') totalCorrections++;
    });
  });

  var html = '<!DOCTYPE html><html lang="en"><head>'+
    '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'+
    '<title>Admin — Physics Foundations</title>'+
    '<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;700;800;900&family=Space+Mono&display=swap" rel="stylesheet">'+
    _adminStyles()+
    '</head><body>'+

    // Header
    '<header>'+
    '<div class="header-inner">'+
    '<div><div class="lbl">Physics Foundations — Admin</div><h1>Dashboard</h1></div>'+
    '<a href="https://misra-ravi.github.io/physics-foundation/" class="back-link">← Public Site</a>'+
    '</div>'+
    '</header>'+

    // Stats
    '<div class="stats-bar">'+
    '<div class="stat"><div class="num">'+students.length+'</div><div class="lbl">Students</div></div>'+
    '<div class="stat"><div class="num">'+totalComplete+'</div><div class="lbl">Units Complete</div></div>'+
    '<div class="stat"><div class="num">'+hwCount+'</div><div class="lbl">HW Submitted</div></div>'+
    '<div class="stat"><div class="num">'+totalPending+'</div><div class="lbl">Awaiting Review</div></div>'+
    '<div class="stat"><div class="num">'+totalCorrections+'</div><div class="lbl">Needs Corrections</div></div>'+
    '</div>'+

    '<main>'+

    // Add student form
    '<section>'+
    '<h2>Add New Student</h2>'+
    '<form action="?action=addStudent" method="get" class="add-form">'+
    '<input type="hidden" name="action" value="addStudent">'+
    '<div class="form-grid">'+
    '<input name="sid"         placeholder="Student ID (e.g. s004)"   required>'+
    '<input name="name"        placeholder="Full Name"                 required>'+
    '<input name="email"       placeholder="Student Gmail"             required type="email">'+
    '<input name="parentEmail" placeholder="Parent Gmail"              required type="email">'+
    '<input name="parentName"  placeholder="Parent Name (e.g. Mr & Mrs Khan)" required>'+
    '<button type="submit" class="btn-primary">+ Enrol Student</button>'+
    '</div>'+
    '</form>'+
    '</section>'+

    // Pending reviews
    '<section>'+
    '<h2>Pending Parent Reviews</h2>'+
    '<table><thead><tr><th>Student</th><th>Unit ID</th><th>Unit</th><th>Submitted</th><th>Homework</th><th>Action</th></tr></thead>'+
    '<tbody>'+pendingRows+'</tbody></table>'+
    '</section>'+

    // Progress grid
    '<section>'+
    '<h2>Progress Grid</h2>'+
    '<p class="hint">Click "unlock" in any 🔒 cell to unlock that unit for the student. An email is sent automatically.</p>'+
    '<div class="grid-scroll"><table class="grid-table">'+
    '<thead><tr>'+gridHeaders+'</tr></thead>'+
    '<tbody>'+gridRows+'</tbody>'+
    '</table></div>'+
    '</section>'+

    // Homework completion
    '<section>'+
    '<h2>Homework Submissions ('+hwCount+' total)</h2>'+
    '<table><thead><tr>'+
    '<th>Student</th><th>Unit ID</th><th>Unit Name</th><th>Submitted</th>'+
    '<th>Status</th><th>Parent Decision</th><th>Comments</th><th>Homework</th>'+
    '</tr></thead><tbody>'+hwRows+'</tbody></table>'+
    '</section>'+

    // Email log
    '<section>'+
    '<h2>Recent Email Log</h2>'+
    '<table><thead><tr><th>Time</th><th>Type</th><th>Student</th><th>Unit</th><th>Recipient</th><th>Status</th></tr></thead>'+
    '<tbody>'+logRows+'</tbody></table>'+
    '</section>'+

    '</main>'+
    '<footer>Physics Foundations by Ravi &nbsp;·&nbsp; Admin Dashboard &nbsp;·&nbsp; misra.ravikant@gmail.com</footer>'+
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Admin — Physics Foundations')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── ACTION: UNLOCK UNIT ──────────────────────────────────────────────────────
function _adminActionUnlock(e) {
  var sid    = e.parameter.sid;
  var uid    = e.parameter.uid;
  var cfg    = _cfg();
  var ss     = SpreadsheetApp.openById(cfg.SHEET_ID);
  var student = _studentById(ss, sid);
  var unit    = _unit(ss, uid);

  if (!student || !unit) {
    return HtmlService.createHtmlOutput('<p>Error: student or unit not found. <a href="?">Back</a></p>');
  }

  // Mark current unit complete if it was awaiting_review
  var prog = _prog(ss, sid, uid);
  if (prog && prog.Status === 'awaiting_review') {
    _setProg(ss, sid, uid, { Status:'complete', ParentReviewedAt:new Date(), ParentDecision:'approved' });
  } else {
    _setProg(ss, sid, uid, { Status:'available', UnlockedAt:new Date() });
  }

  // Find and unlock next unit
  var nextUnit = _nextUnit(ss, uid);
  if (nextUnit) {
    _setProg(ss, sid, nextUnit.UnitID, { Status:'available', UnlockedAt:new Date() });
  }

  // Email student AND parent separately with richer content
  var studentBody = _unlockEmailHtml(student, nextUnit || unit, 'student');
  var parentBody  = _unlockEmailHtml(student, nextUnit || unit, 'parent');
  var subj        = '🔓 New unit unlocked — '+(nextUnit ? nextUnit.UnitName : unit.UnitName);
  var parentSubj  = '🔓 [Physics Foundations] '+student.StudentName+'\'s next unit is unlocked';
  try {
    GmailApp.sendEmail(student.StudentEmail, subj, _stripHtml(studentBody), {
      htmlBody:studentBody, name:PORTAL_NAME
    });
    GmailApp.sendEmail(student.ParentEmail, parentSubj, _stripHtml(parentBody), {
      htmlBody:parentBody, name:PORTAL_NAME
    });
    _logEmail(ss,'admin_unlock',sid,uid,student.StudentEmail,subj,'sent');
    _logEmail(ss,'admin_unlock',sid,uid,student.ParentEmail,parentSubj,'sent');
  } catch(err) {
    _logEmail(ss,'admin_unlock',sid,uid,student.StudentEmail,subj,'error:'+err);
  }

  // Redirect back to dashboard with success message
  return HtmlService.createHtmlOutput(
    '<html><head><meta http-equiv="refresh" content="2;url=?"><style>body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh;background:#0f172a;color:white;text-align:center;}</style></head>'+
    '<body><div><h2>✅ Done</h2><p>Unlocked <strong>'+(nextUnit?nextUnit.UnitName:unit.UnitName)+'</strong> for <strong>'+student.StudentName+'</strong>.</p><p>Email sent. Redirecting...</p></div></body></html>'
  );
}

// ─── ACTION: ADD STUDENT ──────────────────────────────────────────────────────
function _adminActionAddStudent(e) {
  var p   = e.parameter;
  var cfg = _cfg();
  var ss  = SpreadsheetApp.openById(cfg.SHEET_ID);
  var sh  = ss.getSheetByName('Roster');

  // Check ID not already used
  var existing = sh.getDataRange().getValues().map(function(r){ return r[0]; });
  if (existing.indexOf(p.sid) >= 0) {
    return HtmlService.createHtmlOutput(
      '<html><head><meta http-equiv="refresh" content="3;url=?"></head>'+
      '<body style="font-family:sans-serif;text-align:center;padding:60px;background:#0f172a;color:white;">'+
      '<h2>⚠️ Student ID '+p.sid+' already exists.</h2><p>Redirecting back...</p></body></html>'
    );
  }

  sh.appendRow([p.sid, p.name, p.email, p.parentEmail, p.parentName, new Date(), true]);

  // Seed progress — first 3 unlocked
  var units   = ss.getSheetByName('Units').getDataRange().getValues().slice(1);
  var progSh  = ss.getSheetByName('Progress');
  var newRows = units.map(function(u, idx) {
    return [p.sid+'_'+u[0], p.sid, u[0], idx<3?'available':'locked','','','','','','','',''];
  });
  if (newRows.length > 0) progSh.getRange(progSh.getLastRow()+1, 1, newRows.length, 12).setValues(newRows);

  // Welcome emails
  _sendWelcomeEmail(ss, [p.sid, p.name, p.email, p.parentEmail, p.parentName, new Date(), true]);

  // Get first 3 unlocked units for confirmation display
  var first3 = units.slice(0,3).map(function(u){ return u[7]; }).join(', ');

  // Return confirmation page with progress summary before redirecting
  return HtmlService.createHtmlOutput(
    '<html><head>'+
    '<meta http-equiv="refresh" content="4;url=?">'+
    '<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@700;800;900&display=swap" rel="stylesheet">'+
    '<style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:Nunito,sans-serif;background:#0f172a;color:white;display:flex;align-items:center;justify-content:center;min-height:100vh;}'+
    '.box{text-align:center;padding:48px 32px;max-width:480px;}'+
    'h2{font-size:1.6rem;font-weight:900;margin-bottom:12px;}'+
    'p{color:rgba(255,255,255,.7);line-height:1.8;margin-bottom:12px;font-size:.95rem;}'+
    '.badge{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);border-radius:10px;padding:12px 20px;margin:16px 0;font-size:.88rem;}'+
    '.badge strong{color:#4ade80;display:block;margin-bottom:6px;}'+
    '.progress-bar{height:4px;background:rgba(255,255,255,.1);border-radius:2px;margin-top:24px;overflow:hidden;}'+
    '.progress-fill{height:100%;background:#1d4ed8;animation:fill 4s linear forwards;}'+
    '@keyframes fill{from{width:0}to{width:100%}}</style>'+
    '</head><body><div class="box">'+
    '<div style="font-size:3rem;margin-bottom:16px;">✅</div>'+
    '<h2>'+p.name+' enrolled!</h2>'+
    '<div class="badge"><strong>📖 First 3 units unlocked</strong>'+first3+'</div>'+
    '<div class="badge"><strong>📧 Welcome emails sent to</strong>'+p.email+'<br>'+p.parentEmail+'</div>'+
    '<p style="margin-top:16px;font-size:.82rem;color:rgba(255,255,255,.4);">Returning to dashboard in 4 seconds...</p>'+
    '<div class="progress-bar"><div class="progress-fill"></div></div>'+
    '</div></body></html>'
  );
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function _getEmailLog(ss) {
  var sh = ss.getSheetByName('EmailLog');
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues().slice(1);
  return data.reverse(); // newest first
}

function _formatDate(d) {
  if (!d) return '—';
  try {
    var dt = new Date(d);
    return dt.toLocaleDateString('en-GB', {day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
  } catch(e){ return String(d); }
}

// ─── ADMIN CSS ────────────────────────────────────────────────────────────────
function _adminStyles() {
  return '<style>'+
    '*{box-sizing:border-box;margin:0;padding:0;}'+
    'body{font-family:"Nunito",sans-serif;background:#f0f4f8;color:#0f172a;}'+
    'header{background:linear-gradient(135deg,#0f172a,#1e3a5f);color:white;padding:24px 32px;}'+
    '.header-inner{max-width:1400px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;}'+
    '.header-inner .lbl{font-size:.72rem;text-transform:uppercase;letter-spacing:2px;opacity:.6;margin-bottom:4px;}'+
    '.header-inner h1{font-size:1.6rem;font-weight:900;}'+
    '.back-link{color:rgba(255,255,255,.7);font-size:.82rem;font-weight:700;text-decoration:none;}'+
    '.stats-bar{background:#1e293b;padding:16px 32px;display:flex;gap:32px;max-width:100%;}'+
    '.stat{text-align:center;}.stat .num{font-size:1.6rem;font-weight:900;color:white;}'+
    '.stat .lbl{font-size:.7rem;text-transform:uppercase;letter-spacing:1px;color:#94a3b8;}'+
    'main{max-width:1400px;margin:0 auto;padding:28px 24px 80px;}'+
    'section{background:white;border-radius:16px;padding:24px;margin-bottom:24px;box-shadow:0 2px 12px rgba(0,0,0,.07);}'+
    'h2{font-size:1rem;font-weight:800;margin-bottom:16px;color:#0f172a;}'+
    '.hint{font-size:.8rem;color:#64748b;margin-bottom:12px;}'+
    'table{width:100%;border-collapse:collapse;}'+
    'thead th{background:#f1f5f9;font-size:.72rem;font-weight:800;text-transform:uppercase;letter-spacing:.5px;padding:8px 12px;text-align:left;border-bottom:2px solid #e2e8f0;}'+
    'td{padding:8px 12px;font-size:.85rem;border-bottom:1px solid #f1f5f9;vertical-align:middle;}'+
    '.sent{color:#15803d;font-weight:700;}.error{color:#dc2626;font-weight:700;}'+
    '.form-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px;align-items:end;}'+
    '.form-grid input{border:2px solid #e2e8f0;border-radius:8px;padding:9px 14px;font-family:inherit;font-size:.88rem;width:100%;}'+
    '.form-grid input:focus{outline:none;border-color:#3b82f6;}'+
    '.btn-primary{background:#1d4ed8;color:white;border:none;border-radius:8px;padding:10px 20px;font-weight:800;font-size:.88rem;cursor:pointer;font-family:inherit;width:100%;}'+
    '.btn-sm{display:inline-block;padding:4px 10px;border-radius:6px;font-size:.72rem;font-weight:800;text-decoration:none;white-space:nowrap;}'+
    '.btn-green{background:#22c55e;color:white;}'+
    '.unlock-btn{display:block;font-size:.65rem;color:#1d4ed8;text-decoration:none;font-weight:800;margin-top:2px;}'+
    '.unlock-btn:hover{text-decoration:underline;}'+
    '.hw-badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:.72rem;font-weight:800;}'+
    '.hw-done{background:#dcfce7;color:#15803d;}'+
    '.hw-pending{background:#fef9c3;color:#854d0e;}'+
    '.hw-corr{background:#fee2e2;color:#dc2626;}'+
    '.hw-sub{background:#dbeafe;color:#1d4ed8;}'+
    '.grid-scroll{overflow-x:auto;max-height:500px;overflow-y:auto;}'+
    '.grid-table{border-collapse:collapse;font-size:.78rem;}'+
    '.grid-table th,.grid-table td{border:1px solid #e2e8f0;padding:4px 6px;text-align:center;white-space:nowrap;}'+
    '.unit-th{font-size:.65rem;font-weight:700;background:#f1f5f9;writing-mode:vertical-rl;transform:rotate(180deg);padding:8px 4px;max-height:80px;}'+
    '.sticky-col{position:sticky;left:0;background:white;z-index:1;font-weight:700;min-width:140px;text-align:left !important;border-right:2px solid #e2e8f0;}'+
    'footer{background:#0f172a;color:rgba(255,255,255,.4);text-align:center;padding:20px;font-size:.78rem;}'+
    '</style>';
}
