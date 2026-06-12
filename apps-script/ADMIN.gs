// ═══════════════════════════════════════════════════════════════════════════════
// PHYSICS FOUNDATION — ADMIN DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────
// The doGet router in BOOTSTRAP.gs handles both deployments:
// - Admin deployment (Execute as: Me) → shows admin dashboard
// - Student deployment (Execute as: User) → shows student dashboard
// No doGet needed in this file.
// ═══════════════════════════════════════════════════════════════════════════════

var ADMIN_EMAIL_ADDR = 'misra.ravikant@gmail.com';

// ═══════════════════════════════════════════════════════════════════════════════
// SERVER FUNCTIONS — called via google.script.run from the client
// ═══════════════════════════════════════════════════════════════════════════════

function adminAddStudent(p) {
  if (!p.sid || !p.name || !p.email || !p.parentEmail || !p.parentName)
    return {ok: false, msg: 'All 5 fields are required.'};
  var ss = SpreadsheetApp.openById(_cfg().SHEET_ID);
  var sh = ss.getSheetByName('Roster');
  var ids = sh.getDataRange().getValues().map(function(r){ return String(r[0]).toLowerCase().trim(); });
  if (ids.indexOf(p.sid.toLowerCase().trim()) >= 0)
    return {ok: false, msg: 'Student ID ' + p.sid + ' already exists. Choose a different ID.'};
  sh.appendRow([p.sid, p.name, p.email, p.parentEmail, p.parentName, new Date(), true]);
  var units  = ss.getSheetByName('Units').getDataRange().getValues().slice(1);
  var progSh = ss.getSheetByName('Progress');
  var rows   = units.map(function(u, i) {
    return [p.sid+'_'+u[0], p.sid, u[0], i < 3 ? 'available' : 'locked', '','','','','','','',''];
  });
  if (rows.length > 0) progSh.getRange(progSh.getLastRow()+1, 1, rows.length, 12).setValues(rows);
  SpreadsheetApp.flush();
  try { _sendWelcomeEmail(ss, [p.sid, p.name, p.email, p.parentEmail, p.parentName, new Date(), true]); } catch(e){}
  var first3 = units.slice(0,3).map(function(u){ return u[7]; }).join(', ');
  return {ok: true, msg: p.name + ' enrolled! First 3 units unlocked: ' + first3 + '. Welcome emails sent.'};
}

function adminUnlock(sid, uid) {
  var ss  = SpreadsheetApp.openById(_cfg().SHEET_ID);
  var stu = _studentById(ss, sid);
  var u   = _unit(ss, uid);
  if (!stu || !u) return {ok: false, msg: 'Student or unit not found.'};
  var prog = _prog(ss, sid, uid);
  if (prog && prog.Status === 'awaiting_review')
    _setProg(ss, sid, uid, {Status:'complete', ParentReviewedAt:new Date(), ParentDecision:'approved'});
  else
    _setProg(ss, sid, uid, {Status:'available', UnlockedAt:new Date()});
  var nxt = _nextUnit(ss, uid);
  if (nxt) _setProg(ss, sid, nxt.UnitID, {Status:'available', UnlockedAt:new Date()});
  var ul = nxt || u;
  var sb = _unlockEmailHtml(stu, ul, 'student');
  var pb = _unlockEmailHtml(stu, ul, 'parent');
  var ss2 = '🔓 New unit unlocked — ' + ul.UnitName;
  var ps2 = '🔓 ' + stu.StudentName + "'s next unit is unlocked — Physics Foundations";
  try {
    GmailApp.sendEmail(stu.StudentEmail, ss2, _stripHtml(sb), {htmlBody:sb, name:PORTAL_NAME});
    GmailApp.sendEmail(stu.ParentEmail,  ps2, _stripHtml(pb), {htmlBody:pb, name:PORTAL_NAME});
    _logEmail(ss, 'admin_unlock', sid, uid, stu.StudentEmail, ss2, 'sent');
    _logEmail(ss, 'admin_unlock', sid, uid, stu.ParentEmail,  ps2, 'sent');
  } catch(e2){ _logEmail(ss, 'admin_unlock', sid, uid, stu.StudentEmail, ss2, 'error:'+e2); }
  return {ok: true, msg: ul.UnitName + ' unlocked for ' + stu.StudentName + '. Emails sent to student and parent.'};
}

function adminRemove(sid) {
  var ss   = SpreadsheetApp.openById(_cfg().SHEET_ID);
  var sh   = ss.getSheetByName('Roster');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(sid)) {
      sh.getRange(i+1, 7).setValue(false);
      SpreadsheetApp.flush();
      return {ok: true, msg: data[i][1] + ' has been deactivated. Progress history preserved.'};
    }
  }
  return {ok: false, msg: 'Student ID ' + sid + ' not found.'};
}

function getAdminHtml() {
  return _buildDashboard().getContent();
}

// ═══════════════════════════════════════════════════════════════════════════════
// DASHBOARD HTML
// ═══════════════════════════════════════════════════════════════════════════════

function _buildDashboard() {
  var ss  = SpreadsheetApp.openById(_cfg().SHEET_ID);
  var stu = _students(ss);
  var un  = _units(ss).filter(function(u){ return u.LessonURL !== ''; });
  var el  = _getEmailLog(ss);
  var now = new Date();

  var pd  = ss.getSheetByName('Progress').getDataRange().getValues();
  var apm = {};
  pd.slice(1).forEach(function(r){
    var p = _rObj(pd[0], r);
    if (!apm[p.StudentID]) apm[p.StudentID] = {};
    apm[p.StudentID][p.UnitID] = p;
  });
  var um = {};
  un.forEach(function(u){ um[u.UnitID] = u; });
  function pl(sid){ var m=apm[sid]||{}; return Object.keys(m).map(function(k){return m[k];}); }
  function pm(sid){ return apm[sid]||{}; }

  var tC=0,tP=0,tX=0,tH=0;
  stu.forEach(function(s){
    pl(s.StudentID).forEach(function(p){
      if(p.Status==='complete')tC++;
      if(p.Status==='awaiting_review')tP++;
      if(p.Status==='corrections')tX++;
      if(p.HomeworkSubmittedAt&&p.HomeworkSubmittedAt!=='')tH++;
    });
  });

  var si={complete:'✅',available:'📖',in_progress:'✏',awaiting_review:'🟡',corrections:'❌',locked:'🔒'};
  var sb={complete:'#dcfce7',available:'#dbeafe',in_progress:'#dbeafe',awaiting_review:'#fef9c3',corrections:'#fee2e2',locked:'#f1f5f9'};

  // Roster rows
  var rH = stu.map(function(s){
    var p2=pl(s.StudentID);
    var dn=p2.filter(function(p){return p.Status==='complete';}).length;
    var av=p2.filter(function(p){return p.Status==='available';}).length;
    var ls=p2.reduce(function(l,p){if(!p.HomeworkSubmittedAt)return l;var d=new Date(p.HomeworkSubmittedAt);return(!l||d>l)?d:l;},null);
    var dy=ls?Math.floor((now-ls)/86400000):null;
    var dl=dy===null?'—':dy===0?'<b style="color:#15803d">Today</b>':dy<=7?'<b style="color:#1d4ed8">'+dy+'d ago</b>':dy<=14?'<b style="color:#854d0e">'+dy+'d ago</b>':'<b style="color:#dc2626">'+dy+'d ago ⚠</b>';
    return '<tr>'+
      '<td><b>'+s.StudentName+'</b></td>'+
      '<td style="color:#94a3b8;font-size:.78rem">'+s.StudentID+'</td>'+
      '<td style="font-size:.78rem">'+s.StudentEmail+'</td>'+
      '<td style="font-size:.78rem">'+s.ParentEmail+'</td>'+
      '<td>'+dn+' ✅ / '+av+' 📖</td>'+
      '<td>'+dl+'</td>'+
      '<td>'+
        '<a href="mailto:'+s.StudentEmail+'" class="bs be">Email</a> '+
        '<a href="mailto:'+s.ParentEmail+'" class="bs bg">Parent</a> '+
        '<button class="bs br" onclick="doRemove(\''+s.StudentID+'\',\''+s.StudentName+'\')">✕ Remove</button>'+
      '</td></tr>';
  }).join('');

  // Attention
  var aH='';
  stu.forEach(function(s){
    pl(s.StudentID).filter(function(p){return p.Status==='corrections';}).forEach(function(p){
      var u=um[p.UnitID],d=p.ParentReviewedAt?Math.floor((now-new Date(p.ParentReviewedAt))/86400000):'?';
      aH+='<tr style="background:#fff7ed"><td>⚠ <b>'+s.StudentName+'</b></td><td>Corrections not resubmitted</td><td>'+(u?u.UnitName:p.UnitID)+'</td><td>'+d+'d</td><td><a href="mailto:'+s.StudentEmail+'" class="bs bg">📧 Nudge</a></td></tr>';
    });
    pl(s.StudentID).filter(function(p){return p.Status==='awaiting_review'&&p.HomeworkSubmittedAt;}).forEach(function(p){
      var d=Math.floor((now-new Date(p.HomeworkSubmittedAt))/86400000); if(d<3) return;
      var u=um[p.UnitID];
      aH+='<tr style="background:#fef9c3"><td>🕐 <b>'+s.StudentName+'</b></td><td>Parent review overdue</td><td>'+(u?u.UnitName:p.UnitID)+'</td><td>'+d+'d</td><td><a href="mailto:'+s.ParentEmail+'" class="bs" style="background:#1d4ed8;color:white">📧 Remind</a></td></tr>';
    });
  });
  if(!aH) aH='<tr><td colspan="5" style="text-align:center;color:#94a3b8;padding:12px">✅ No issues</td></tr>';

  // Pending
  var pH='';
  stu.forEach(function(s){
    pl(s.StudentID).filter(function(p){return p.Status==='awaiting_review';}).forEach(function(p){
      var u=um[p.UnitID];
      pH+='<tr>'+
        '<td><b>'+s.StudentName+'</b></td>'+
        '<td>'+p.UnitID+'</td>'+
        '<td>'+(u?u.UnitName:'')+'</td>'+
        '<td>'+_fd(p.HomeworkSubmittedAt)+'</td>'+
        '<td>'+(p.HomeworkDriveURL?'<a href="'+p.HomeworkDriveURL+'" target="_blank" class="bs be">View</a>':'—')+'</td>'+
        '<td><button class="bs bg" onclick="doUnlock(\''+s.StudentID+'\',\''+p.UnitID+'\',\''+s.StudentName+'\')">✅ Approve &amp; Unlock</button></td>'+
        '</tr>';
    });
  });
  if(!pH) pH='<tr><td colspan="6" style="text-align:center;color:#94a3b8;padding:12px">No pending submissions</td></tr>';

  // Progress grid
  var gH='<th class="sc">Student</th>'+un.map(function(u){return '<th class="uth" title="'+u.UnitName+'">'+u.UnitID+'</th>';}).join('');
  var gR=stu.map(function(s){
    var cells=un.map(function(u){
      var p2=pm(s.StudentID)[u.UnitID]||{Status:'locked'}, st=p2.Status||'locked';
      var btn=(st==='locked'||st==='corrections')?'<button class="ub" onclick="doUnlock(\''+s.StudentID+'\',\''+u.UnitID+'\',\''+s.StudentName+'\')">unlock</button>':'';
      return '<td style="background:'+(sb[st]||'#f1f5f9')+';text-align:center" title="'+s.StudentName+' — '+st+'">'+(si[st]||'🔒')+'<br>'+btn+'</td>';
    }).join('');
    return '<tr><td class="sc"><b>'+s.StudentName+'</b></td>'+cells+'</tr>';
  }).join('');

  // HW
  var hH='',hC=0;
  stu.forEach(function(s){
    pl(s.StudentID).filter(function(p){return p.HomeworkSubmittedAt&&p.HomeworkSubmittedAt!=='';}).forEach(function(p){
      var u=um[p.UnitID];
      var badge={complete:'<span class="hb hc">✅ Complete</span>',awaiting_review:'<span class="hb hp">🟡 Awaiting</span>',corrections:'<span class="hb hx">❌ Corrections</span>'}[p.Status]||p.Status;
      var pd2=p.ParentDecision?(p.ParentDecision==='approved'?'✅ Approved':'❌ Corrections'):'⏳ Pending';
      hH+='<tr><td><b>'+s.StudentName+'</b></td><td style="font-family:monospace;font-size:.78rem">'+p.UnitID+'</td><td>'+(u?u.UnitName:'')+'</td><td>'+_fd(p.HomeworkSubmittedAt)+'</td><td>'+badge+'</td><td>'+pd2+'</td><td>'+(p.ParentComments?'<span title="'+p.ParentComments+'" style="cursor:help;color:#64748b">💬 '+p.ParentComments.substring(0,30)+'</span>':'—')+'</td><td>'+(p.HomeworkDriveURL?'<a href="'+p.HomeworkDriveURL+'" target="_blank" class="bs be">View</a>':'—')+'</td></tr>';
      hC++;
    });
  });
  if(!hH) hH='<tr><td colspan="8" style="text-align:center;color:#94a3b8;padding:12px">No homework yet</td></tr>';

  // Email log
  var eH=el.slice(0,20).map(function(r){
    return '<tr><td>'+_fd(r[0])+'</td><td>'+r[1]+'</td><td>'+r[2]+'</td><td>'+r[3]+'</td><td>'+r[4]+'</td><td class="'+(r[6]==='sent'?'st2':'er')+'">'+r[6]+'</td></tr>';
  }).join('');
  if(!eH) eH='<tr><td colspan="6" style="text-align:center;color:#94a3b8;padding:12px">No emails yet</td></tr>';

  var html =
    '<!DOCTYPE html><html lang="en"><head>' +
    '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Admin — Physics Foundations</title>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com">' +
    '<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;700;800;900&display=swap" rel="stylesheet">' +
    _css() +
    '</head><body>' +
    '<div id="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:9999;display:none;align-items:center;justify-content:center"><div style="background:white;border-radius:16px;padding:40px 48px;text-align:center;max-width:400px"><div id="ov_icon" style="font-size:2.5rem;margin-bottom:12px">⏳</div><div id="ov_msg" style="font-size:1rem;font-weight:700;color:#0f172a;line-height:1.6"></div></div></div>' +
    '<header><div class="hi">' +
      '<div><div class="lbl">Physics Foundations Admin</div><h1>Dashboard</h1></div>' +
      '<div style="display:flex;gap:10px;align-items:center">' +
        '<span style="color:rgba(255,255,255,.35);font-size:.7rem">'+_fd(now)+'</span>' +
        '<button class="bl" onclick="doRefresh()">🔄 Refresh</button>' +
        '<a href="https://misra-ravi.github.io/physics-foundation/class-schedule.html" target="_blank" class="bl">📅 Schedule</a>' +
        '<a href="https://misra-ravi.github.io/physics-foundation/" target="_blank" class="bl">← Site</a>' +
      '</div>' +
    '</div></header>' +
    '<div class="sb">' +
      '<div class="st"><div class="sn">'+stu.length+'</div><div class="sl">Students</div></div>' +
      '<div class="st"><div class="sn">'+tC+'</div><div class="sl">Complete</div></div>' +
      '<div class="st"><div class="sn">'+tH+'</div><div class="sl">HW Submitted</div></div>' +
      '<div class="st"><div class="sn">'+tP+'</div><div class="sl">Awaiting Review</div></div>' +
      '<div class="st"><div class="sn">'+tX+'</div><div class="sl">Corrections</div></div>' +
    '</div>' +
    '<main>' +
      '<section><h2>⚠ Needs Attention</h2>' +
      '<table><thead><tr><th>Student</th><th>Issue</th><th>Unit</th><th>Days</th><th>Action</th></tr></thead>' +
      '<tbody>'+aH+'</tbody></table></section>' +

      '<section><h2>Add New Student</h2>' +
      '<div class="fg">' +
        '<input id="a_sid" placeholder="Student ID (e.g. s004)">' +
        '<input id="a_nm"  placeholder="Full Name">' +
        '<input id="a_em"  placeholder="Student Gmail" type="email">' +
        '<input id="a_pe"  placeholder="Parent Gmail" type="email">' +
        '<input id="a_pn"  placeholder="Parent Name (e.g. Mr Khan)">' +
        '<button class="bp" onclick="doAdd()">+ Enrol Student</button>' +
      '</div></section>' +

      '<section><h2>Student Roster</h2>' +
      '<table><thead><tr><th>Name</th><th>ID</th><th>Student Email</th><th>Parent Email</th><th>Progress</th><th>Last Activity</th><th>Actions</th></tr></thead>' +
      '<tbody>'+rH+'</tbody></table></section>' +

      '<section><h2>Pending Reviews ('+tP+')</h2>' +
      '<table><thead><tr><th>Student</th><th>Unit ID</th><th>Unit</th><th>Submitted</th><th>Homework</th><th>Action</th></tr></thead>' +
      '<tbody>'+pH+'</tbody></table></section>' +

      '<section><h2>Progress Grid</h2>' +
      '<p class="hint">Click "unlock" on any 🔒 cell to unlock that unit and email student + parent.</p>' +
      '<div class="gs"><table class="gt"><thead><tr>'+gH+'</tr></thead><tbody>'+gR+'</tbody></table></div></section>' +

      '<section><h2>Homework Submissions ('+hC+')</h2>' +
      '<table><thead><tr><th>Student</th><th>Unit ID</th><th>Unit</th><th>Submitted</th><th>Status</th><th>Parent Decision</th><th>Comments</th><th>File</th></tr></thead>' +
      '<tbody>'+hH+'</tbody></table></section>' +

      '<section><h2>Email Log (last 20)</h2>' +
      '<table><thead><tr><th>Time</th><th>Type</th><th>Student</th><th>Unit</th><th>Recipient</th><th>Status</th></tr></thead>' +
      '<tbody>'+eH+'</tbody></table></section>' +
    '</main>' +
    '<footer>Physics Foundations by Ravi · Admin Dashboard</footer>' +

    '<script>' +
    'function showOverlay(icon,msg){var o=document.getElementById("overlay");document.getElementById("ov_icon").textContent=icon;document.getElementById("ov_msg").innerHTML=msg;o.style.display="flex";}' +
    'function hideOverlay(){document.getElementById("overlay").style.display="none";}' +
    'function doRefresh(){' +
      'showOverlay("⏳","Refreshing dashboard...");' +
      'google.script.run.withSuccessHandler(function(html){document.open();document.write(html);document.close();}).withFailureHandler(fail).getAdminHtml();' +
    '}' +
    'function done(r){' +
      'hideOverlay();' +
      'if(r&&r.ok===false){alert("Error: "+r.msg);}' +
      'else{' +
        'alert(r&&r.msg?r.msg:"Done.");' +
        'google.script.run.withSuccessHandler(function(html){' +
          'document.open();document.write(html);document.close();' +
        '}).withFailureHandler(function(e){' +
          'alert("Dashboard refresh failed: "+e+". Please close and reopen the admin portal.");' +
        '}).getAdminHtml();' +
      '}' +
    '}' +
    'function fail(e){hideOverlay();alert("Error: "+(e&&e.message?e.message:e));}' +
    'function doAdd(){' +
      'var sid=document.getElementById("a_sid").value.trim();' +
      'var nm=document.getElementById("a_nm").value.trim();' +
      'var em=document.getElementById("a_em").value.trim();' +
      'var pe=document.getElementById("a_pe").value.trim();' +
      'var pn=document.getElementById("a_pn").value.trim();' +
      'if(!sid||!nm||!em||!pe||!pn){alert("Please fill in all 5 fields.");return;}' +
      'showOverlay("⏳","Enrolling "+nm+"...");' +
      'google.script.run.withSuccessHandler(done).withFailureHandler(fail).adminAddStudent({sid:sid,name:nm,email:em,parentEmail:pe,parentName:pn});' +
    '}' +
    'function doUnlock(sid,uid,name){' +
      'if(!confirm("Unlock unit "+uid+" for "+name+"?"))return;' +
      'showOverlay("⏳","Unlocking unit for "+name+"...");' +
      'google.script.run.withSuccessHandler(done).withFailureHandler(fail).adminUnlock(sid,uid);' +
    '}' +
    'function doRemove(sid,name){' +
      'if(!confirm("Remove "+name+"? Their progress history will be kept."))return;' +
      'showOverlay("⏳","Removing "+name+"...");' +
      'google.script.run.withSuccessHandler(done).withFailureHandler(fail).adminRemove(sid);' +
    '}' +
    '</script>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Admin — Physics Foundations')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _getEmailLog(ss) {
  var sh = ss.getSheetByName('EmailLog');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getDataRange().getValues().slice(1).reverse();
}

function _fd(d) {
  if (!d) return '—';
  try { return new Date(d).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'}); }
  catch(e){ return String(d); }
}

function _css() {
  return '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:Nunito,sans-serif;background:#f0f4f8;color:#0f172a}' +
    'header{background:linear-gradient(135deg,#0f172a,#1e3a5f);color:white;padding:16px 24px}' +
    '.hi{max-width:1400px;margin:0 auto;display:flex;align-items:center;justify-content:space-between}' +
    '.lbl{font-size:.68rem;text-transform:uppercase;letter-spacing:2px;opacity:.5;margin-bottom:3px}' +
    'h1{font-size:1.35rem;font-weight:900}' +
    '.bl{color:rgba(255,255,255,.75);font-size:.76rem;font-weight:700;text-decoration:none;background:rgba(255,255,255,.1);padding:5px 11px;border-radius:6px;border:none;cursor:pointer}' +
    '.sb{background:#1e293b;padding:10px 24px;display:flex;gap:20px;flex-wrap:wrap}' +
    '.st{text-align:center}.sn{font-size:1.35rem;font-weight:900;color:white}.sl{font-size:.62rem;text-transform:uppercase;letter-spacing:1px;color:#94a3b8}' +
    'main{max-width:1400px;margin:0 auto;padding:18px 18px 80px}' +
    'section{background:white;border-radius:12px;padding:18px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,.07)}' +
    'h2{font-size:.9rem;font-weight:800;margin-bottom:11px}' +
    '.hint{font-size:.74rem;color:#64748b;margin-bottom:9px}' +
    'table{width:100%;border-collapse:collapse}' +
    'thead th{background:#f1f5f9;font-size:.66rem;font-weight:800;text-transform:uppercase;letter-spacing:.5px;padding:7px 9px;text-align:left;border-bottom:2px solid #e2e8f0}' +
    'td{padding:6px 9px;font-size:.81rem;border-bottom:1px solid #f1f5f9;vertical-align:middle}' +
    'tr:hover td{background:#fafafa}' +
    '.st2{color:#15803d;font-weight:700}.er{color:#dc2626;font-weight:700}' +
    '.fg{display:grid;grid-template-columns:repeat(auto-fill,minmax(165px,1fr));gap:7px;align-items:end}' +
    '.fg input{border:2px solid #e2e8f0;border-radius:6px;padding:7px 9px;font-family:inherit;font-size:.81rem;width:100%}' +
    '.fg input:focus{outline:none;border-color:#3b82f6}' +
    '.bp{background:#1d4ed8;color:white;border:none;border-radius:6px;padding:8px 14px;font-weight:800;font-size:.81rem;cursor:pointer;font-family:inherit;width:100%}' +
    '.bs{display:inline-block;padding:3px 8px;border-radius:5px;font-size:.68rem;font-weight:800;text-decoration:none;white-space:nowrap;border:none;cursor:pointer;font-family:inherit}' +
    '.be{background:#eff6ff;color:#1d4ed8}' +
    '.bg{background:#22c55e;color:white}' +
    '.br{background:#fee2e2;color:#dc2626}' +
    '.ub{display:block;font-size:.6rem;color:#1d4ed8;background:none;border:none;cursor:pointer;font-weight:800;font-family:inherit;padding:0;margin-top:2px}' +
    '.ub:hover{text-decoration:underline}' +
    '.hb{display:inline-block;padding:2px 7px;border-radius:14px;font-size:.68rem;font-weight:800}' +
    '.hc{background:#dcfce7;color:#15803d}.hp{background:#fef9c3;color:#854d0e}.hx{background:#fee2e2;color:#dc2626}' +
    '.gs{overflow-x:auto;max-height:440px;overflow-y:auto}' +
    '.gt{border-collapse:collapse;font-size:.74rem}' +
    '.gt th,.gt td{border:1px solid #e2e8f0;padding:4px 5px;text-align:center;white-space:nowrap}' +
    '.uth{font-size:.6rem;font-weight:700;background:#f1f5f9;writing-mode:vertical-rl;transform:rotate(180deg);padding:6px 3px}' +
    '.sc{position:sticky;left:0;background:white;z-index:1;font-weight:700;min-width:110px;text-align:left!important;border-right:2px solid #e2e8f0}' +
    'footer{background:#0f172a;color:rgba(255,255,255,.3);text-align:center;padding:14px;font-size:.74rem}' +
    '</style>';
}
