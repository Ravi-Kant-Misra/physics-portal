// ═══════════════════════════════════════════════════════════════════════════════
// IB PHYSICS PORTAL — BOOTSTRAP
// ═══════════════════════════════════════════════════════════════════════════════
//
// HOW TO USE
// ──────────
// 1. Go to https://script.google.com → click "New project"
// 2. Delete everything in the editor
// 3. Paste this entire file
// 4. Click Save (floppy disk icon), name the project "IB Physics Portal"
// 5. Make sure the function selector at the top shows "bootstrapPortal"
// 6. Click ▶ Run
// 7. Click "Review permissions" → choose your Google account → Allow
// 8. Wait ~60 seconds — check the Execution Log at the bottom for the output
// 9. Copy the URLs that appear in the log — you will need them for Step 11
// 10. Follow the AFTER BOOTSTRAP instructions printed in the log
//
// ═══════════════════════════════════════════════════════════════════════════════

var ADMIN_EMAIL  = 'misra.ravikant@gmail.com';
var PORTAL_NAME  = 'IB Physics Portal';
var BASE_URL     = 'https://ravi-kant-misra.github.io/physics-portal/';

// ─── MAIN ENTRY POINT ─────────────────────────────────────────────────────────
function bootstrapPortal() {
  Logger.log('═══════════════════════════════════════');
  Logger.log(' IB Physics Portal — Bootstrap');
  Logger.log('═══════════════════════════════════════');

  // 1. Drive folders
  Logger.log('\n📁 Creating Drive folders...');
  var rootFolder    = _getOrCreateFolder(null,         PORTAL_NAME);
  var submissionsF  = _getOrCreateFolder(rootFolder,   'Student Submissions');
  var answerKeysF   = _getOrCreateFolder(rootFolder,   'Answer Keys');
  Logger.log('   Root folder:    ' + rootFolder.getUrl());
  Logger.log('   Submissions:    ' + submissionsF.getUrl());
  Logger.log('   Answer Keys:    ' + answerKeysF.getUrl());

  // 2. Google Sheet
  Logger.log('\n📊 Creating Portal Tracker sheet...');
  var ss = _getOrCreateSheet(rootFolder, 'IB Physics — Portal Tracker');
  var sheetId = ss.getId();
  Logger.log('   Sheet URL: ' + ss.getUrl());

  // 3. Sheet tabs
  Logger.log('\n📋 Setting up sheet tabs...');
  _setupRoster(ss);
  _setupUnits(ss);
  _setupProgress(ss);
  _setupEmailLog(ss);
  SpreadsheetApp.flush();
  Logger.log('   Tabs created: Roster, Units, Progress, EmailLog');

  // 4. Seed units
  Logger.log('\n🔢 Seeding 130 unit rows...');
  _seedUnits(ss);
  Logger.log('   Done.');

  // 5. Seed progress
  Logger.log('\n📈 Seeding progress rows for sample students...');
  _seedProgress(ss);
  Logger.log('   Done.');

  // 6. Form A — Homework Submission
  Logger.log('\n📝 Creating Form A: Homework Submission...');
  var formA = _createHomeworkForm(rootFolder, ss);
  var formAId = formA.getId();
  Logger.log('   Form A URL: ' + formA.getPublishedUrl());

  // 7. Form B — Parent Approval
  Logger.log('\n📝 Creating Form B: Parent Approval...');
  var formB = _createApprovalForm(rootFolder);
  var formBId = formB.getId();
  Logger.log('   Form B URL: ' + formB.getPublishedUrl());

  // 8. Write final config into this script's Properties so the triggers can read them
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    SHEET_ID:              sheetId,
    HOMEWORK_FORM_ID:      formAId,
    APPROVAL_FORM_ID:      formBId,
    SUBMISSIONS_FOLDER_ID: submissionsF.getId(),
    ANSWER_KEYS_FOLDER_ID: answerKeysF.getId(),
  });
  Logger.log('\n✅ Config saved to Script Properties.');

  // 9. Install form triggers
  Logger.log('\n⚡ Installing form submit triggers...');
  _installTriggers(ss, formA, formB);
  Logger.log('   Triggers installed.');

  Logger.log('\n═══════════════════════════════════════');
  Logger.log(' BOOTSTRAP COMPLETE');
  Logger.log('═══════════════════════════════════════');
  Logger.log('\nSheet:  ' + ss.getUrl());
  Logger.log('Form A: ' + formA.getPublishedUrl());
  Logger.log('Form B: ' + formB.getPublishedUrl());
  Logger.log('\nNEXT STEP:');
  Logger.log('Deploy this project as a Web App:');
  Logger.log('  Deploy → New Deployment → Web App');
  Logger.log('  Execute as: User accessing the web app');
  Logger.log('  Who has access: Anyone with a Google Account');
  Logger.log('  → Share that URL with your students.');
}

// ─── DRIVE ────────────────────────────────────────────────────────────────────
function _getOrCreateFolder(parent, name) {
  var search = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (search.hasNext()) return search.next();
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

function _getOrCreateSheet(folder, name) {
  var files = folder.getFilesByName(name);
  if (files.hasNext()) {
    return SpreadsheetApp.openById(files.next().getId());
  }
  var ss = SpreadsheetApp.create(name);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  return ss;
}

// ─── SHEET TABS ───────────────────────────────────────────────────────────────
function _setupRoster(ss) {
  var sh = ss.getSheetByName('Roster') || ss.insertSheet('Roster');
  sh.clearContents();
  var h = ['StudentID','StudentName','StudentEmail','ParentEmail','ParentName','StartDate','Active'];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold');
  sh.getRange(2,1,2,h.length).setValues([
    ['s001','Alice Chen',  'alice.chen.student@gmail.com',  'parent.chen@gmail.com',  'Mr & Mrs Chen',  new Date(), true],
    ['s002','Bob Patel',   'bob.patel.student@gmail.com',   'parent.patel@gmail.com', 'Mr & Mrs Patel', new Date(), true],
  ]);
  sh.setFrozenRows(1);
}

function _setupUnits(ss) {
  var sh = ss.getSheetByName('Units') || ss.insertSheet('Units');
  sh.clearContents();
  var h = ['UnitID','UnitOrder','TopicNum','TopicName','SectionNum','SectionName',
           'UnitNum','UnitName','Type','LessonURL','ClassworkURL','HomeworkURL',
           'TopicSlug','SectionSlug','UnitSlug','AnswerKeyDriveID','PrevUnitID'];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function _setupProgress(ss) {
  var sh = ss.getSheetByName('Progress') || ss.insertSheet('Progress');
  sh.clearContents();
  var h = ['ProgressID','StudentID','UnitID','Status','LessonOpened',
           'HomeworkSubmittedAt','HomeworkDriveURL','SubmissionToken',
           'ParentReviewedAt','ParentDecision','ParentComments','UnlockedAt'];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function _setupEmailLog(ss) {
  var sh = ss.getSheetByName('EmailLog') || ss.insertSheet('EmailLog');
  sh.clearContents();
  var h = ['Timestamp','Type','StudentID','UnitID','Recipient','Subject','Status'];
  sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold');
  sh.setFrozenRows(1);
  // Remove the default blank sheet if it exists
  var def = ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);
}

// ─── SEED UNITS ───────────────────────────────────────────────────────────────
function _seedUnits(ss) {
  var sh   = ss.getSheetByName('Units');
  var base = BASE_URL;
  var rows = [];
  var order = 1;

  function makeRow(id, ord, tNum, tName, sNum, sName, uNum, uName, type, tSlug, sSlug, uSlug, prev) {
    var lesson    = base + tSlug + '/' + sSlug + '/' + uSlug + '/lesson.html';
    var classwork = (type==='full') ? base + tSlug + '/' + sSlug + '/' + uSlug + '/classwork.html' : '';
    var homework  = (type==='full') ? base + tSlug + '/' + sSlug + '/' + uSlug + '/homework.html'  : '';
    return [id,ord,tNum,tName,sNum,sName,uNum,uName,type,lesson,classwork,homework,tSlug,sSlug,uSlug,'',prev];
  }

  // 0.1 Background Math
  var bg = [
    'Significant Figures','Calculating Significant Figures','Variables, Units & Prefixes',
    'Derived Units','Unit Conversion','Proportional Relationships','Uncertainty Notation',
    'Absolute, Fractional & Percentage Uncertainty','Propagation: Add & Subtract',
    'Propagation: Multiply & Divide','Propagation: Powers','Propagation in Physics',
    'Measurement Uncertainty','Data Tables','Logger Pro','Line of Best Fit',
    'Vectors & Scalars','Labelling & Adding Vectors','Recording Angles','Finding Angles',
    'X & Y Components','Finding X & Y with Trig','Complex Vectors',
    'Review: Trig Sides','Review: Trig','Inverse Trig','3D Vectors'
  ];
  bg.forEach(function(name,i){
    var n=i+1, uSlug='u'+_pad(n), prev=n>1?'0.1.'+_pad(n-1):'';
    rows.push(makeRow('0.1.'+_pad(n),order++,0,'Appendix','0.1','Background Math',n,name,'full','t0-appendix','s0.1-background-math',uSlug,prev));
  });

  // 0.2 Stories
  var stories = [
    'Introduction to Physics Stories','Galileo & the Pendulum','Newton & the Apple',
    'Faraday & the Field','Light: Wave or Particle?','Marie & Pierre Curie',
    'Einstein & Relativity','Einstein\'s Chauffeur','Schrödinger\'s Cat',
    'Lise Meitner','Satyendra Bose'
  ];
  stories.forEach(function(name,i){
    var uSlug='u'+_pad(i), prev=i>0?'0.2.'+_pad(i-1):'';
    rows.push(makeRow('0.2.'+_pad(i),order++,0,'Appendix','0.2','Stories of Science',i,name,'story','t0-appendix','s0.2-stories',uSlug,prev));
  });

  // 1.1 Kinematics
  var kin = [
    'Distance & Displacement','Speed & Velocity','Acceleration',
    'Average vs Instantaneous Velocity','Position-Time Graphs',
    'Velocity-Time Graphs I','Velocity-Time Graphs II','PT to VT Conversion',
    'Acceleration-Time Graphs','Kinematics Transition','Kinematic Equations',
    'Kinematics & Geometry','Free Fall','Projectile Motion','Air Resistance','Motion Maps'
  ];
  kin.forEach(function(name,i){
    var n=i+1, uSlug='u'+_pad(n), prev=n>1?'1.1.'+_pad(n-1):'';
    rows.push(makeRow('1.1.'+_pad(n),order++,1,'Space, Time and Motion','1.1','Kinematics',n,name,'full','t1-space-time-motion','s1.1-kinematics',uSlug,prev));
  });

  // Remaining topics — stub rows (no URLs yet)
  var stubs = [
    [1,'Space, Time and Motion','1.2','Forces & Momentum',17,'t1-space-time-motion','s1.2-forces-momentum'],
    [1,'Space, Time and Motion','1.3','Work, Energy & Power',12,'t1-space-time-motion','s1.3-work-energy-power'],
    [2,'The Particulate Nature of Matter','2.1','Thermal Energy Transfers',7,'t2-particulate-matter','s2.1-thermal-energy'],
    [2,'The Particulate Nature of Matter','2.2','Greenhouse Effect',1,'t2-particulate-matter','s2.2-greenhouse-effect'],
    [2,'The Particulate Nature of Matter','2.3','Gas Laws',4,'t2-particulate-matter','s2.3-gas-laws'],
    [2,'The Particulate Nature of Matter','2.4','Current & Circuits',10,'t2-particulate-matter','s2.4-current-circuits'],
    [3,'Wave Behavior','3.1','Simple Harmonic Motion',2,'t3-wave-behavior','s3.1-simple-harmonic'],
    [3,'Wave Behavior','3.2','Wave Model',4,'t3-wave-behavior','s3.2-wave-model'],
    [3,'Wave Behavior','3.3','Wave Phenomena',11,'t3-wave-behavior','s3.3-wave-phenomena'],
    [3,'Wave Behavior','3.4','Standing Waves & Resonance',3,'t3-wave-behavior','s3.4-standing-waves'],
    [4,'Fields','4.1','Gravitational Fields',1,'t4-fields','s4.1-gravitational'],
    [4,'Fields','4.2','Electric & Magnetic Fields',8,'t4-fields','s4.2-electric-magnetic'],
    [5,'Nuclear and Quantum Physics','5.1','Structure of the Atom',6,'t5-nuclear-quantum','s5.1-structure-atom'],
    [5,'Nuclear and Quantum Physics','5.2','Radioactive Decay',6,'t5-nuclear-quantum','s5.2-radioactive-decay'],
    [5,'Nuclear and Quantum Physics','5.3','Fission',1,'t5-nuclear-quantum','s5.3-fission'],
    [5,'Nuclear and Quantum Physics','5.4','Fusion & Stars',4,'t5-nuclear-quantum','s5.4-fusion-stars'],
  ];
  stubs.forEach(function(st){
    var tNum=st[0],tName=st[1],sNum=st[2],sName=st[3],count=st[4],tSlug=st[5],sSlug=st[6];
    for(var n=1;n<=count;n++){
      var id=sNum+'.'+_pad(n), prev=n>1?sNum+'.'+_pad(n-1):'';
      rows.push([id,order++,tNum,tName,sNum,sName,n,sName+' — Unit '+n,'full','','','',tSlug,sSlug,'u'+_pad(n),'',prev]);
    }
  });

  sh.getRange(2,1,rows.length,17).setValues(rows);
  Logger.log('   ' + rows.length + ' units seeded.');
}

function _pad(n){ return String(n).padStart(2,'0'); }

// ─── SEED PROGRESS ────────────────────────────────────────────────────────────
function _seedProgress(ss) {
  var rosterSh   = ss.getSheetByName('Roster');
  var unitsSh    = ss.getSheetByName('Units');
  var progressSh = ss.getSheetByName('Progress');

  var rData = rosterSh.getDataRange().getValues();
  var uData = unitsSh.getDataRange().getValues();

  var students = rData.slice(1).filter(function(r){ return r[6]; }); // Active=true
  var units    = uData.slice(1);

  if(progressSh.getLastRow() > 1){
    progressSh.deleteRows(2, progressSh.getLastRow()-1);
  }

  var rows = [];
  students.forEach(function(s){
    var sid = s[0];
    units.forEach(function(u, idx){
      var uid    = u[0];
      var status = (idx === 0) ? 'available' : 'locked';
      rows.push([sid+'_'+uid, sid, uid, status, '','','','','','','','']);
    });
  });

  if(rows.length > 0){
    progressSh.getRange(2,1,rows.length,12).setValues(rows);
  }
  Logger.log('   ' + rows.length + ' progress rows seeded (' + students.length + ' students × ' + units.length + ' units).');
}

// ─── FORM A: HOMEWORK SUBMISSION ──────────────────────────────────────────────
function _createHomeworkForm(folder, ss) {
  var form = FormApp.create('IB Physics — Homework Submission');
  form.setDescription('Submit your completed homework for parent review.');
  form.setCollectEmail(true);
  form.setLimitOneResponsePerUser(false);

  // Build student name list from Roster
  var rData    = ss.getSheetByName('Roster').getDataRange().getValues();
  var students = rData.slice(1).map(function(r){ return r[1]; }).filter(Boolean);

  // Q1 — Student name
  var q1 = form.addListItem();
  q1.setTitle('Your name').setRequired(true);
  q1.setChoiceValues(students.length > 0 ? students : ['Alice Chen','Bob Patel']);

  // Q2 — Unit (top 54 real units — Background Math + Stories + Kinematics)
  var uData     = ss.getSheetByName('Units').getDataRange().getValues();
  var unitLabels = uData.slice(1)
    .filter(function(u){ return u[9] !== ''; }) // only units with a LessonURL
    .map(function(u){ return u[0] + ' — ' + u[7]; }); // "1.1.01 — Distance & Displacement"

  var q2 = form.addListItem();
  q2.setTitle('Unit completed').setRequired(true);
  q2.setChoiceValues(unitLabels.length > 0 ? unitLabels : ['0.1.01 — Significant Figures']);

  // Q3 — Homework link (file upload not available via Apps Script API)
  var q3p = form.addParagraphTextItem();
  q3p.setTitle('Paste a link to your homework (Google Drive share link or photo link)').setRequired(true);

  // Q4 — Confirmation
  var q4 = form.addCheckboxItem();
  q4.setTitle('Confirmation').setRequired(true);
  q4.setChoices([q4.createChoice('I completed the lesson and classwork before submitting this homework')]);

  // Q5 — Notes
  form.addParagraphTextItem().setTitle('Anything you found difficult? (optional)').setRequired(false);

  // Move form file into portal folder
  var formFile = DriveApp.getFileById(form.getId());
  formFile.moveTo(folder);

  // Link responses to the sheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  return form;
}

// ─── FORM B: PARENT APPROVAL ──────────────────────────────────────────────────
function _createApprovalForm(folder) {
  var form = FormApp.create('IB Physics — Parent Homework Review');
  form.setDescription('Review your child\'s homework and record your decision below.');
  form.setCollectEmail(false);

  // Hidden pre-filled fields (parents never see these — they arrive via pre-filled URL)
  form.addTextItem().setTitle('studentId').setRequired(false);
  form.addTextItem().setTitle('unitId').setRequired(false);
  form.addTextItem().setTitle('token').setRequired(false);
  form.addTextItem().setTitle('studentName').setRequired(false);
  form.addTextItem().setTitle('unitName').setRequired(false);

  // Decision
  var dec = form.addMultipleChoiceItem();
  dec.setTitle('Your decision').setRequired(true);
  dec.setChoices([
    dec.createChoice('Homework is complete — approve and unlock next chapter'),
    dec.createChoice('Needs corrections — send back to student'),
  ]);

  // Comments
  form.addParagraphTextItem()
    .setTitle('Comments for your child (required if requesting corrections)')
    .setRequired(false);

  // Parent name
  form.addTextItem().setTitle('Your name').setRequired(true);

  // Move into portal folder
  DriveApp.getFileById(form.getId()).moveTo(folder);

  return form;
}

// ─── TRIGGERS ─────────────────────────────────────────────────────────────────
function _installTriggers(ss, formA, formB) {
  // Remove any existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('onHomeworkSubmit')
    .forForm(formA)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('onParentApproval')
    .forForm(formB)
    .onFormSubmit()
    .create();

  Logger.log('   onHomeworkSubmit trigger → Form A');
  Logger.log('   onParentApproval trigger → Form B');
}

// ═══════════════════════════════════════════════════════════════════════════════
// CONFIG — reads from Script Properties (set by bootstrapPortal)
// ═══════════════════════════════════════════════════════════════════════════════
function _cfg() {
  var p = PropertiesService.getScriptProperties().getProperties();
  return {
    SHEET_ID:              p.SHEET_ID,
    HOMEWORK_FORM_ID:      p.HOMEWORK_FORM_ID,
    APPROVAL_FORM_ID:      p.APPROVAL_FORM_ID,
    SUBMISSIONS_FOLDER_ID: p.SUBMISSIONS_FOLDER_ID,
    ANSWER_KEYS_FOLDER_ID: p.ANSWER_KEYS_FOLDER_ID,
    BASE_URL:              BASE_URL,
    ADMIN_EMAIL:           ADMIN_EMAIL,
    PORTAL_NAME:           PORTAL_NAME,
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SHARED DATA HELPERS
// ═══════════════════════════════════════════════════════════════════════════════
function _ss()   { return SpreadsheetApp.openById(_cfg().SHEET_ID); }
function _rObj(h,r){ var o={}; h.forEach(function(k,i){o[k]=r[i];}); return o; }

function _students(ss){ var d=ss.getSheetByName('Roster').getDataRange().getValues(); return d.slice(1).map(function(r){return _rObj(d[0],r);}).filter(function(s){return s.Active;}); }
function _studentByName(ss,n){ return _students(ss).find(function(s){return s.StudentName===n;})||null; }
function _studentById(ss,id)  { return _students(ss).find(function(s){return s.StudentID===id;})||null; }
function _studentByEmail(ss,e){ return _students(ss).find(function(s){return s.StudentEmail===e;})||null; }

function _units(ss){ var d=ss.getSheetByName('Units').getDataRange().getValues(); return d.slice(1).map(function(r){return _rObj(d[0],r);}); }
function _unit(ss,id){ return _units(ss).find(function(u){return u.UnitID===id;})||null; }
function _nextUnit(ss,id){ var us=_units(ss),i=us.findIndex(function(u){return u.UnitID===id;}); return(i>=0&&i<us.length-1)?us[i+1]:null; }

function _progress(ss,sid){ var d=ss.getSheetByName('Progress').getDataRange().getValues(); return d.slice(1).map(function(r){return _rObj(d[0],r);}).filter(function(p){return p.StudentID===sid;}); }
function _prog(ss,sid,uid){ var d=ss.getSheetByName('Progress').getDataRange().getValues(),h=d[0],pid=sid+'_'+uid; for(var i=1;i<d.length;i++){if(d[i][0]===pid)return _rObj(h,d[i]);} return null; }

function _setProg(ss,sid,uid,fields){
  var sh=ss.getSheetByName('Progress'),d=sh.getDataRange().getValues(),h=d[0],pid=sid+'_'+uid;
  for(var i=1;i<d.length;i++){
    if(d[i][0]===pid){ Object.keys(fields).forEach(function(k){var c=h.indexOf(k);if(c>=0)sh.getRange(i+1,c+1).setValue(fields[k]);}); return; }
  }
  var row=h.map(function(k){return fields[k]!==undefined?fields[k]:'';});
  row[0]=pid; row[h.indexOf('StudentID')]=sid; row[h.indexOf('UnitID')]=uid;
  sh.appendRow(row);
}

function _logEmail(ss,type,sid,uid,to,subj,status){
  ss.getSheetByName('EmailLog').appendRow([new Date(),type,sid,uid,to,subj,status||'sent']);
}

function _prefilledUrl(formId,vals){
  var form=FormApp.openById(formId),items=form.getItems(),params=[];
  items.forEach(function(item){var t=item.getTitle();if(vals[t]!==undefined)params.push('entry.'+item.getId()+'='+encodeURIComponent(vals[t]));});
  return 'https://docs.google.com/forms/d/'+formId+'/viewform?'+params.join('&');
}

function _alertAdmin(msg){ GmailApp.sendEmail(ADMIN_EMAIL,'[IB Physics Portal] Error',msg,{name:PORTAL_NAME}); }
function _first(arr){ return(arr&&arr.length>0)?arr[0]:''; }
function _stripHtml(h){ return h.replace(/<[^>]+>/g,' ').replace(/\s+/g,' ').trim(); }

// ═══════════════════════════════════════════════════════════════════════════════
// TRIGGER A — HOMEWORK SUBMISSION
// ═══════════════════════════════════════════════════════════════════════════════
function onHomeworkSubmit(e) {
  var cfg = _cfg();
  var ss  = SpreadsheetApp.openById(cfg.SHEET_ID);
  var row = e.namedValues;

  var studentName = _first(row['Your name']);
  var unitLabel   = _first(row['Unit completed']);
  var fileUrl     = _first(row['Paste a link to your homework (Google Drive share link or photo link)']);
  var notes       = _first(row['Anything you found difficult? (optional)']) || '';
  var unitId      = unitLabel ? unitLabel.split(' — ')[0].trim() : '';

  var student = _studentByName(ss, studentName);
  var unit    = _unit(ss, unitId);
  if (!student || !unit) { _alertAdmin('onHomeworkSubmit: unknown student "'+studentName+'" or unit "'+unitId+'"'); return; }

  var prog = _prog(ss, student.StudentID, unitId);
  if (prog && ['awaiting_review','complete'].indexOf(prog.Status) >= 0) { Logger.log('Duplicate ignored'); return; }

  var token = Utilities.getUuid();
  _setProg(ss, student.StudentID, unitId, { Status:'awaiting_review', HomeworkSubmittedAt:new Date(), HomeworkDriveURL:fileUrl, SubmissionToken:token });

  var approvalUrl = _prefilledUrl(cfg.APPROVAL_FORM_ID, { studentId:student.StudentID, unitId:unitId, token:token, studentName:student.StudentName, unitName:unit.UnitName });

  var attachments = [];
  if (unit.AnswerKeyDriveID) {
    try { attachments.push(DriveApp.getFileById(unit.AnswerKeyDriveID).getBlob().setName('AnswerKey_'+unit.UnitName.replace(/[^a-zA-Z0-9]/g,'_')+'.pdf')); } catch(err) {}
  }

  var subj = '[IB Physics] '+student.StudentName+' submitted homework — '+unit.UnitName;
  var body = _parentEmail(student, unit, fileUrl, approvalUrl, notes);
  try {
    GmailApp.sendEmail(student.ParentEmail, subj, _stripHtml(body), { htmlBody:body, attachments:attachments, name:PORTAL_NAME, replyTo:ADMIN_EMAIL });
    _logEmail(ss,'homework_notification',student.StudentID,unitId,student.ParentEmail,subj,'sent');
  } catch(err) {
    _logEmail(ss,'homework_notification',student.StudentID,unitId,student.ParentEmail,subj,'error:'+err);
    _alertAdmin('Failed to email parent: '+err);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// TRIGGER B — PARENT APPROVAL
// ═══════════════════════════════════════════════════════════════════════════════
function onParentApproval(e) {
  var cfg = _cfg();
  var ss  = SpreadsheetApp.openById(cfg.SHEET_ID);
  var row = e.namedValues;

  var studentId  = _first(row['studentId']);
  var unitId     = _first(row['unitId']);
  var token      = _first(row['token']);
  var decision   = _first(row['Your decision']);
  var comments   = _first(row['Comments for your child (required if requesting corrections)']) || '';
  var parentName = _first(row['Your name']) || 'Parent';
  var approved   = decision.toLowerCase().indexOf('approve') >= 0;

  var student = _studentById(ss, studentId);
  var unit    = _unit(ss, unitId);
  if (!student || !unit) { _alertAdmin('onParentApproval: unknown student "'+studentId+'" or unit "'+unitId+'"'); return; }

  var prog = _prog(ss, studentId, unitId);
  if (!prog || prog.SubmissionToken !== token) { _alertAdmin('Invalid token for '+studentId+'/'+unitId); return; }
  if (prog.Status === 'complete') { Logger.log('Already approved'); return; }

  _setProg(ss, studentId, unitId, { Status:approved?'complete':'corrections', ParentReviewedAt:new Date(), ParentDecision:approved?'approved':'corrections', ParentComments:comments });

  if (approved) {
    var nextUnit = _nextUnit(ss, unitId);
    if (nextUnit) _setProg(ss, studentId, nextUnit.UnitID, { Status:'available', UnlockedAt:new Date() });
    var subj = '✅ Approved! Your next chapter is ready — '+(nextUnit?nextUnit.UnitName:'All done!');
    var body = _approvalEmail(student, unit, nextUnit);
    try {
      GmailApp.sendEmail(student.StudentEmail, subj, _stripHtml(body), { htmlBody:body, cc:student.ParentEmail, name:PORTAL_NAME });
      _logEmail(ss,'approval',studentId,unitId,student.StudentEmail,subj,'sent');
    } catch(err) { _logEmail(ss,'approval',studentId,unitId,student.StudentEmail,subj,'error:'+err); }
  } else {
    var resubUrl = _prefilledUrl(cfg.HOMEWORK_FORM_ID, { 'Your name':student.StudentName, 'Unit completed':unitId+' — '+unit.UnitName });
    var subj2 = '📝 Corrections needed — '+unit.UnitName;
    var body2 = _correctionsEmail(student, unit, comments, parentName, resubUrl);
    try {
      GmailApp.sendEmail(student.StudentEmail, subj2, _stripHtml(body2), { htmlBody:body2, cc:student.ParentEmail, name:PORTAL_NAME });
      _logEmail(ss,'corrections',studentId,unitId,student.StudentEmail,subj2,'sent');
    } catch(err) { _logEmail(ss,'corrections',studentId,unitId,student.StudentEmail,subj2,'error:'+err); }
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// WEB APP — STUDENT DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════
function doGet() {
  var email   = Session.getActiveUser().getEmail();
  var cfg     = _cfg();
  var ss      = SpreadsheetApp.openById(cfg.SHEET_ID);
  var student = _studentByEmail(ss, email);

  if (!student) {
    return HtmlService.createHtmlOutput(
      '<html><head><meta charset="UTF-8"><style>body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;background:#f0f4f8;}.box{background:white;border-radius:16px;padding:48px;text-align:center;max-width:420px;box-shadow:0 4px 24px rgba(0,0,0,.08);}h2{color:#1d4ed8;}p{color:#64748b;margin-top:12px;}</style></head>' +
      '<body><div class="box"><h2>Not Enrolled</h2><p>The email <strong>'+email+'</strong> is not in the portal.</p><p>Contact your teacher at <a href="mailto:'+ADMIN_EMAIL+'">'+ADMIN_EMAIL+'</a>.</p></div></body></html>'
    ).setTitle('IB Physics Portal').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var allUnits = _units(ss);
  var progMap  = {};
  _progress(ss, student.StudentID).forEach(function(p){ progMap[p.UnitID]=p; });

  // Group by section
  var sections={}, secOrder=[];
  allUnits.forEach(function(u){
    var k=u.SectionNum+'|'+u.SectionName;
    if(!sections[k]){sections[k]=[];secOrder.push(k);}
    sections[k].push(u);
  });

  var hwFormUrl = 'https://docs.google.com/forms/d/'+cfg.HOMEWORK_FORM_ID+'/viewform';

  var secHtml = secOrder.map(function(k){
    var parts=k.split('|'), sNum=parts[0], sName=parts[1], list=sections[k];
    var done=list.filter(function(u){var p=progMap[u.UnitID];return p&&p.Status==='complete';}).length;
    var cards=list.map(function(u){
      var p=progMap[u.UnitID]||{Status:'locked'};
      return _unitCard(u,p,hwFormUrl);
    }).join('');
    return '<div class="sec"><div class="sec-hdr"><span class="snum">'+sNum+'</span><span class="sname">'+sName+'</span><span class="sprog">'+done+'/'+list.length+'</span></div><div class="ulist">'+cards+'</div></div>';
  }).join('');

  var css = '<style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:Nunito,Arial,sans-serif;background:#f0f4f8;background-image:radial-gradient(circle,#c7d2fe 1px,transparent 1px);background-size:28px 28px;}header{background:linear-gradient(135deg,#1e3a5f,#1d4ed8);color:white;padding:32px 24px;}.hi{max-width:860px;margin:0 auto;}.lbl{font-size:11px;text-transform:uppercase;letter-spacing:2px;opacity:.75;margin-bottom:6px;}header h1{font-size:1.7rem;font-weight:900;}main{max-width:860px;margin:28px auto;padding:0 16px 60px;}.sec{margin-bottom:28px;}.sec-hdr{display:flex;align-items:center;gap:10px;margin-bottom:10px;padding:0 4px;}.snum{font-size:.7rem;font-weight:800;text-transform:uppercase;letter-spacing:1px;color:#64748b;min-width:36px;}.sname{font-size:.98rem;font-weight:800;flex:1;}.sprog{font-size:.75rem;font-weight:700;color:#64748b;}.ulist{display:flex;flex-direction:column;gap:8px;}.uc{background:white;border-radius:12px;padding:14px 18px;border:2px solid #e2e8f0;}.uc.av{border-color:#93c5fd;}.uc.co{border-color:#86efac;background:#f0fdf4;}.uc.pe{border-color:#fde68a;background:#fffbeb;}.uc.cr{border-color:#fca5a5;background:#fff1f2;}.uc.lk{opacity:.45;}.utop{display:flex;align-items:center;gap:10px;}.uicon{font-size:1.1rem;flex-shrink:0;}.uname{flex:1;font-size:.9rem;font-weight:700;}.ust{font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#64748b;}.links{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px;align-items:center;}.lb{display:flex;align-items:center;justify-content:center;padding:6px 14px;border-radius:8px;font-size:.78rem;font-weight:700;text-decoration:none;background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;white-space:nowrap;}.lb.sub{background:#1d4ed8;color:white;border-color:#1d4ed8;}.fb{font-size:.8rem;color:#92400e;background:#fff7ed;border-radius:6px;padding:8px 12px;margin-top:8px;}@import url("https://fonts.googleapis.com/css2?family=Nunito:wght@700;800;900&display=swap");</style>';

  var html = '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>IB Physics Portal</title>'+css+'</head><body>'+
    '<header><div class="hi"><div class="lbl">IB Physics Portal</div><h1>'+student.StudentName+'</h1></div></header>'+
    '<main>'+secHtml+'</main></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('IB Physics Portal — '+student.StudentName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _unitCard(unit, prog, hwFormUrl) {
  var st=prog.Status||'locked';
  var icons={locked:'🔒',available:'📖',in_progress:'✏️',awaiting_review:'🟡',corrections:'❌',complete:'✅'};
  var labels={locked:'Locked',available:'Ready',in_progress:'In Progress',awaiting_review:'Awaiting Review',corrections:'Corrections Required',complete:'Complete'};
  var cls={locked:'lk',available:'av',in_progress:'av',awaiting_review:'pe',corrections:'cr',complete:'co'};

  var links='';
  if(st!=='locked'){
    if(unit.Type==='story'){
      links='<div class="links"><a href="'+unit.LessonURL+'" target="_blank" class="lb">📖 Read Story</a></div>';
    } else {
      var l=unit.LessonURL    ?'<a href="'+unit.LessonURL+'"    target="_blank" class="lb">📖 Lesson</a>'   :'';
      var c=unit.ClassworkURL ?'<a href="'+unit.ClassworkURL+'" target="_blank" class="lb">✏️ Classwork</a>'  :'';
      var hw='';
      if(['available','in_progress','corrections'].indexOf(st)>=0){
        var hwPage = unit.HomeworkURL ? '<a href="'+unit.HomeworkURL+'" target="_blank" class="lb">📝 Homework</a>' : '';
        hw = hwPage + '<a href="'+hwFormUrl+'" target="_blank" class="lb sub">📤 Submit</a>';
      }
      links='<div class="links">'+l+c+hw+'</div>';
    }
  }

  var fb=(st==='corrections'&&prog.ParentComments)?'<div class="fb">💬 '+prog.ParentComments+'</div>':'';

  return '<div class="uc '+cls[st]+'"><div class="utop"><span class="uicon">'+icons[st]+'</span><span class="uname">'+unit.UnitName+'</span><span class="ust">'+labels[st]+'</span></div>'+fb+links+'</div>';
}

// ═══════════════════════════════════════════════════════════════════════════════
// EMAIL TEMPLATES
// ═══════════════════════════════════════════════════════════════════════════════
var _ES='<style>@import url("https://fonts.googleapis.com/css2?family=Nunito:wght@700;800&display=swap");body{font-family:Nunito,Arial,sans-serif;background:#f8faff;margin:0;padding:0;}.w{max-width:560px;margin:32px auto;background:white;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08);}.h{background:linear-gradient(135deg,#1e3a5f,#1d4ed8);color:white;padding:32px 36px;}.h .lbl{font-size:11px;text-transform:uppercase;letter-spacing:2px;opacity:.8;margin-bottom:6px;}.h h1{font-size:1.35rem;font-weight:800;margin:0;}.b{padding:32px 36px;color:#1a1a2e;font-size:.94rem;line-height:1.7;}.badge{background:#eff6ff;border:2px solid #bfdbfe;border-radius:10px;padding:14px 18px;margin:18px 0;font-weight:700;color:#1d4ed8;}.btn{display:inline-block;padding:13px 26px;border-radius:10px;font-weight:800;font-size:.92rem;text-decoration:none;margin:6px 6px 6px 0;}.bg{background:#22c55e;color:white;}.br{background:#ef4444;color:white;}.bb{background:#1d4ed8;color:white;}.note{font-size:.8rem;color:#64748b;margin-top:18px;padding-top:14px;border-top:1px solid #e2e8f0;}.fb{background:#fff7ed;border:2px solid #fed7aa;border-radius:10px;padding:14px 18px;margin:16px 0;color:#92400e;}</style>';

function _parentEmail(student,unit,fileUrl,approvalUrl,notes){
  var keyUrl = unit.HomeworkURL ? unit.HomeworkURL + '?key=show' : '';
  return '<!DOCTYPE html><html><head>'+_ES+'</head><body><div class="w">'+
    '<div class="h"><div class="lbl">IB Physics Portal</div><h1>'+student.StudentName+' submitted homework</h1></div>'+
    '<div class="b"><p>Dear '+student.ParentName+',</p>'+
    '<p>'+student.StudentName+' has submitted homework for:</p>'+
    '<div class="badge">📖 '+unit.SectionNum+' '+unit.SectionName+' — '+unit.UnitName+'</div>'+
    (fileUrl?'<p><a href="'+fileUrl+'" class="btn bb">View '+student.StudentName+'\'s Homework →</a></p>':'')+
    (keyUrl?'<p><a href="'+keyUrl+'" class="btn bb">🔑 View Answer Key →</a></p>':'')+
    '<p>Please review '+student.StudentName+'\'s work against the answer key, then click a button:</p>'+
    (notes?'<div class="fb"><strong>Note from '+student.StudentName+':</strong><br>'+notes+'</div>':'')+
    '<p><a href="'+approvalUrl+'" class="btn bg">✅ Approve &amp; Unlock Next Chapter</a>'+
    '<a href="'+approvalUrl+'" class="btn br">❌ Request Corrections</a></p>'+
    '<p class="note">These links open a pre-filled Google Form — just choose your decision and submit.</p>'+
    '</div></div></body></html>';
}

function _approvalEmail(student,unit,nextUnit){
  var nxt=nextUnit&&nextUnit.LessonURL
    ?'<p>Your next chapter is now unlocked:</p><div class="badge">🔓 '+nextUnit.SectionNum+' '+nextUnit.SectionName+' — '+nextUnit.UnitName+'</div><p><a href="'+nextUnit.LessonURL+'" class="btn bb">Start '+nextUnit.UnitName+' →</a></p>'
    :'<p>You\'ve completed all available units — well done! 🎉</p>';
  return '<!DOCTYPE html><html><head>'+_ES+'</head><body><div class="w">'+
    '<div class="h"><div class="lbl">IB Physics Portal</div><h1>Chapter Approved ✅</h1></div>'+
    '<div class="b"><p>Hi '+student.StudentName+',</p>'+
    '<p>Your parent approved your homework for:</p>'+
    '<div class="badge">✅ '+unit.UnitName+'</div>'+nxt+
    '<p class="note">Keep it up!</p></div></div></body></html>';
}

function _correctionsEmail(student,unit,comments,parentName,resubUrl){
  return '<!DOCTYPE html><html><head>'+_ES+'</head><body><div class="w">'+
    '<div class="h"><div class="lbl">IB Physics Portal</div><h1>Corrections Needed 📝</h1></div>'+
    '<div class="b"><p>Hi '+student.StudentName+',</p>'+
    '<p>Your parent reviewed your homework for:</p>'+
    '<div class="badge">📖 '+unit.UnitName+'</div>'+
    '<p>'+(parentName||'Your parent')+' has some feedback:</p>'+
    '<div class="fb">'+(comments||'Please review and redo your homework.')+'</div>'+
    '<p><a href="'+resubUrl+'" class="btn bb">📤 Resubmit Homework →</a></p>'+
    '<p class="note">Your next chapter unlocks once your parent approves your corrected work.</p>'+
    '</div></div></body></html>';
}
