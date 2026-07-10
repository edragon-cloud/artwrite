/*───────────────────────────────────────────────────────────────
  ArticuWrite — Google Apps Script Backend (Code.gs)
  Spreadsheet-backed API for student writing + teacher dashboard.

  DEPLOY: Extensions ▸ Apps Script ▸ paste ▸ Deploy ▸ New deployment
          ▸ Web app ▸ Execute as: Me ▸ Who has access: Anyone
          ▸ copy /exec URL into the GAS constant in each .html

  Router: both doGet (JSONP) and doPost (fetch) dispatch on `action`.
  All responses: { success:Boolean, data|error, ... }
───────────────────────────────────────────────────────────────*/

const SHEET_ID = '1-rtR_ikTdRx1R-2OiIZDxECoH2KY_Fa1Fqsuixab8Y0'; // <-- your sheet
const LIVE_STALE_MS = 25 * 1000; // heartbeat older than this = student went offline

// Tab names
const T = {
  STUDENTS: 'Students',
  TEACHERS: 'Teachers',
  HISTORY:  'History',
  FREE:     'Free-writing',
  HOMEWORK: 'Homework',
  INCLASS:  'In-class Practice',
  LIBRARY:  'Library',
  ASSIGN:   'Assignments',   // teacher-created homework/in-class tasks
  LIVE:     'LiveSessions',  // 1 row/student, overwritten each heartbeat
  ANNOT:    'Annotations',   // teacher inline suggestions + scores
  BOARDS:   'Boards',        // interactive class boards (teacher-editable)
};

// Column headers per tab (order matters — used on auto-create)
const HEADERS = {
  [T.STUDENTS]: ['Student ID','Name','Class','Birthdate','Password','Phone','Email','CreatedAt'],
  [T.TEACHERS]: ['Name','Class','Birthdate','Password','Phone','Email','CreatedAt'],
  [T.HISTORY]:  ['Timestamp','Student ID','Name','Practice Mode','Topic','Topic ID'],
  [T.FREE]:     ['Timestamp','Student ID','Name','Topic','Topic ID','Start time','Finish time','Duration','CC','TA','Grammar','AI Grading','Teacher Grading','Attempt','Essay'],
  [T.HOMEWORK]: ['Timestamp','Student ID','Name','Topic','Topic ID','Start time','Finish time','Duration','CC','TA','Grammar','AI Grading','Teacher Grading','Google Doc Link','Attempt','Essay'],
  [T.INCLASS]:  ['Timestamp','Student ID','Name','Topic','Topic ID','Start time','Finish time','Duration','CC','TA','Grammar','AI Grading','Teacher Grading','Google Doc Link','Attempt','Essay'],
  [T.LIBRARY]:  ['Topic','Question','Task Type','Difficulty','Model Paragraph'],
  [T.ASSIGN]:   ['Topic ID','Mode','Class','Topic','Prompt','Required Attempts','Duration Min','CreatedAt','Active'],
  [T.LIVE]:     ['Student ID','Name','Class','Topic ID','Topic','Mode','Status','Word Count','Snapshot','Updated'],
  [T.ANNOT]:    ['Timestamp','Student ID','Topic ID','Mode','Teacher','Suggestions','Annotated HTML','TR','CC','LR','GRA','Note'],
  [T.BOARDS]:   ['Board ID','Class','Title','Content','Owner','CreatedAt','UpdatedAt'],
};

/*──────────────── ROUTER ────────────────*/
function doGet(e)  { return handle(e, 'GET'); }
function doPost(e) { return handle(e, 'POST'); }

function handle(e, method) {
  var params  = (e && e.parameter) || {};
  var action  = params.action || '';
  var payload = {};
  try {
    if (method === 'POST' && e.postData && e.postData.contents) {
      var body = JSON.parse(e.postData.contents);
      action  = body.action || action;
      payload = body.payload || {};
    } else if (params.payload) {
      payload = JSON.parse(params.payload);
    }
  } catch (err) {
    return respond(e, { success: false, error: 'Bad payload: ' + err.message });
  }

  var out;
  try {
    out = dispatch(action, payload);
  } catch (err) {
    out = { success: false, error: err.message, action: action };
  }
  return respond(e, out);
}

function dispatch(action, p) {
  switch (action) {
    // ── auth ──
    case 'auth.studentSignup':  return studentSignup(p);
    case 'auth.studentLogin':   return studentLogin(p);
    case 'auth.teacherLogin':   return teacherLogin(p);
    case 'auth.teacherSignup':  return teacherSignup(p);
    // ── library ──
    case 'write.getLibrary':    return getLibrary();
    // ── student writing ──
    case 'write.saveResult':    return saveResult(p);
    case 'write.getHistory':    return getHistory(p);
    case 'write.getAttempt':    return getAttempt(p);   // for Re-try
    case 'write.heartbeat':     return heartbeat(p);
    case 'write.getAssignments':return getAssignments(p);
    // ── teacher ──
    case 'teacher.getLive':     return getLive(p);
    case 'teacher.getResults':  return getResults(p);
    case 'teacher.saveAnnotation': return saveAnnotation(p);
    case 'teacher.getOverview': return getOverview(p);
    case 'teacher.createAssignment': return createAssignment(p);
    // ── boards ──
    case 'board.create':        return boardCreate(p);
    case 'board.list':          return boardList(p);
    case 'board.get':           return boardGet(p);
    case 'board.save':          return boardSave(p);
    case 'board.delete':        return boardDelete(p);
    case 'board.uploadImage':   return boardUploadImage(p);
    case 'ping':                return { success: true, data: 'pong', time: new Date().toISOString() };
    default:  return { success: false, error: 'Unknown action: ' + action };
  }
}

/*──────────────── RESPONSE (JSONP-aware) ────────────────*/
function respond(e, obj) {
  var json = JSON.stringify(obj);
  var cb   = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/*──────────────── SHEET HELPERS ────────────────*/
function ss() { return SpreadsheetApp.openById(SHEET_ID); }

function sheet(name) {
  var book = ss();
  var sh = book.getSheetByName(name);
  if (!sh) {
    sh = book.insertSheet(name);
    var head = HEADERS[name];
    if (head) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

// Return array of row objects keyed by header
function readAll(name) {
  var sh = sheet(name);
  var rng = sh.getDataRange().getValues();
  if (rng.length < 2) return [];
  var head = rng[0];
  var rows = [];
  for (var i = 1; i < rng.length; i++) {
    var o = { _row: i + 1 };
    for (var c = 0; c < head.length; c++) o[head[c]] = rng[i][c];
    rows.push(o);
  }
  return rows;
}

function headerIndex(name) {
  var sh = sheet(name);
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idx = {};
  head.forEach(function (h, i) { idx[h] = i; });
  return idx;
}

function appendRowByHeader(name, obj) {
  var sh = sheet(name);
  var head = HEADERS[name] || sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = head.map(function (h) { return obj[h] != null ? obj[h] : ''; });
  sh.appendRow(row);
  return sh.getLastRow();
}

function nowIso() { return new Date().toISOString(); }
function uid(prefix) { return (prefix||'') + Date.now().toString(36) + Math.floor(Math.random()*1e4).toString(36); }

/*──────────────── AUTH ────────────────*/
function studentSignup(p) {
  if (!p.studentId || !p.password) return { success:false, error:'Thiếu Student ID hoặc mật khẩu.' };
  var rows = readAll(T.STUDENTS);
  if (rows.some(function(r){ return String(r['Student ID']) === String(p.studentId); }))
    return { success:false, error:'Student ID đã tồn tại.' };
  appendRowByHeader(T.STUDENTS, {
    'Student ID': p.studentId, 'Name': p.name||'', 'Class': p.class||'',
    'Birthdate': p.birthdate||'', 'Password': p.password, 'Phone': p.phone||'',
    'Email': p.email||'', 'CreatedAt': nowIso()
  });
  return { success:true, data:{ studentId:p.studentId, name:p.name, class:p.class } };
}

function studentLogin(p) {
  var rows = readAll(T.STUDENTS);
  var u = rows.filter(function(r){
    return String(r['Student ID'])===String(p.studentId) && String(r['Password'])===String(p.password);
  })[0];
  if (!u) return { success:false, error:'Sai Student ID hoặc mật khẩu.' };
  return { success:true, data:{ studentId:u['Student ID'], name:u['Name'], class:u['Class'], email:u['Email'] } };
}

function teacherSignup(p) {
  if (!p.email || !p.password) return { success:false, error:'Thiếu email hoặc mật khẩu.' };
  var rows = readAll(T.TEACHERS);
  if (rows.some(function(r){ return String(r['Email']).toLowerCase()===String(p.email).toLowerCase(); }))
    return { success:false, error:'Email đã tồn tại.' };
  appendRowByHeader(T.TEACHERS, {
    'Name': p.name||'', 'Class': p.class||'', 'Birthdate': p.birthdate||'',
    'Password': p.password, 'Phone': p.phone||'', 'Email': p.email, 'CreatedAt': nowIso()
  });
  return { success:true, data:{ name:p.name, email:p.email, class:p.class } };
}

function teacherLogin(p) {
  var rows = readAll(T.TEACHERS);
  var u = rows.filter(function(r){
    return String(r['Email']).toLowerCase()===String(p.email).toLowerCase() && String(r['Password'])===String(p.password);
  })[0];
  if (!u) return { success:false, error:'Sai email hoặc mật khẩu.' };
  return { success:true, data:{ name:u['Name'], email:u['Email'], class:u['Class'] } };
}

/*──────────────── LIBRARY ────────────────*/
// Returns [{topic, questions:[{question, taskType, difficulty, modelParagraph}]}]
function getLibrary() {
  var rows = readAll(T.LIBRARY);
  var byTopic = {};
  rows.forEach(function(r){
    var topic = r['Topic'] || 'Uncategorized';
    if (!byTopic[topic]) byTopic[topic] = { topic: topic, questions: [] };
    byTopic[topic].questions.push({
      question: r['Question'] || '',
      taskType: r['Task Type'] || '',
      difficulty: r['Difficulty'] || '',
      modelParagraph: r['Model Paragraph'] || ''
    });
  });
  return { success:true, data: Object.keys(byTopic).map(function(k){ return byTopic[k]; }) };
}

/*──────────────── SAVE RESULT ────────────────
  Called after client-side Gemini grading. Writes to the mode-specific
  tab AND appends a History row. Attempt # derived from prior rows with
  same Student ID + Topic ID.
  p = { studentId, name, mode('free'|'homework'|'inclass'), topic, topicId,
        startTime, finishTime, duration, cc, ta, grammar, aiGrading,
        teacherGrading, docLink, essay }
─────────────────────────────────────────────*/
function saveResult(p) {
  var tabMap = { free:T.FREE, homework:T.HOMEWORK, inclass:T.INCLASS };
  var tab = tabMap[p.mode];
  if (!tab) return { success:false, error:'Mode không hợp lệ: ' + p.mode };

  // attempt number
  var prior = readAll(tab).filter(function(r){
    return String(r['Student ID'])===String(p.studentId) && String(r['Topic ID'])===String(p.topicId);
  });
  var attempt = prior.length + 1;

  var row = {
    'Timestamp': nowIso(), 'Student ID': p.studentId, 'Name': p.name||'',
    'Topic': p.topic||'', 'Topic ID': p.topicId||'',
    'Start time': p.startTime||'', 'Finish time': p.finishTime||'',
    'Duration': p.duration||'', 'CC': p.cc!=null?p.cc:'', 'TA': p.ta!=null?p.ta:'',
    'Grammar': p.grammar!=null?p.grammar:'', 'AI Grading': p.aiGrading!=null?p.aiGrading:'',
    'Teacher Grading': p.teacherGrading!=null?p.teacherGrading:'',
    'Attempt': attempt, 'Essay': p.essay||''
  };
  if (tab === T.HOMEWORK || tab === T.INCLASS) row['Google Doc Link'] = p.docLink||'';

  appendRowByHeader(tab, row);

  appendRowByHeader(T.HISTORY, {
    'Timestamp': nowIso(), 'Student ID': p.studentId, 'Name': p.name||'',
    'Practice Mode': p.mode, 'Topic': p.topic||'', 'Topic ID': p.topicId||''
  });

  // clear this student's live session (they finished)
  clearLive(p.studentId);

  return { success:true, data:{ attempt:attempt, topicId:p.topicId } };
}

/*──────────────── HISTORY + RE-TRY ────────────────*/
function getHistory(p) {
  var rows = readAll(T.HISTORY).filter(function(r){
    return String(r['Student ID'])===String(p.studentId);
  });
  // dedupe by Topic ID keeping latest, so Re-try list shows one row per topic
  var seen = {};
  rows.reverse().forEach(function(r){ if(!seen[r['Topic ID']]) seen[r['Topic ID']] = r; });
  return { success:true, data: Object.keys(seen).map(function(k){
    return { topicId:k, topic:seen[k]['Topic'], mode:seen[k]['Practice Mode'], timestamp:seen[k]['Timestamp'] };
  }) };
}

// Return the most recent attempt for a topicId (question + feedback) so student can re-write
function getAttempt(p) {
  var tabs = [T.FREE, T.HOMEWORK, T.INCLASS];
  var found = null;
  tabs.forEach(function(tab){
    readAll(tab).forEach(function(r){
      if (String(r['Student ID'])===String(p.studentId) && String(r['Topic ID'])===String(p.topicId)) {
        if (!found || new Date(r['Timestamp']) > new Date(found['Timestamp'])) found = r;
      }
    });
  });
  if (!found) return { success:false, error:'Không tìm thấy bài cũ.' };
  return { success:true, data:{
    topic: found['Topic'], topicId: found['Topic ID'], essay: found['Essay'],
    cc: found['CC'], ta: found['TA'], grammar: found['Grammar'],
    aiGrading: found['AI Grading'], teacherGrading: found['Teacher Grading']
  } };
}

/*──────────────── ASSIGNMENTS (teacher-created tasks) ────────────────*/
function createAssignment(p) {
  var topicId = p.topicId || uid('A');
  appendRowByHeader(T.ASSIGN, {
    'Topic ID': topicId, 'Mode': p.mode||'homework', 'Class': p.class||'',
    'Topic': p.topic||'', 'Prompt': p.prompt||'',
    'Required Attempts': p.requiredAttempts||1, 'Duration Min': p.durationMin||'',
    'CreatedAt': nowIso(), 'Active': true
  });
  return { success:true, data:{ topicId:topicId } };
}

function getAssignments(p) {
  var rows = readAll(T.ASSIGN).filter(function(r){
    if (r['Active'] === false) return false;
    if (p.class && String(r['Class']) !== String(p.class)) return false;
    if (p.mode && String(r['Mode']) !== String(p.mode)) return false;
    return true;
  });
  return { success:true, data: rows.map(function(r){
    return { topicId:r['Topic ID'], mode:r['Mode'], class:r['Class'], topic:r['Topic'],
             prompt:r['Prompt'], requiredAttempts:r['Required Attempts'], durationMin:r['Duration Min'] };
  }) };
}

/*──────────────── LIVE OBSERVATION ────────────────
  heartbeat: upsert one row per student into LiveSessions.
  Snapshot text is stored so the teacher's click is instant (no extra call).
─────────────────────────────────────────────*/
function heartbeat(p) {
  if (!p.studentId) return { success:false, error:'Thiếu studentId.' };
  var sh = sheet(T.LIVE);
  var idx = headerIndex(T.LIVE);
  var data = sh.getDataRange().getValues();
  var rowNum = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idx['Student ID']]) === String(p.studentId)) { rowNum = i + 1; break; }
  }
  var vals = [
    p.studentId, p.name||'', p.class||'', p.topicId||'', p.topic||'',
    p.mode||'', p.status||'Writing', p.wordCount||0, (p.snapshot||'').slice(0, 8000), nowIso()
  ];
  if (rowNum === -1) sh.appendRow(vals);
  else sh.getRange(rowNum, 1, 1, vals.length).setValues([vals]);
  return { success:true };
}

function clearLive(studentId) {
  var sh = sheet(T.LIVE);
  var idx = headerIndex(T.LIVE);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idx['Student ID']]) === String(studentId)) {
      // mark Submitted rather than delete, so teacher sees final state briefly
      sh.getRange(i + 1, idx['Status'] + 1).setValue('Submitted');
      sh.getRange(i + 1, idx['Updated'] + 1).setValue(nowIso());
      return;
    }
  }
}

// Teacher poll: return all live sessions, marking stale ones Offline
function getLive(p) {
  var rows = readAll(T.LIVE);
  var now = Date.now();
  var out = rows.filter(function(r){
    return !p.class || String(r['Class'])===String(p.class);
  }).map(function(r){
    var age = now - new Date(r['Updated']).getTime();
    var status = r['Status'];
    if (status !== 'Submitted' && age > LIVE_STALE_MS) status = 'Offline';
    return {
      studentId:r['Student ID'], name:r['Name'], class:r['Class'],
      topicId:r['Topic ID'], topic:r['Topic'], mode:r['Mode'],
      status:status, wordCount:r['Word Count'], snapshot:r['Snapshot'],
      updated:r['Updated']
    };
  });
  return { success:true, data: out };
}

/*──────────────── TEACHER ANNOTATION / SCORING ────────────────*/
function saveAnnotation(p) {
  appendRowByHeader(T.ANNOT, {
    'Timestamp': nowIso(), 'Student ID': p.studentId, 'Topic ID': p.topicId,
    'Mode': p.mode||'', 'Teacher': p.teacher||'',
    'Suggestions': p.suggestions ? JSON.stringify(p.suggestions) : '',
    'Annotated HTML': (p.annotatedHtml||'').slice(0, 45000),
    'TR': p.tr!=null?p.tr:'', 'CC': p.cc!=null?p.cc:'',
    'LR': p.lr!=null?p.lr:'', 'GRA': p.gra!=null?p.gra:'', 'Note': p.note||''
  });

  // also write teacher overall back to the mode tab (latest attempt for that topicId)
  if (p.teacherGrading != null && p.mode) {
    var tabMap = { free:T.FREE, homework:T.HOMEWORK, inclass:T.INCLASS };
    var tab = tabMap[p.mode];
    if (tab) {
      var sh = sheet(tab);
      var idx = headerIndex(tab);
      var data = sh.getDataRange().getValues();
      var target = -1;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idx['Student ID']])===String(p.studentId) &&
            String(data[i][idx['Topic ID']])===String(p.topicId)) target = i + 1;
      }
      if (target > -1) sh.getRange(target, idx['Teacher Grading'] + 1).setValue(p.teacherGrading);
    }
  }
  return { success:true };
}

/*──────────────── RESULTS TABLE (teacher) ────────────────*/
function getResults(p) {
  var tabMap = { free:T.FREE, homework:T.HOMEWORK, inclass:T.INCLASS };
  var tab = tabMap[p.mode] || T.FREE;
  var rows = readAll(tab).filter(function(r){
    return !p.class || String(r['Class'])===String(p.class);
  });
  return { success:true, data: rows.map(function(r){
    return {
      studentId:r['Student ID'], name:r['Name'], topic:r['Topic'], topicId:r['Topic ID'],
      aiScore:r['AI Grading'], teacherScore:r['Teacher Grading'],
      attempt:r['Attempt'], cc:r['CC'], ta:r['TA'], grammar:r['Grammar'],
      timestamp:r['Timestamp'], docLink:r['Google Doc Link']||''
    };
  }) };
}

/*──────────────── CLASS OVERVIEW (teacher KPIs) ────────────────*/
function getOverview(p) {
  var tabs = [T.FREE, T.HOMEWORK, T.INCLASS];
  var all = [];
  tabs.forEach(function(tab){
    readAll(tab).forEach(function(r){
      if (!p.class || String(r['Class'])===String(p.class)) all.push(r);
    });
  });
  var totalEssays = all.length;
  var scores = all.map(function(r){ return parseFloat(r['AI Grading']); }).filter(function(n){ return !isNaN(n); });
  var avg = scores.length ? (scores.reduce(function(a,b){return a+b;},0)/scores.length) : 0;
  var students = {};
  all.forEach(function(r){ students[r['Student ID']] = true; });
  var pending = all.filter(function(r){ return r['Teacher Grading']===''||r['Teacher Grading']==null; }).length;

  return { success:true, data:{
    avgScore: Math.round(avg*100)/100,
    totalEssays: totalEssays,
    activeStudents: Object.keys(students).length,
    feedbackPending: pending
  } };
}

/*──────────────── INTERACTIVE BOARDS ────────────────
  Teacher creates a board for a class. Students in that class
  see it (read-only). Only the teacher edits/formats. Content
  is HTML (rich text + <img> with Drive links or small base64).
─────────────────────────────────────────────*/
function boardCreate(p) {
  var id = uid('B');
  appendRowByHeader(T.BOARDS, {
    'Board ID': id, 'Class': p.class||'', 'Title': p.title||'Untitled Board',
    'Content': '', 'Owner': p.owner||'', 'CreatedAt': nowIso(), 'UpdatedAt': nowIso()
  });
  return { success:true, data:{ boardId:id, title:p.title||'Untitled Board' } };
}

function boardList(p) {
  var rows = readAll(T.BOARDS).filter(function(r){
    return !p.class || String(r['Class'])===String(p.class);
  });
  return { success:true, data: rows.map(function(r){
    return { boardId:r['Board ID'], class:r['Class'], title:r['Title'],
             owner:r['Owner'], updatedAt:r['UpdatedAt'] };
  }) };
}

function boardGet(p) {
  var b = readAll(T.BOARDS).filter(function(r){ return String(r['Board ID'])===String(p.boardId); })[0];
  if (!b) return { success:false, error:'Board không tồn tại.' };
  return { success:true, data:{
    boardId:b['Board ID'], class:b['Class'], title:b['Title'],
    content:b['Content'], owner:b['Owner'], updatedAt:b['UpdatedAt']
  } };
}

function boardSave(p) {
  var sh = sheet(T.BOARDS);
  var idx = headerIndex(T.BOARDS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idx['Board ID']])===String(p.boardId)) {
      if (p.title != null) sh.getRange(i+1, idx['Title']+1).setValue(p.title);
      if (p.content != null) sh.getRange(i+1, idx['Content']+1).setValue(String(p.content).slice(0, 48000));
      sh.getRange(i+1, idx['UpdatedAt']+1).setValue(nowIso());
      return { success:true, data:{ boardId:p.boardId } };
    }
  }
  return { success:false, error:'Board không tồn tại.' };
}

function boardDelete(p) {
  var sh = sheet(T.BOARDS);
  var idx = headerIndex(T.BOARDS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idx['Board ID']])===String(p.boardId)) {
      sh.deleteRow(i+1); return { success:true };
    }
  }
  return { success:false, error:'Board không tồn tại.' };
}

// Upload a base64 image to Drive, return a public view link (teacher only)
function boardUploadImage(p) {
  if (!p.dataUrl) return { success:false, error:'Thiếu ảnh.' };
  try {
    var parts = p.dataUrl.match(/^data:([^;]+);base64,(.+)$/);
    if (!parts) return { success:false, error:'Định dạng ảnh không hợp lệ.' };
    var blob = Utilities.newBlob(Utilities.base64Decode(parts[2]), parts[1], p.name || ('board-' + Date.now()));
    var folder = getBoardFolder();
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var id = file.getId();
    return { success:true, data:{ url: 'https://drive.google.com/uc?export=view&id=' + id, id:id } };
  } catch (err) {
    return { success:false, error:'Upload lỗi: ' + err.message };
  }
}

function getBoardFolder() {
  var name = 'ArticuWrite Board Images';
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/*──────────────── OPTIONAL: one-time setup ────────────────
  Run manually in the Apps Script editor (Run ▸ setup) to
  pre-create every tab with headers. Not required — tabs are
  auto-created on first write — but handy for a clean start.
─────────────────────────────────────────────*/
function setup() {
  Object.keys(HEADERS).forEach(function(name){ sheet(name); });
  Logger.log('All tabs created: ' + Object.keys(HEADERS).join(', '));
}
