/**
 * SmartFlex — Google Apps Script backend (Code.gs) v2
 * รองรับ: โหมดส่วนบุคคล + โหมดห้องเรียน (รายบุคคล), เหตุการณ์เตือน, ผู้ใช้, การตั้งค่า
 *
 * ติดตั้ง
 * 1. เปิด Google Sheet → Extensions → Apps Script → วางไฟล์นี้แทน Code.gs → Save
 * 2. Deploy → New deployment → Web app
 *      Execute as: Me     |     Who has access: Anyone
 * 3. คัดลอก URL /exec ไปใส่ในตัวแปร DEFAULT_URL ของ index.html
 * 4. แก้โค้ดครั้งต่อไป: Deploy → Manage deployments → Edit → Version: New version
 *
 * ชีทที่สร้างอัตโนมัติ: Sessions · StudentResults · Events · Users · Settings
 */

var SHEETS = {
  Sessions: ['sessionId','savedAt','appMode','userId','name','room','activity','device',
             'cameraAngle','activityMode','totalSeconds','goodSeconds','badSeconds','unknownSeconds',
             'coverage','score','grade','alertCount','mainIssue','advice','startedAt','endedAt'],
  StudentResults: ['sessionId','savedAt','room','seatNo','studentId','name',
                   'score','grade','goodSeconds','badSeconds','alertCount','mainIssue','advice'],
  Events: ['sessionId','at','elapsed','who','issue','angle','value','level'],
  Users: ['userId','name','room','school','grade','updatedAt'],
  Settings: ['key','value','updatedAt']
};

function doGet(e)  { return handle((e && e.parameter) || {}); }
function doPost(e) {
  var p = {};
  try {
    if (e.postData && e.postData.contents && e.postData.type.indexOf('json') > -1) p = JSON.parse(e.postData.contents);
    else p = (e && e.parameter) || {};
  } catch (err) { p = (e && e.parameter) || {}; }
  return handle(p);
}

function handle(p) {
  var action = p.action || 'ping';
  var out;
  try {
    if (action === 'ping')             out = { ok: true, action: 'ping', sheets: Object.keys(SHEETS) };
    else if (action === 'save')        out = { ok: true, action: 'save', row: saveRow('Sessions', p) };
    else if (action === 'saveStudent') out = { ok: true, action: 'saveStudent', row: saveRow('StudentResults', p, 'sessionId', 'seatNo') };
    else if (action === 'saveEvent')   out = { ok: true, action: 'saveEvent', row: appendRow('Events', p) };
    else if (action === 'saveUser')    out = { ok: true, action: 'saveUser', row: saveRow('Users', p, 'userId') };
    else if (action === 'saveBatch')   out = { ok: true, action: 'saveBatch', rows: saveBatch(p) };
    else if (action === 'list')        out = { ok: true, action: 'list', data: listRows(p) };
    else out = { ok: false, error: 'unknown action: ' + action };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return reply(out, p.callback);
}

/* ---------- helpers ---------- */

function reply(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function sheet(name) {
  var cols = SHEETS[name];
  if (!cols) throw new Error('unknown sheet: ' + name);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(cols);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold').setBackground('#E7F8F0');
    sh.setFrozenRows(1);
  }
  return sh;
}

var NUMERIC = ['totalSeconds','goodSeconds','badSeconds','unknownSeconds','score','alertCount',
               'coverage','seatNo','elapsed','value'];

function toValues(name, p) {
  return SHEETS[name].map(function (key) {
    var v = p[key];
    if (v === undefined || v === null || v === '') return '';
    if (NUMERIC.indexOf(key) > -1) { var n = Number(v); return isNaN(n) ? '' : n; }
    return String(v);
  });
}

/** บันทึกแบบกันซ้ำ: คีย์เดียวหรือคีย์ผสม (เช่น sessionId + seatNo) */
function saveRow(name, p, k1, k2) {
  k1 = k1 || 'sessionId';
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet(name), cols = SHEETS[name];
    var key1 = String(p[k1] || '').trim();
    if (!key1) throw new Error('missing ' + k1);
    var values = toValues(name, p);
    var savedIdx = cols.indexOf('savedAt');
    if (savedIdx > -1 && !values[savedIdx]) values[savedIdx] = new Date().toISOString();

    var found = findRow(sh, cols, k1, key1, k2, k2 ? String(p[k2] || '') : null);
    if (found > 0) { sh.getRange(found, 1, 1, cols.length).setValues([values]); return found; }
    sh.appendRow(values);
    return sh.getLastRow();
  } finally { lock.releaseLock(); }
}

function appendRow(name, p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet(name);
    sh.appendRow(toValues(name, p));
    return sh.getLastRow();
  } finally { lock.releaseLock(); }
}

function findRow(sh, cols, k1, v1, k2, v2) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var i1 = cols.indexOf(k1), i2 = k2 ? cols.indexOf(k2) : -1;
  var data = sh.getRange(2, 1, last - 1, cols.length).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][i1]) !== String(v1)) continue;
    if (i2 > -1 && String(data[i][i2]) !== String(v2)) continue;
    return i + 2;
  }
  return -1;
}

/** รับหลายแถวพร้อมกัน: p.rows = JSON string ของ [{sheet:'StudentResults', ...}, ...] */
function saveBatch(p) {
  var rows = [];
  try { rows = JSON.parse(p.rows || '[]'); } catch (e) { throw new Error('rows is not valid JSON'); }
  var n = 0;
  rows.forEach(function (r) {
    var name = r.sheet || 'Sessions';
    if (name === 'Events') appendRow(name, r);
    else if (name === 'StudentResults') saveRow(name, r, 'sessionId', 'seatNo');
    else if (name === 'Users') saveRow(name, r, 'userId');
    else saveRow('Sessions', r);
    n++;
  });
  return n;
}

function listRows(p) {
  var name = p.sheet || 'Sessions';
  var cols = SHEETS[name];
  var sh = sheet(name), last = sh.getLastRow();
  if (last < 2) return [];
  var limit = Number(p.limit) || 100;
  var data = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var out = [];
  for (var i = data.length - 1; i >= 0 && out.length < limit; i--) {
    var o = {};
    for (var c = 0; c < cols.length; c++) o[cols[c]] = data[i][c];
    if (p.userId && String(o.userId) !== String(p.userId)) continue;
    if (p.room   && String(o.room)   !== String(p.room))   continue;
    if (p.appMode && String(o.appMode) !== String(p.appMode)) continue;
    if (p.from && String(o.startedAt || o.savedAt) < String(p.from)) continue;
    if (p.to   && String(o.startedAt || o.savedAt) > String(p.to))   continue;
    out.push(o);
  }
  return out;
}

/** ทดสอบจากตัวแก้ไข */
function testSave() {
  var id = 'TEST-' + Date.now();
  saveRow('Sessions', { sessionId: id, appMode: 'classroom', userId: 'SF-0001', name: 'ทดสอบ',
    room: 'ม.5/2', activity: 'เรียน', device: 'มือถือ', cameraAngle: 'side', activityMode: 'side',
    totalSeconds: 600, goodSeconds: 450, badSeconds: 150, unknownSeconds: 0, coverage: 92,
    score: 75, grade: 'ควรปรับ', alertCount: 2, mainIssue: 'คอยื่น', advice: 'ยกจอขึ้นระดับสายตา',
    startedAt: new Date().toISOString(), endedAt: new Date().toISOString() });
  saveRow('StudentResults', { sessionId: id, room: 'ม.5/2', seatNo: 1, studentId: '01', name: 'นักเรียน 1',
    score: 82, grade: 'ดี', goodSeconds: 500, badSeconds: 100, alertCount: 1, mainIssue: 'ไหล่เอียง',
    advice: 'ปรับไหล่ให้เท่ากัน' }, 'sessionId', 'seatNo');
  appendRow('Events', { sessionId: id, at: new Date().toISOString(), elapsed: 120,
    who: 'นักเรียน 1', issue: 'คอยื่น', angle: 'forwardHead', value: 42, level: 'alert' });
  Logger.log('ok ' + id);
}
