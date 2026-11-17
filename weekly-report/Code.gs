/**
 * CSR협력2팀 주간보고 — 구글 시트에 저장하는 웹앱 (서버 쪽 코드)
 *
 * 이 스크립트는 구글 시트에 붙여서 씁니다. 처음 열릴 때 시트에 탭 3개를 직접 만듭니다.
 *   주차  : 주차별 모금현황, 주간일정, 공통업무
 *   업무  : 담당자·기업별 업무 한 줄씩
 *   설정  : 담당자 목록, 연간 모금 목표, 이 도구를 쓸 수 있는 사람(허용 이메일)
 */

const SHEET = { WEEKS: '주차', ITEMS: '업무', SETTINGS: '설정' };
const WEEK_HEAD = ['주차(시작일)', '종료일', '주차명', '누적 모금액(원)', '달성률(%)', '주간일정', '공통업무', '수정시각'];
const ITEM_HEAD = ['ID', '주차(시작일)', '담당자', '기업명', '주요 보고사항 및 업무 진행 내용', '마감기한', '피드백', '순서', '수정시각', '수정자'];
const ITEM_COL = { owner: 3, company: 4, content: 5, due: 6, feedback: 7 };
const WEEK_COL = { fundAmount: 4, fundPct: 5, schedule: 6, common: 7 };
const SET_ROW = { members: 2, fundTarget: 3, allow: 4 };
const TITLE = 'CSR협력2팀 주간보고';

/* ---------- 화면 열기 ---------- */
function doGet() {
  ensureSheets_();
  if (!allowed_()) {
    const who = (me_() || '확인할 수 없는 계정').replace(/[<>&"']/g, '');
    return HtmlService.createHtmlOutput(
      '<div style="font:16px/1.7 sans-serif;max-width:34em;margin:14vh auto;padding:0 20px">' +
      '<h2 style="margin:0 0 8px">접근 권한이 없습니다</h2>' +
      '<p>지금 로그인한 계정: <b>' + who + '</b></p>' +
      '<p>이 도구는 등록된 팀원만 쓸 수 있습니다. 팀 담당자에게 위 이메일을 등록해 달라고 요청하세요.</p></div>'
    ).setTitle(TITLE);
  }
  return HtmlService.createHtmlOutputFromFile('Index').setTitle(TITLE).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 화면에서 부르는 기능 ---------- */

/** 주차 목록, 설정, 선택한 주차의 업무를 한 번에 돌려줍니다. */
function getState(weekId) {
  guard_();
  ensureSheets_();
  const weeks = readWeeks_();
  const id = weekId && weeks.some(function (w) { return w.id === weekId; }) ? weekId : (weeks[0] ? weeks[0].id : '');
  const all = rows_(sheet_(SHEET.ITEMS));
  const items = [];
  const prev = {};
  let prevId = '';
  for (let i = 0; i < weeks.length; i++) { if (id && weeks[i].id < id) { prevId = weeks[i].id; break; } }
  all.forEach(function (r) {
    const wk = str_(r[1]);
    if (!str_(r[0])) return;
    if (id && wk === id) items.push(itemObj_(r));
    else if (prevId && wk === prevId && text_(r[4])) prev[text_(r[2]) + '|' + text_(r[3])] = text_(r[4]);
  });
  return { me: me_(), weeks: weeks, settings: readSettings_(), weekId: id, items: items, prev: prev };
}

function saveItem(id, patch) {
  guard_();
  return withLock_(function () {
    const sh = sheet_(SHEET.ITEMS);
    const row = findRow_(sh, id);
    if (row < 0) throw new Error('다른 팀원이 이 줄을 삭제했습니다. 잠시 후 화면이 새로 고쳐집니다.');
    Object.keys(patch || {}).forEach(function (k) {
      if (ITEM_COL[k]) sh.getRange(row, ITEM_COL[k]).setValue(esc_(patch[k]));
    });
    sh.getRange(row, 9, 1, 2).setValues([[now_(), me_()]]);
    return true;
  });
}

function addItem(weekId, owner) {
  guard_();
  checkWeekId_(weekId);
  return withLock_(function () {
    const sh = sheet_(SHEET.ITEMS);
    const item = { id: newId_(), owner: String(owner || ''), company: '', content: '', due: '', feedback: '', order: Date.now() };
    sh.getRange(sh.getLastRow() + 1, 1, 1, ITEM_HEAD.length).setValues([itemRow_(weekId, item)]);
    return item;
  });
}

function deleteItem(id) {
  guard_();
  return withLock_(function () {
    const sh = sheet_(SHEET.ITEMS);
    const row = findRow_(sh, id);
    if (row > 0) sh.deleteRow(row);
    return true;
  });
}

function saveWeek(weekId, patch) {
  guard_();
  return withLock_(function () {
    const sh = sheet_(SHEET.WEEKS);
    const row = findRow_(sh, weekId);
    if (row < 0) throw new Error('이 주차가 삭제되었습니다. 잠시 후 화면이 새로 고쳐집니다.');
    Object.keys(patch || {}).forEach(function (k) {
      if (!WEEK_COL[k]) return;
      const v = patch[k];
      const isNum = k === 'fundAmount' || k === 'fundPct';
      sh.getRange(row, WEEK_COL[k]).setValue(isNum ? (v === null || v === '' || isNaN(Number(v)) ? '' : Number(v)) : esc_(v));
    });
    sh.getRange(row, 8).setValue(now_());
    return true;
  });
}

/** 새 주차를 만듭니다. copy가 true면 직전 주차의 담당자·기업 목록을 내용 없이 가져옵니다. */
function createWeek(startId, copy) {
  guard_();
  checkWeekId_(startId);
  withLock_(function () {
    const weeks = readWeeks_();
    if (weeks.some(function (w) { return w.id === startId; })) return;
    const wsh = sheet_(SHEET.WEEKS);
    wsh.getRange(wsh.getLastRow() + 1, 1, 1, WEEK_HEAD.length)
      .setValues([[startId, fmtDay_(addDays_(parseDay_(startId), 4)), weekName_(startId), '', '', '', '', now_()]]);
    if (!copy) return;
    let source = '';
    for (let i = 0; i < weeks.length; i++) { if (weeks[i].id < startId) { source = weeks[i].id; break; } }
    if (!source) return;
    const ish = sheet_(SHEET.ITEMS);
    const rows = rows_(ish).filter(function (r) { return str_(r[1]) === source && str_(r[0]); })
      .map(itemObj_).sort(function (a, b) { return a.order - b.order; });
    if (!rows.length) return;
    const out = rows.map(function (x, i) {
      return itemRow_(startId, { id: newId_(), owner: x.owner, company: x.company, content: '', due: '', feedback: '', order: i + 1 });
    });
    ish.getRange(ish.getLastRow() + 1, 1, out.length, ITEM_HEAD.length).setValues(out);
  });
  return getState(startId);
}

function deleteWeek(weekId) {
  guard_();
  withLock_(function () {
    const ish = sheet_(SHEET.ITEMS);
    const rows = rows_(ish);
    for (let i = rows.length - 1; i >= 0; i--) { if (str_(rows[i][1]) === weekId) ish.deleteRow(i + 2); }
    const wsh = sheet_(SHEET.WEEKS);
    const row = findRow_(wsh, weekId);
    if (row > 0) wsh.deleteRow(row);
  });
  return getState('');
}

function saveSettings(patch) {
  guard_();
  return withLock_(function () {
    const sh = sheet_(SHEET.SETTINGS);
    if (patch && Array.isArray(patch.members)) {
      const names = patch.members.map(function (m) { return String(m).replace(/,/g, ' ').trim(); }).filter(String);
      sh.getRange(SET_ROW.members, 2).setValue(esc_(names.join(', ')));
    }
    if (patch && 'fundTarget' in patch) {
      const t = patch.fundTarget;
      sh.getRange(SET_ROW.fundTarget, 2).setValue(t === null || t === '' || isNaN(Number(t)) ? '' : String(Math.round(Number(t))));
    }
    return true;
  });
}

/* ---------- 권한 ---------- */
function me_() { return String(Session.getActiveUser().getEmail() || '').toLowerCase(); }
function owner_() { return String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); }
function allowed_() {
  const me = me_();
  if (me && me === owner_()) return true;
  const sh = sheet_(SHEET.SETTINGS);
  const list = sh ? text_(sh.getRange(SET_ROW.allow, 2).getValue()).toLowerCase().split(/[\s,;]+/).filter(String) : [];
  if (list.indexOf('*') >= 0) return true;
  return !!me && list.indexOf(me) >= 0;
}
function guard_() {
  if (!allowed_()) throw new Error('이 도구를 사용할 권한이 없습니다. 팀 담당자에게 이메일 등록을 요청하세요.');
}

/* ---------- 시트 준비 ---------- */
function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(SHEET.WEEKS) && ss.getSheetByName(SHEET.ITEMS) && ss.getSheetByName(SHEET.SETTINGS)) return;
  withLock_(function () {
    if (!ss.getSheetByName(SHEET.WEEKS)) makeSheet_(ss, SHEET.WEEKS, WEEK_HEAD, [1, 2, 3, 6, 7, 8], [6, 7]);
    if (!ss.getSheetByName(SHEET.ITEMS)) makeSheet_(ss, SHEET.ITEMS, ITEM_HEAD, [1, 2, 3, 4, 5, 6, 7, 9, 10], [5, 7]);
    if (!ss.getSheetByName(SHEET.SETTINGS)) {
      const sh = makeSheet_(ss, SHEET.SETTINGS, ['항목', '값'], [1, 2], [2]);
      sh.getRange(2, 1, 4, 2).setValues([
        ['담당자 목록', ''],
        ['연간 모금 목표(원)', ''],
        ['허용 이메일', owner_()],
        ['안내', '허용 이메일 칸에 이 도구를 쓸 팀원의 구글 계정 이메일을 쉼표로 구분해 적습니다. 별표(*)만 적으면 링크를 열 수 있는 조직 내 모든 사람이 쓸 수 있습니다. 담당자 목록과 목표는 도구 화면에서 바꾸는 것을 권장합니다.']
      ]);
    }
  });
}
function makeSheet_(ss, name, head, textCols, wideCols) {
  const sh = ss.insertSheet(name);
  textCols.forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  wideCols.forEach(function (c) { sh.setColumnWidth(c, 420); sh.getRange(1, c, sh.getMaxRows(), 1).setWrap(true); });
  return sh;
}

/* ---------- 읽고 쓰는 도우미 ---------- */
function sheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function rows_(sh) {
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
}
function findRow_(sh, key) {
  const n = sh.getLastRow();
  if (n < 2 || !key) return -1;
  const col = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) { if (str_(col[i][0]) === String(key)) return i + 2; }
  return -1;
}
function readWeeks_() {
  return rows_(sheet_(SHEET.WEEKS)).filter(function (r) { return str_(r[0]); }).map(function (r) {
    return { id: str_(r[0]), fundAmount: num_(r[3]), fundPct: num_(r[4]), schedule: text_(r[5]), common: text_(r[6]) };
  }).sort(function (a, b) { return a.id < b.id ? 1 : -1; });
}
function readSettings_() {
  const v = sheet_(SHEET.SETTINGS).getRange(2, 2, 2, 1).getValues();
  return {
    members: text_(v[0][0]).split(',').map(function (s) { return s.trim(); }).filter(String),
    fundTarget: num_(String(v[1][0]).replace(/[^\d.]/g, ''))
  };
}
function itemObj_(r) {
  return { id: str_(r[0]), owner: text_(r[2]), company: text_(r[3]), content: text_(r[4]), due: str_(r[5]), feedback: text_(r[6]), order: Number(r[7]) || 0 };
}
function itemRow_(weekId, x) {
  return [x.id, weekId, esc_(x.owner), esc_(x.company), esc_(x.content), esc_(x.due), esc_(x.feedback), x.order, now_(), me_()];
}
function tz_() { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Asia/Seoul'; }
/** 시트가 날짜로 바꿔 버린 값도 "2026-11-16" 글자로 되돌립니다. */
function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v).trim();
}
function text_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const s = v === null || v === undefined ? '' : String(v);
  return /^'[=+]/.test(s) ? s.slice(1) : s;
}
function num_(v) { return v === '' || v === null || v === undefined || isNaN(Number(v)) ? null : Number(v); }
/** "="나 "+"로 시작하는 글이 수식으로 바뀌지 않도록 막습니다. */
function esc_(v) {
  const s = v === null || v === undefined ? '' : String(v);
  return /^[=+]/.test(s) ? "'" + s : s;
}
function now_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'); }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function checkWeekId_(id) { if (!/^\d{4}-\d{2}-\d{2}$/.test(String(id))) throw new Error('주차 날짜가 올바르지 않습니다.'); }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function parseDay_(s) { const p = String(s).split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function addDays_(d, n) { const x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
function fmtDay_(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
/** 주차 이름은 그 주 목요일이 속한 달 기준입니다. (예: 2026-11-16 → 11월 3주차) */
function weekName_(id) {
  const thu = addDays_(parseDay_(id), 3);
  return (thu.getMonth() + 1) + '월 ' + Math.ceil(thu.getDate() / 7) + '주차';
}
