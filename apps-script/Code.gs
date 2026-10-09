/**
 * 인력관리 대시보드용 Google Sheets 백엔드 (Apps Script 웹앱)
 * 시트 구성: '직원'(이름,국적,구분) / '근무기록'(날짜,이름,근무형태,시간,메모)
 */
const TOKEN = '';  // 비밀번호처럼 사용할 문자열 (비워두면 URL만 알면 접근 가능). 대시보드에도 같은 값을 입력.
const LABEL = {full: '전일', half: '반일', part: '시간제'};
const TYPE = {'전일': 'full', '반일': 'half', '시간제': 'part'};
const H_STAFF = ['이름', '국적', '구분'];
const H_REC = ['날짜', '이름', '근무형태', '시간', '메모'];

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  if (name === '근무기록') s.getRange('A:A').setNumberFormat('@');  // 날짜가 Date로 바뀌지 않게
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function auth_(t) {
  if (TOKEN && t !== TOKEN) throw new Error('unauthorized');
}
function safe_(v) {  // 수식 주입 방지
  v = String(v == null ? '' : v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function rows_(s, width) {
  const n = s.getLastRow();
  return n < 2 ? [] : s.getRange(2, 1, n - 1, width).getDisplayValues();
}
function write_(s, width, rows) {
  const n = s.getLastRow();
  if (n > 1) s.getRange(2, 1, n - 1, s.getMaxColumns()).clearContent();
  if (rows.length) s.getRange(2, 1, rows.length, width).setValues(rows);
}
function readAll_() {
  const staff = rows_(sheet_('직원', H_STAFF), 3).filter(r => r[0])
    .map(r => ({name: r[0], nat: r[1], grp: r[2], memo: ''}));
  const rec = {};
  rows_(sheet_('근무기록', H_REC), 5).forEach(r => {
    if (!r[0] || !r[1]) return;
    (rec[r[0]] = rec[r[0]] || {})[r[1]] = {type: TYPE[r[2]] || 'full', hours: r[3] === '' ? '' : Number(r[3]), memo: r[4]};
  });
  return {staff: staff, rec: rec};
}
function recRows_(date, entries) {
  return Object.keys(entries || {}).map(n => {
    const x = entries[n];
    return [date, n, LABEL[x.type] || '전일', x.hours === '' || x.hours == null ? '' : Number(x.hours), safe_(x.memo)];
  });
}

// 브라우저에서 웹앱 URL을 열면 대시보드(index.html)를 보여준다.
// (token 파라미터가 있는 요청은 외부에 올린 페이지용 JSON API)
function doGet(e) {
  if (e && e.parameter && 'token' in e.parameter) {
    try {
      auth_(e.parameter.token);
      return out_(readAll_());
    } catch (err) {
      return out_({error: String(err.message || err)});
    }
  }
  // 붙여넣기가 중간에 잘렸는지 검사 (각 파일 맨 끝에 END-OF-파일명 표식이 있어야 함)
  const parts = ['index', 'style', 'app1', 'app2'];
  const bad = parts.filter(n => HtmlService.createHtmlOutputFromFile(n).getContent().indexOf('END-OF-' + n) < 0);
  if (bad.length) {
    return HtmlService.createHtmlOutput(
      '<p style="font:16px sans-serif">다음 파일이 끝까지 붙여넣어지지 않았습니다: ' + bad.join(', ') +
      '.html — 해당 파일 내용을 지우고 다시 붙여넣은 뒤 새 버전으로 배포하세요.</p>');
  }
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('인력관리 대시보드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 대시보드(google.script.run)에서 호출하는 함수들
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
function apiRead() {
  return readAll_();
}
function apiCall(b) {
  return locked_(() => apply_(b));
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function apply_(b) {
  if (b.action === 'setStaff') {
    write_(sheet_('직원', H_STAFF), 3, b.staff.map(s => [safe_(s.name), safe_(s.nat), safe_(s.grp)]));
  } else if (b.action === 'setDay') {
    const s = sheet_('근무기록', H_REC);
    const keep = rows_(s, 5).filter(r => r[0] && r[0] !== b.date);
    write_(s, 5, keep.concat(recRows_(b.date, b.entries)));
  } else if (b.action === 'setAll') {
    write_(sheet_('직원', H_STAFF), 3, b.staff.map(s => [safe_(s.name), safe_(s.nat), safe_(s.grp)]));
    let all = [];
    Object.keys(b.rec).sort().forEach(d => { all = all.concat(recRows_(d, b.rec[d])); });
    write_(sheet_('근무기록', H_REC), 5, all);
  } else {
    throw new Error('unknown action');
  }
  return {ok: true};
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    auth_(b.token);
    return out_(locked_(() => apply_(b)));
  } catch (err) {
    return out_({error: String(err.message || err)});
  }
}
