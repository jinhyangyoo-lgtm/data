/**
 * 인력관리 대시보드 — Google Sheets 백엔드 (Apps Script 웹앱)
 * 데이터는 '데이터' 탭에 문서(JSON) 단위로 저장된다: staff / cfg / m:YYYY-MM(근무기록) / p:YYYY-MM(급여지급)
 * 변경된 문서만 한 번에 저장하므로 입력이 빠르다. 기존 '직원'·'근무기록' 탭이 있으면 처음 한 번 자동으로 가져온다.
 */
const PARTS = ['index', 'style1', 'style2', 'app1', 'app2', 'app3', 'app4', 'app5', 'app6', 'app7', 'app8'];  // build.py가 자동으로 채움
const SHEET = '데이터';

function doGet() {
  // 붙여넣기가 중간에 잘렸는지 검사 (각 파일 맨 끝에 END-OF-파일명 표식이 있어야 함)
  const bad = PARTS.filter(n => HtmlService.createHtmlOutputFromFile(n).getContent().indexOf('END-OF-' + n) < 0);
  if (bad.length) {
    return HtmlService.createHtmlOutput(
      '<p style="font:16px sans-serif">다음 파일이 끝까지 붙여넣어지지 않았습니다: ' + bad.join(', ') +
      '.html — 해당 파일 내용을 지우고 다시 붙여넣은 뒤 새 버전으로 배포하세요.</p>');
  }
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('인력관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
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

function store_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.getRange(1, 1, 1, 3).setValues([['키', '값(JSON)', '수정시각']]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.getRange('A:C').setNumberFormat('@');
    s.setColumnWidth(2, 480);
    migrate_(s);
  }
  return s;
}

// 예전 버전('직원', '근무기록', '일별메모', '급여지급' 탭)의 데이터를 새 형식으로 한 번 가져온다. (원본 탭은 그대로 둔다)
function migrate_(s) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const get = n => { const t = ss.getSheetByName(n); return t ? t.getDataRange().getDisplayValues().slice(1) : []; };
  const docs = legacyDocs_(get('직원'), get('근무기록'), get('일별메모'), get('급여지급'));
  const now = new Date().toISOString();
  const rows = Object.keys(docs).map(k => [k, JSON.stringify(docs[k]), now]);
  if (rows.length) s.getRange(2, 1, rows.length, 3).setValues(rows);
}

// 순수 함수: 예전 탭의 행들 → {키: 문서}
function legacyDocs_(staffRows, workRows, memoRows, payRows) {
  const T = {'전일': 'full', '반일': 'half', '시간제': 'part', '시간': 'part'};
  const docs = {};
  const staff = staffRows.filter(r => r[0]).map(r => ({name: r[0], nat: r[1], grp: r[2], active: true}));
  if (!staff.length) return docs;
  docs['staff'] = staff;
  const mon = ym => docs['m:' + ym] || (docs['m:' + ym] = {rec: {}, note: {}, off: {}});
  workRows.forEach(r => {
    if (!r[0] || !r[1]) return;
    (mon(r[0].slice(0, 7)).rec[r[0]] = mon(r[0].slice(0, 7)).rec[r[0]] || {})[r[1]] =
      {t: T[r[2]] || 'full', h: r[3] === '' || r[3] == null ? '' : Number(r[3]), m: r[4] || ''};
  });
  memoRows.forEach(r => {
    const d = r[0], t = String(r[1] || '').trim();
    if (!d || !t) return;
    if (t === '휴무') mon(d.slice(0, 7)).off[d] = true;
    else mon(d.slice(0, 7)).note[d] = t;
  });
  payRows.forEach(r => {   // 기간시작, 기간종료, 이름, 근무일수, 지급액, 지급일, 메모
    if (!r[0] || !r[2] || !r[5]) return;
    const ym = r[0].slice(0, 7), per = r[0].slice(8, 10) === '01' ? 'p1' : 'p2';
    const doc = docs['p:' + ym] || (docs['p:' + ym] = {});
    (doc[per] = doc[per] || {})[r[2]] = {paid: true, paidAt: r[5], fixed: Number(String(r[4]).replace(/[^0-9.-]/g, '')) || 0};
  });
  return docs;
}

// 화면에서 호출 (google.script.run)
function apiLoad() {
  const s = store_(), n = s.getLastRow(), docs = {};
  if (n > 1) {
    s.getRange(2, 1, n - 1, 2).getValues().forEach(r => { if (r[0] && r[1]) docs[r[0]] = String(r[1]); });
  }
  return {docs: docs};
}

function apiSave(changes) {
  return locked_(() => {
    const s = store_(), n = s.getLastRow(), idx = {};
    if (n > 1) s.getRange(2, 1, n - 1, 1).getValues().forEach((r, i) => { idx[r[0]] = i + 2; });
    const now = new Date().toISOString(), add = [];
    Object.keys(changes).forEach(k => {
      const v = changes[k];
      if (v && v.length > 49000) throw new Error('데이터가 너무 큽니다: ' + k);
      if (idx[k]) s.getRange(idx[k], 2, 1, 2).setValues([[v, now]]);
      else add.push([k, v, now]);
    });
    if (add.length) s.getRange(n + 1, 1, add.length, 3).setValues(add);
    return {ok: true};
  });
}

// 보기 좋은 표로 내보내기 (탭 이름 → 2차원 배열)
function apiExport(sheets) {
  return locked_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(sheets).forEach(name => {
      const rows = sheets[name];
      const s = ss.getSheetByName(name) || ss.insertSheet(name);
      s.clear();
      s.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      s.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
      s.setFrozenRows(1);
    });
    return {ok: true};
  });
}
