/**
 * 투자수익관리 (고속 조회 버전) — 기존 '투자관리 DB' 시트를 읽기 전용으로 사용한다.
 * 기존 페이지는 그대로 두고, 이 웹앱을 별도 배포하여 사용한다.
 * 속도: ① 4개 탭을 getDisplayValues 한 번씩만 읽음 ② 서버 캐시(gzip) ③ 브라우저 localStorage 선표시 후 백그라운드 갱신
 */
const SS_ID = '1k_VsiEwP6IoarVPnikxtAZ8S0ruGqjvJCA74G75bqyQ';
const CACHE_KEY = 'inv_v1';
const CACHE_SEC = 300;           // 5분. 화면의 '새로고침'은 캐시를 무시한다.
const SCH_COLS = [0, 1, 2, 3, 4, 5, 6, 7, 9, 12, 13, 14, 15, 16]; // 입금일정에서 쓰는 열만 전송

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('투자수익관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function getData(force) {
  const cache = CacheService.getScriptCache();
  if (!force) {
    const hit = readCache_(cache);
    if (hit) return hit;
  }
  const ss = SpreadsheetApp.openById(SS_ID);
  const read = (name, cols) => {
    const s = ss.getSheetByName(name);
    if (!s || s.getLastRow() < 2) return [];
    const lastCol = cols ? Math.max.apply(null, cols) + 1 : s.getLastColumn();
    const rows = s.getRange(2, 1, s.getLastRow() - 1, lastCol).getDisplayValues();
    return rows.filter(r => String(r[0]).trim() !== '')
      .map(r => cols ? cols.map(i => r[i]) : r);
  };
  const data = {
    inv: read('투자목록'),
    cond: read('조건이력'),
    sch: read('입금일정', SCH_COLS),
    ext: read('연장이력'),
    ts: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
  };
  writeCache_(cache, data);
  return data;
}

function writeCache_(cache, data) {
  try {
    const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(data), 'application/json')).getBytes());
    const size = 90000, n = Math.ceil(b64.length / size), obj = {};
    for (let i = 0; i < n; i++) obj[CACHE_KEY + '_' + i] = b64.substr(i * size, size);
    obj[CACHE_KEY + '_n'] = String(n);
    cache.putAll(obj, CACHE_SEC);
  } catch (e) { /* 캐시 실패는 무시 */ }
}
function readCache_(cache) {
  try {
    const n = Number(cache.get(CACHE_KEY + '_n'));
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
    const got = cache.getAll(keys);
    let b64 = ''; for (let i = 0; i < n; i++) { if (!got[keys[i]]) return null; b64 += got[keys[i]]; }
    const blob = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip'));
    return JSON.parse(blob.getDataAsString());
  } catch (e) { return null; }
}
