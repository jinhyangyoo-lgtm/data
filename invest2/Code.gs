/**
 * 투자 수익 관리 웹앱
 * 데이터: '투자관리 DB' 스프레드시트 (투자목록 · 조건이력 · 입금일정 · 공휴일 · 과거수익)
 */
const DB_ID = '1k_VsiEwP6IoarVPnikxtAZ8S0ruGqjvJCA74G75bqyQ';
const TZ = 'Asia/Seoul';
const PENDING_MARK = '[계약서 대기]';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('투자 수익 관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

let SS_MEMO_ = null;
function ss_() {
  if (!SS_MEMO_) SS_MEMO_ = SpreadsheetApp.openById(DB_ID);
  return SS_MEMO_;
}

function ymd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function rows_(name) {
  const sh = ss_().getSheetByName(name);
  const last = lastDataRow_(sh);
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return v.filter(function (r) { return r[0] !== '' && r[0] !== null; });
}

/* ===== 속도: 서버 캐시(gzip) =====
 * 데이터는 5분간 캐시한다. 이 앱에서 저장·연장·등록을 하면 캐시를 즉시 비운다.
 * DB 시트를 직접 고쳤을 때는 화면의 '새로고침'(force)으로 바로 반영한다. */
const CACHE_KEY = 'inv2_data';
const CACHE_SEC = 300;
function bust_() {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get(CACHE_KEY + '_n')) || 0;
    const keys = [CACHE_KEY + '_n'];
    for (let i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
    c.removeAll(keys);
  } catch (e) { }
}
function getData(force) {
  const c = CacheService.getScriptCache();
  if (!force) {
    try {
      const n = Number(c.get(CACHE_KEY + '_n')) || 0;
      if (n) {
        const keys = []; for (let i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
        const got = c.getAll(keys);
        let b64 = ''; let ok = true;
        for (let i = 0; i < n; i++) { if (!got[keys[i]]) { ok = false; break; } b64 += got[keys[i]]; }
        if (ok) return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString();
      }
    } catch (e) { }
  }
  const json = build_();
  try {
    const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json')).getBytes());
    const size = 90000, n = Math.ceil(b64.length / size), obj = {};
    for (let i = 0; i < n; i++) obj[CACHE_KEY + '_' + i] = b64.substr(i * size, size);
    obj[CACHE_KEY + '_n'] = String(n);
    c.putAll(obj, CACHE_SEC);
  } catch (e) { }
  return json;
}

/** 시트에서 전체 데이터를 읽어 JSON 문자열로 만든다. */
function build_() {
  const inv = rows_('투자목록').map(function (r) {
    return {
      id: r[0], name: r[1], nick: r[2], investor: r[3], kind: r[4],
      start: ymd_(r[5]), end: ymd_(r[6]), payDay: num_(r[7]), bonus: num_(r[8]) || 0,
      investee: r[9], corpNo: String(r[10] || ''), guarantor: r[11],
      docName: r[12], docUrl: r[13], doc2Url: r[14], memo: r[15],
      principal0: num_(r[16]), deposit: ymd_(r[18])
    };
  });
  const cond = rows_('조건이력').map(function (r) {
    return { id: r[0], from: ymd_(r[1]), principal: num_(r[2]) || 0, rate: num_(r[3]) || 0, note: r[4] };
  });
  const sch = rows_('입금일정').map(function (r) {
    return {
      sid: r[0], id: r[1], type: r[2], due: ymd_(r[3]), changed: ymd_(r[4]),
      manual: num_(r[5]), paid: r[6] === true, actual: num_(r[7]),
      paidAt: r[8] instanceof Date ? Utilities.formatDate(r[8], TZ, 'yyyy-MM-dd HH:mm') : '',
      memo: r[9] || '', payDate: ymd_(r[12]), principal: num_(r[13]), rate: num_(r[14]),
      expected: num_(r[15]) || 0
    };
  });
  const hol = rows_('공휴일').map(function (r) { return { date: ymd_(r[0]), name: r[1] }; });
  const extSh = ss_().getSheetByName('연장이력');
  const ext = extSh ? rows_('연장이력').map(function (r) {
    return {
      id: r[0], extDate: ymd_(r[1]), oldEnd: ymd_(r[2]), newEnd: ymd_(r[3]),
      oldPrincipal: num_(r[4]), newPrincipal: num_(r[5]), repay: num_(r[6]),
      oldRate: num_(r[7]), newRate: num_(r[8]), oldDocName: r[9] || '', oldDocUrl: r[10] || '',
      newDocName: r[11] || '', newDocUrl: r[12] || '',
      at: r[13] instanceof Date ? Utilities.formatDate(r[13], TZ, 'yyyy-MM-dd HH:mm') : '', memo: r[14] || ''
    };
  }) : [];
  const hist = rows_('과거수익').map(function (r) { return { year: Number(r[0]), amount: num_(r[1]) || 0 }; });
  return JSON.stringify({
    inv: inv, cond: cond, sch: sch, hol: hol, hist: hist, ext: ext,
    today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
    dbUrl: ss_().getUrl(),
    ts: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')
  });
}

function findRow_(sh, sid) {
  const ids = sh.getRange(2, 1, sh.getMaxRows() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === sid) return i + 2;
  }
  throw new Error('일정을 찾을 수 없습니다: ' + sid);
}

function toDate_(s) {
  const p = s.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

/**
 * 입금일정 한 줄을 수정한다.
 * patch: { paid: bool, actual: number|null, changed: 'yyyy-MM-dd'|'', memo: string }
 */
function updateSchedule(sid, patch) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = ss_().getSheetByName('입금일정');
    const row = findRow_(sh, sid);
    if ('changed' in patch) sh.getRange(row, 5).setValue(patch.changed ? toDate_(patch.changed) : '');
    if ('paid' in patch) {
      sh.getRange(row, 7).setValue(!!patch.paid);
      sh.getRange(row, 9).setValue(patch.paid ? new Date() : '');
    }
    if ('actual' in patch) {
      sh.getRange(row, 8).setValue(patch.actual === null || patch.actual === '' ? '' : Number(patch.actual));
    }
    if ('memo' in patch) sh.getRange(row, 10).setValue(patch.memo || '');
    SpreadsheetApp.flush();
    const r = sh.getRange(row, 1, 1, 17).getValues()[0];
    return JSON.stringify({
      payDate: ymd_(r[12]), expected: num_(r[15]) || 0,
      paidAt: r[8] instanceof Date ? Utilities.formatDate(r[8], TZ, 'yyyy-MM-dd HH:mm') : ''
    });
  } finally {
    bust_();
    lock.releaseLock();
  }
}

function addMonths_(y, m, k, day) {
  // m: 1~12. k개월 뒤 같은 날(말일 보정)
  const t = new Date(y, m - 1 + k, 1);
  const last = new Date(t.getFullYear(), t.getMonth() + 1, 0).getDate();
  return new Date(t.getFullYear(), t.getMonth(), Math.min(day, last));
}

/**
 * 투자목록을 보고 빠진 일정을 채운다. (from 이전 날짜는 만들지 않음)
 * 장기: 매월 이자(시작 다음 달 ~ 만기월), 만기 추가이자, 원금상환
 * 단기: 원금상환만 (단기 이자는 계약마다 달라 직접 입력)
 * onlyId 를 주면 그 투자만 처리한다. 잠금은 호출하는 쪽에서 건다.
 */
function generateCore_(onlyId, from) {
  const d = JSON.parse(build_());
  const sh = ss_().getSheetByName('입금일정');
  const key = function (id, type, date) { return id + '|' + type + '|' + date; };
  const exist = {};
  d.sch.forEach(function (s) { exist[key(s.id, s.type, s.due)] = true; });
  const hasCond = {};
  d.cond.forEach(function (c) { hasCond[c.id] = true; });
  let maxN = 0;
  d.sch.forEach(function (s) {
    const n = parseInt(String(s.sid).replace(/\D/g, ''), 10);
    if (n > maxN) maxN = n;
  });
  const add = [];
  const missing = [];
  function push(id, type, date) {
    if (date < from) return;
    const ds = Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
    if (exist[key(id, type, ds)]) return;
    exist[key(id, type, ds)] = true;
    maxN++;
    add.push(['S' + ('0000' + maxN).slice(-4), id, type, date]);
  }
  d.inv.forEach(function (x) {
    if (onlyId && x.id !== onlyId) return;
    if (!x.start || !x.end) return;
    const s = toDate_(x.start), e = toDate_(x.end);
    if (e < from) return;
    if (!hasCond[x.id]) missing.push(x.id);
    if (x.kind === '장기' && x.payDay) {
      for (let k = 1; k < 240; k++) {
        const dt = addMonths_(s.getFullYear(), s.getMonth() + 1, k, x.payDay);
        if (dt.getFullYear() * 12 + dt.getMonth() > e.getFullYear() * 12 + e.getMonth()) break;
        push(x.id, '월이자', dt);
      }
      if (x.bonus) push(x.id, '추가이자', e);
    }
    push(x.id, '원금상환', e);
  });
  if (add.length) {
    const startRow = d.sch.length + 2;
    sh.getRange(startRow, 1, add.length, 4).setValues(add);
  }
  return { added: add.length, missing: missing, maxN: maxN };
}

function generateSchedule() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const t = toDate_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
    const r = generateCore_(null, new Date(t.getFullYear(), t.getMonth(), 1));
    return JSON.stringify({ added: r.added, missing: r.missing });
  } finally {
    bust_();
    lock.releaseLock();
  }
}

/** A열 기준 마지막 데이터 행 (배열 수식이 있는 시트는 getLastRow가 빈 줄까지 센다) */
function lastDataRow_(sh) {
  const v = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '' && v[i][0] !== null) return i + 1;
  return 1;
}

function extSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName('연장이력');
  if (!sh) {
    sh = ss.insertSheet('연장이력');
    sh.getRange(1, 1, 1, 15).setValues([['투자ID', '연장일(기존 만기일)', '기존 만기일', '새 만기일', '기존 원금', '연장 원금',
      '상환액', '기존 이율', '새 이율', '기존 계약서명', '기존 계약서링크', '새 계약서명', '새 계약서링크', '처리일시', '메모']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * 만기 연장 처리.
 * p: { id, extDate, newEnd, newPrincipal, newRate(소수), payDay, bonus(소수), kind,
 *      docName, docUrl, memo, shortInterest: [{date, amount}] }
 * - 기존 만기일 원금상환: 원금을 줄이면 차액만 상환으로 남기고, 그대로/늘리면 삭제
 * - 기존 만기 추가이자: 금액을 기존 조건으로 고정
 * - 조건이력: 연장일 다음 날부터 새 원금·이율
 * - 투자목록: 만기일·지급일·추가이자율·구분·계약서 갱신 (기존 계약서는 연장이력에 보관)
 * - 연장 기간의 이자·원금상환 일정 생성
 */
function extendInvestment(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    const invSh = ss.getSheetByName('투자목록');
    const invVals = invSh.getRange(2, 1, Math.max(1, lastDataRow_(invSh) - 1), 18).getValues();
    let ir = -1;
    for (let i = 0; i < invVals.length; i++) if (invVals[i][0] === p.id) { ir = i; break; }
    if (ir < 0) throw new Error('투자를 찾을 수 없습니다: ' + p.id);
    const old = invVals[ir];
    const invRow = ir + 2;
    const oldEnd = ymd_(old[6]);
    const extDate = p.extDate || oldEnd;
    if (!p.newEnd || p.newEnd <= extDate) throw new Error('새 만기일은 연장일 이후여야 합니다.');
    const newPrincipal = Number(p.newPrincipal);
    const newRate = Number(p.newRate);
    if (!(newPrincipal > 0)) throw new Error('연장 원금을 입력해 주세요.');
    if (!(newRate >= 0)) throw new Error('새 이율을 입력해 주세요.');

    // 연장일 기준 기존 조건
    const condSh = ss.getSheetByName('조건이력');
    const condLast = lastDataRow_(condSh);
    const condVals = condLast > 1 ? condSh.getRange(2, 1, condLast - 1, 5).getValues() : [];
    let oc = null;
    condVals.forEach(function (c) {
      const f = ymd_(c[1]);
      if (c[0] === p.id && f <= extDate && (!oc || f >= oc.from)) oc = { from: f, principal: Number(c[2]) || 0, rate: Number(c[3]) || 0 };
    });
    const oldPrincipal = oc ? oc.principal : (Number(old[16]) || 0);
    const oldRate = oc ? oc.rate : 0;
    const oldBonus = Number(old[8]) || 0;
    const repay = Math.max(0, oldPrincipal - newPrincipal);

    // 입금일정 정리 (아래쪽부터 지워야 행 번호가 안 밀린다)
    const schSh = ss.getSheetByName('입금일정');
    const n = lastDataRow_(schSh) - 1;
    const sch = n > 0 ? schSh.getRange(2, 1, n, 10).getValues() : [];
    const del = [];
    for (let i = 0; i < sch.length; i++) {
      const r = sch[i];
      if (r[1] !== p.id || r[6] === true) continue;
      const due = ymd_(r[3]);
      const row = i + 2;
      if (r[2] === '원금상환' && due === oldEnd) {
        if (repay > 0) {
          schSh.getRange(row, 6).setValue(repay);
          schSh.getRange(row, 10).setValue('연장 시 부분상환 (연장 원금 ' + newPrincipal.toLocaleString() + '원)');
        } else del.push(row);
      } else if (r[2] === '추가이자' && due === oldEnd) {
        if (r[5] === '' || r[5] === null) schSh.getRange(row, 6).setValue(Math.round(oldPrincipal * oldBonus));
      } else if (r[2] === '월이자' && due > extDate) {
        del.push(row); // 연장 조건으로 다시 만든다
      }
    }
    del.sort(function (a, b) { return b - a; }).forEach(function (row) { schSh.deleteRow(row); });

    // 조건이력: 연장일 다음 날부터
    const from = toDate_(extDate); from.setDate(from.getDate() + 1);
    const fromS = Utilities.formatDate(from, TZ, 'yyyy-MM-dd');
    const note = '연장 (' + extDate + ' → ' + p.newEnd + ')';
    let updated = false;
    for (let i = 0; i < condVals.length; i++) {
      if (condVals[i][0] === p.id && ymd_(condVals[i][1]) === fromS) {
        condSh.getRange(i + 2, 3, 1, 3).setValues([[newPrincipal, newRate, note]]);
        updated = true; break;
      }
    }
    if (!updated) condSh.getRange(condLast + 1, 1, 1, 5).setValues([[p.id, from, newPrincipal, newRate, note]]);

    // 연장이력
    const ext = extSheet_();
    ext.getRange(lastDataRow_(ext) + 1, 1, 1, 15).setValues([[
      p.id, toDate_(extDate), old[6], toDate_(p.newEnd), oldPrincipal, newPrincipal, repay, oldRate, newRate,
      old[12] || '', old[13] || '', p.docName || '', p.docUrl || '', new Date(), p.memo || ''
    ]]);

    // 투자목록
    if (p.kind) invSh.getRange(invRow, 5).setValue(p.kind);
    invSh.getRange(invRow, 7).setValue(toDate_(p.newEnd));
    if (p.payDay !== undefined) invSh.getRange(invRow, 8).setValue(p.payDay ? Number(p.payDay) : '');
    if (p.bonus !== undefined) invSh.getRange(invRow, 9).setValue(p.bonus ? Number(p.bonus) : '');
    if (p.docUrl) {
      invSh.getRange(invRow, 13).setValue(p.docName || '연장 계약서');
      invSh.getRange(invRow, 14).setValue(p.docUrl);
    }
    const memo = String(old[15] || '');
    const add = Utilities.formatDate(new Date(), TZ, 'yyyy.M.d') + ' 연장: 만기 ' + p.newEnd + ', ' +
      (newRate * 100).toFixed(1) + '%, 원금 ' + newPrincipal.toLocaleString() + '원' + (repay ? ' (' + repay.toLocaleString() + '원 상환)' : '');
    invSh.getRange(invRow, 16).setValue(memo ? memo + ' / ' + add : add);
    SpreadsheetApp.flush();

    // 연장 기간 일정
    const g = generateCore_(p.id, from);
    let maxN = g.maxN, extra = 0;
    const si = (p.shortInterest || []).filter(function (x) { return x.date && Number(x.amount) > 0; });
    if (si.length) {
      const rows = si.map(function (x) { maxN++; return ['S' + ('0000' + maxN).slice(-4), p.id, '단기이자', toDate_(x.date), '', Number(x.amount)]; });
      const startRow = lastDataRow_(schSh) + 1;
      schSh.getRange(startRow, 1, rows.length, 6).setValues(rows);
      extra = rows.length;
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ added: g.added + extra, repay: repay, removed: del.length });
  } finally {
    bust_();
    lock.releaseLock();
  }
}

/**
 * 신규 투자 등록.
 * p: { name, nick, investor, kind, start, end, principal, rate(소수), payDay, bonus(소수),
 *      investee, corpNo, guarantor, docName, docUrl, doc2Url, memo, shortInterest: [{date, amount}] }
 * - 투자목록에 새 투자ID(P숫자)로 추가, 조건이력에 시작일 기준 원금·이율 추가
 * - 이자·원금상환 일정을 시작일부터 만든다
 */
function addInvestment(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const name = String(p.name || '').trim();
    const investor = String(p.investor || '').trim();
    if (!name) throw new Error('프로젝트명을 입력해 주세요.');
    if (!investor) throw new Error('투자자를 입력해 주세요.');
    if (!p.start) p.start = p.deposit; // 계약서 연결 전에는 입금일을 임시 시작일로 쓴다
    if (!p.start) throw new Error('투자 입금일을 입력해 주세요.');
    const pending = !p.end; // 만기일 없이 등록 → 계약서를 연결하면 나머지를 채운다
    if (!pending && p.end <= p.start) throw new Error('만기일은 시작일 이후여야 합니다.');
    const principal = Number(p.principal), rate = Number(p.rate);
    if (!(principal > 0)) throw new Error('투자원금을 입력해 주세요.');
    if (!(rate >= 0)) throw new Error('연이율을 입력해 주세요.');
    const kind = pending ? '' : (p.kind === '단기' ? '단기' : '장기');
    if (kind === '장기' && !(Number(p.payDay) >= 1 && Number(p.payDay) <= 31)) throw new Error('이자 지급일(1~31)을 입력해 주세요.');

    const ss = ss_();
    const invSh = ss.getSheetByName('투자목록');
    const invLast = lastDataRow_(invSh);
    const ids = invLast > 1 ? invSh.getRange(2, 1, invLast - 1, 1).getValues() : [];
    let maxId = 0;
    ids.forEach(function (r) {
      const n = parseInt(String(r[0]).replace(/\D/g, ''), 10);
      if (n > maxId) maxId = n;
    });
    const id = 'P' + ('00' + (maxId + 1)).slice(-2);

    invSh.getRange(invLast + 1, 1, 1, 17).setValues([[
      id, name, p.nick || '', investor, kind, toDate_(p.start), pending ? '' : toDate_(p.end),
      kind === '장기' ? Number(p.payDay) : '', p.bonus ? Number(p.bonus) : '',
      p.investee || '', p.corpNo || '', p.guarantor || '', p.docName || '', p.docUrl || '', p.doc2Url || '',
      pending ? (PENDING_MARK + (p.memo ? ' ' + p.memo : '')) : (p.memo || ''), principal
    ]]);
    invSh.getRange(invLast + 1, 11).setNumberFormat('@').setValue(p.corpNo || '');
    if (p.deposit) invSh.getRange(invLast + 1, 19).setValue(toDate_(p.deposit)); // S열: 투자입금일

    const condSh = ss.getSheetByName('조건이력');
    condSh.getRange(lastDataRow_(condSh) + 1, 1, 1, 5).setValues([[id, toDate_(p.start), principal, rate, '신규 등록']]);
    SpreadsheetApp.flush();

    const g = generateCore_(id, toDate_(p.start));
    let maxN = g.maxN, extra = 0;
    const si = (p.shortInterest || []).filter(function (x) { return x.date && Number(x.amount) > 0; });
    if (si.length) {
      const schSh = ss.getSheetByName('입금일정');
      const rows = si.map(function (x) { maxN++; return ['S' + ('0000' + maxN).slice(-4), id, '단기이자', toDate_(x.date), '', Number(x.amount)]; });
      schSh.getRange(lastDataRow_(schSh) + 1, 1, rows.length, 6).setValues(rows);
      extra = rows.length;
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ id: id, added: g.added + extra });
  } finally {
    bust_();
    lock.releaseLock();
  }
}


/* ================= 계약서 자동 읽기 ================= */

function fileIdFromUrl_(u) {
  u = String(u || '').trim();
  const m = u.match(/\/d\/([\w-]{20,})/) || u.match(/[?&]id=([\w-]{20,})/) || u.match(/^([\w-]{25,})$/);
  if (!m) throw new Error('구글 드라이브 파일 링크가 아닙니다. 드라이브에서 \'링크 복사\'한 주소를 붙여넣어 주세요.');
  return m[1];
}

/** 구글 문서로 변환해 글자를 읽고, 임시 문서는 지운다. */
function docText_(blobOrId, name, ocr) {
  const res = { name: '_계약서읽기_' + name, mimeType: MimeType.GOOGLE_DOCS };
  const opt = { supportsAllDrives: true, fields: 'id' };
  if (ocr) opt.ocrLanguage = 'ko';
  const tmp = typeof blobOrId === 'string' ? Drive.Files.copy(res, blobOrId, opt) : Drive.Files.create(res, blobOrId, opt);
  try {
    return DocumentApp.openById(tmp.id).getBody().getText();
  } finally {
    try { Drive.Files.remove(tmp.id, { supportsAllDrives: true }); }
    catch (e) { try { DriveApp.getFileById(tmp.id).setTrashed(true); } catch (e2) { } }
  }
}

/** 바이트 배열에서 같은 길이의 ASCII 문자열을 바꾼다 (PDF 위치표가 깨지지 않도록 길이 유지). */
function swapBytes_(bytes, from, to) {
  const f = from.split('').map(function (c) { return c.charCodeAt(0); });
  const t = to.split('').map(function (c) { return c.charCodeAt(0); });
  for (let i = 0; i <= bytes.length - f.length; i++) {
    if (bytes[i] !== f[0]) continue;
    let ok = true;
    for (let k = 1; k < f.length; k++) if (bytes[i + k] !== f[k]) { ok = false; break; }
    if (ok) { for (let k = 0; k < t.length; k++) bytes[i + k] = t[k]; i += f.length - 1; }
  }
}

/**
 * 계약서 글자 읽기.
 * - fields: PDF에 입력된 칸(투자금·이름 등)의 글자
 * - ocr: 입력 칸을 숨긴 사본을 스캔 OCR 한 본문 글자 (스캔본은 입력 칸이 있으면 OCR을 건너뛰기 때문)
 */
function ocrText_(fileId) {
  const meta = Drive.Files.get(fileId, { fields: 'id,name,mimeType', supportsAllDrives: true });
  if (meta.mimeType === MimeType.GOOGLE_DOCS) {
    return { name: meta.name, ocr: DocumentApp.openById(fileId).getBody().getText(), fields: '' };
  }
  const fields = docText_(fileId, meta.name, true);
  let ocr = '';
  if (meta.mimeType === 'application/pdf') {
    const bytes = DriveApp.getFileById(fileId).getBlob().getBytes();
    swapBytes_(bytes, '/Annots', '/Xnnots');
    swapBytes_(bytes, '/AcroForm', '/XcroForm');
    ocr = docText_(Utilities.newBlob(bytes, 'application/pdf', meta.name), meta.name, true);
  }
  if (ocr.replace(/\s/g, '').length < 50) { ocr = fields; }
  return { name: meta.name, ocr: ocr, fields: fields === ocr ? '' : fields };
}

function pad2_(n) { return ('0' + n).slice(-2); }

/** 'yyyy년 m월 d일' / 'yyyy. m. d' 형태의 날짜를 위치와 함께 모두 찾는다. */
function datesIn_(s) {
  const re = /(20\d{2})\s*(?:년|\.|-|\/)\s*(\d{1,2})\s*(?:월|\.|-|\/)\s*(\d{1,2})\s*일?/g;
  const out = [];
  let m;
  while ((m = re.exec(s))) {
    const mo = Number(m[2]), d = Number(m[3]);
    if (mo < 1 || mo > 12 || d < 1 || d > 31) continue;
    out.push({ d: m[1] + '-' + pad2_(mo) + '-' + pad2_(d), i: m.index, j: m.index + m[0].length });
  }
  return out;
}

function cleanName_(s) {
  return String(s || '').replace(/\s+/g, ' ').replace(/^[\s:"'“”]+|[\s:"'“”,]+$/g, '').trim();
}

/** 계약서 글자에서 투자 조건을 뽑는다. 못 찾은 항목은 비워 둔다. */
function parseContract_(raw, fileName, fieldsRaw) {
  const fl = String(fieldsRaw || '');
  const t = String(raw || '').replace(/\r/g, '').replace(/[“”″＂]/g, '"').replace(/：/g, ':')
    .replace(/[ \t ]+/g, ' ');
  const r = { warn: [] };

  // 제2조(투자조건) 구간
  const a = t.search(/제\s*2\s*조/), b = t.search(/제\s*3\s*조/);
  const sec = a >= 0 ? t.slice(a, b > a ? b : a + 1500) : t;
  const ds = datesIn_(sec);

  // 투자일(시작일)
  let m = sec.match(/투자를\s*((?:20\d{2})\s*(?:년|\.)\s*\d{1,2}\s*(?:월|\.)\s*\d{1,2}\s*일?)/);
  if (m) r.start = datesIn_(m[1])[0].d;
  else if (a >= 0) {
    const k = sec.search(/이하\s*"?\s*투자일/);
    const before = ds.filter(function (x) { return k >= 0 && x.j <= k; });
    if (before.length) r.start = before[before.length - 1].d;
    else if (ds.length) r.start = ds[0].d;
  }

  // 투자금
  m = sec.match(/[₩W\\￦]\s*([\d,.]{7,})/) || t.match(/[₩W\\￦]\s*([\d,.]{7,})/);
  if (m) r.principal = Number(m[1].replace(/[,.]/g, ''));
  else {
    const nums = ((fl + '\n' + sec).match(/\d{1,3}(?:,\d{3}){2,}/g) || []).map(function (x) { return Number(x.replace(/,/g, '')); });
    if (nums.length) r.principal = Math.max.apply(null, nums);
  }

  // 연이율 · 매월 지급일
  m = sec.match(/연\s*(\d{1,2}(?:\.\d+)?)\s*%/);
  if (m) r.rate = Number(m[1]);
  m = sec.match(/매월\s*(\d{1,2})\s*일/);
  if (m) r.payDay = Number(m[1]);

  // 만기일: 투자조건(제2조)에 나온 가장 늦은 날짜 (원금 지급일)
  if (a >= 0 && ds.length) r.end = ds.map(function (x) { return x.d; }).sort().pop();
  if (r.end && r.start && r.end <= r.start) r.end = '';

  // 퍼센트 + 지급일 (단기 이자 / 투자수수료)
  const pays = [];
  const pre = /(\d{1,2}(?:\.\d+)?)\s*%/g;
  const pm = [];
  while ((m = pre.exec(sec))) pm.push({ v: Number(m[1]), i: m.index, j: m.index + m[0].length });
  pm.forEach(function (p, n) {
    const head = sec.slice(Math.max(0, p.i - 6), p.i);
    if (/연\s*$/.test(head)) return; // 연이율
    let seg = sec.slice(p.j, n + 1 < pm.length ? pm[n + 1].i : sec.length);
    const cut = seg.search(/후취|지급\s*시기|입금|계좌|\n\s*\*/);
    if (cut >= 0) seg = seg.slice(0, cut + (seg.slice(cut, cut + 2) === '후취' ? 0 : 0));
    const seen = {};
    datesIn_(seg).forEach(function (x) {
      if (x.d === r.start || seen[x.d]) return;
      seen[x.d] = true;
      pays.push({ date: x.d, pct: p.v, fee: /원금의\s*$/.test(sec.slice(Math.max(0, p.i - 12), p.i)) });
    });
  });

  r.kind = (r.start && r.end) ? (r.end < (Number(r.start.slice(0, 4)) + 1) + r.start.slice(4) ? '단기' : '장기') : '';
  r.extra = []; r.bonus = 0;
  pays.forEach(function (x) {
    if (!r.kind) return;
    if (r.kind === '장기') {
      if (!x.fee) return; // 장기는 '투자원금의 n%' 형태만 별도 지급으로 본다
      if (x.date === r.end) r.bonus = x.pct;
      else r.extra.push({ date: x.date, pct: x.pct, memo: '투자수수료 ' + x.pct + '%' });
    } else {
      r.extra.push({ date: x.date, pct: x.pct, memo: (x.fee ? '투자수수료 ' : '단기 이자 ') + x.pct + '%' });
    }
  });

  // 피투자자 · 연대보증인 (서명란: 이름 (법인등록번호 : 000000-0000000))
  function party(label) {
    const k = t.lastIndexOf(label);
    if (k < 0) return null;
    const s = t.slice(k + label.length, k + label.length + 220);
    const mm = s.match(/([^\n(]{2,40}?)\s*\(\s*법\s*인\s*등\s*록\s*번\s*호\s*:?\s*(\d{6})\s*-\s*(\d{7})/);
    if (mm) return { name: cleanName_(mm[1]), no: mm[2] + '-' + mm[3] };
    const lines = s.split('\n').map(cleanName_).filter(function (x) { return x.length >= 2 && !/\d{4}\s*년|^\)|^\(/.test(x); });
    return lines.length ? { name: lines[0], no: '' } : null;
  }
  const g = party('피투자자)');
  if (g) { r.investee = g.name; r.corpNo = g.no; }
  m = t.match(/피투자자로서\s*(.{2,30}?)\s*\(\s*이하/);
  if (m && (!g || !g.no)) r.investee = cleanName_(m[1]);
  if (!r.corpNo && r.investee) { // 서명란에서 피투자자 이름 뒤의 법인등록번호
    const k = t.lastIndexOf(r.investee);
    const mm = k >= 0 ? t.slice(k, k + 80).match(/법\s*인\s*등\s*록\s*번\s*호\s*:?\s*(\d{6})\s*-\s*(\d{7})/) : null;
    if (mm) r.corpNo = mm[1] + '-' + mm[2];
  }
  const bg = party('연대보증인)');
  if (bg) r.guarantor = bg.name + (bg.no ? ' (' + bg.no + ')' : '');
  else {
    m = t.match(/연대보증인으로서\s*(.{2,30}?)\s*\(\s*이하/);
    if (m) r.guarantor = cleanName_(m[1]);
  }

  // 투자자 성명
  const ek = t.lastIndexOf('(투자자)');
  m = (ek >= 0 ? t.slice(ek, ek + 120) : t).match(/성\s*명\s*:?[ \t]*([가-힣]{2,4})(?![가-힣])/);
  if (m && !/주민|번호|주소/.test(m[1])) r.investorName = m[1];
  if (!r.investorName && fl) {
    const nm = fl.split('\n').map(function (x) { return x.trim(); }).filter(function (x) { return /^[가-힣]{2,4}$/.test(x) && !/만원|억원|천원/.test(x); })[0];
    if (nm) r.investorName = nm;
  }

  // 계약서 이름 (파일명 첫 부분 + 투자계약서)
  const base = String(fileName || '').replace(/\.pdf$/i, '').replace(/\(\d{9,11}\)/g, '');
  const first = cleanName_(base.split('_')[0]).replace(/^[^0-9A-Za-z가-힣]+/, '');
  if (first) r.docName = first + (/계약서|확인서/.test(first) ? '' : ' 투자계약서');

  ['start', 'end'].forEach(function (k) { if (!r[k]) r.warn.push(k); });
  return r;
}

/** 웹앱: 계약서 링크를 읽어 투자 조건을 돌려준다 (저장은 하지 않음). */
function readContract(url) {
  const id = fileIdFromUrl_(url);
  let o = ocrText_(id);
  let r = parseContract_(o.ocr, o.name, o.fields);
  if (r.warn.length) { // 스캔 OCR이 가끔 일부 쪽을 빠뜨려서 한 번 더 읽어 본다
    const o2 = ocrText_(id), r2 = parseContract_(o2.ocr, o2.name, o2.fields);
    if (r2.warn.length < r.warn.length) { o = o2; r = r2; }
  }
  r.fileName = o.name;
  r.url = 'https://drive.google.com/file/d/' + id + '/view';
  r.text = (o.ocr + (o.fields ? '\n\n[입력 칸]\n' + o.fields : '')).slice(0, 5000);
  return JSON.stringify(r);
}

/**
 * 계약서 내용 적용.
 * p: { id, docUrl, docName, investee, corpNo, guarantor, start, end, kind, payDay, bonus(소수), rate(소수),
 *      extra: [{date, amount, memo}] }
 * - 아직 입금일정이 없는 투자(계약서 대기)는 조건·만기를 채우고 일정을 만든다.
 * - 이미 일정이 있는 투자는 피투자자·보증인·계약서 정보만 채운다.
 */
function applyContract(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    const invSh = ss.getSheetByName('투자목록');
    const invVals = invSh.getRange(2, 1, Math.max(1, lastDataRow_(invSh) - 1), 17).getValues();
    let ir = -1;
    for (let i = 0; i < invVals.length; i++) if (invVals[i][0] === p.id) { ir = i; break; }
    if (ir < 0) throw new Error('투자를 찾을 수 없습니다: ' + p.id);
    const old = invVals[ir], row = ir + 2;

    const schSh = ss.getSheetByName('입금일정');
    const nS = lastDataRow_(schSh) - 1;
    const hasSch = nS > 0 && schSh.getRange(2, 2, nS, 1).getValues().some(function (r) { return r[0] === p.id; });

    // 정보 칸 (빈 값으로 덮어쓰지 않음)
    if (p.investee) invSh.getRange(row, 10).setValue(p.investee);
    if (p.corpNo) invSh.getRange(row, 11).setNumberFormat('@').setValue(p.corpNo);
    if (p.guarantor) invSh.getRange(row, 12).setValue(p.guarantor);
    if (p.docUrl) {
      if (old[13] && old[13] !== p.docUrl) invSh.getRange(row, 15).setValue(p.docUrl); // 기존 계약서가 있으면 추가서류로
      else { invSh.getRange(row, 13).setValue(p.docName || '투자계약서'); invSh.getRange(row, 14).setValue(p.docUrl); }
    }
    const memo = String(old[15] || '').replace(PENDING_MARK, '').trim();

    if (hasSch) {
      invSh.getRange(row, 16).setValue(memo);
      SpreadsheetApp.flush();
      return JSON.stringify({ mode: 'info', added: 0 });
    }

    if (!p.start || !p.end || p.end <= p.start) throw new Error('시작일·만기일을 확인해 주세요.');
    const kind = p.kind === '단기' ? '단기' : '장기';
    if (kind === '장기' && !(Number(p.payDay) >= 1 && Number(p.payDay) <= 31)) throw new Error('이자 지급일(1~31)을 확인해 주세요.');
    invSh.getRange(row, 5, 1, 5).setValues([[kind, toDate_(p.start), toDate_(p.end),
      kind === '장기' ? Number(p.payDay) : '', kind === '장기' && p.bonus ? Number(p.bonus) : '']]);
    invSh.getRange(row, 16).setValue(memo);

    // 조건이력: 이 투자의 첫 줄을 시작일·이율로 맞춘다
    const condSh = ss.getSheetByName('조건이력');
    const cl = lastDataRow_(condSh);
    const cv = cl > 1 ? condSh.getRange(2, 1, cl - 1, 4).getValues() : [];
    let ci = -1;
    for (let i = 0; i < cv.length; i++) if (cv[i][0] === p.id) { ci = i; break; }
    const principal = Number(old[16]) || 0;
    const rate = (p.rate !== undefined && p.rate !== null && p.rate !== '') ? Number(p.rate) : (ci >= 0 ? Number(cv[ci][3]) : 0);
    if (ci >= 0) condSh.getRange(ci + 2, 2, 1, 3).setValues([[toDate_(p.start), Number(cv[ci][2]) || principal, rate]]);
    else condSh.getRange(cl + 1, 1, 1, 5).setValues([[p.id, toDate_(p.start), principal, rate, '계약서 등록']]);
    SpreadsheetApp.flush();

    const g = generateCore_(p.id, toDate_(p.start));
    let maxN = g.maxN, extra = 0;
    const ex = (p.extra || []).filter(function (x) { return x.date && Number(x.amount) > 0; });
    if (ex.length) {
      const rows = ex.map(function (x) { maxN++; return ['S' + ('0000' + maxN).slice(-4), p.id, '단기이자', toDate_(x.date), '', Number(x.amount), false, '', '', x.memo || '']; });
      schSh.getRange(lastDataRow_(schSh) + 1, 1, rows.length, 10).setValues(rows);
      extra = rows.length;
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ mode: 'full', added: g.added + extra });
  } finally {
    bust_();
    lock.releaseLock();
  }
}
