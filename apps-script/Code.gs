/**
 * 특별실 예약 - Apps Script API
 *
 * 화면(HTML)은 별도 웹 주소에 두고, 이 스크립트는 JSON만 주고받는다.
 * 일반 교사는 로그인 없이 소속·이름만 적어 예약한다. 관리자 기능만 이름+PIN(사용자 시트)으로 보호한다.
 *
 *  - GET  ?action=config      → 공간/교시/학기/수업잠금/예외일/예약 목록
 *  - POST login               → (관리자) 이름+PIN 확인
 *  - POST book                → 예약 (소속, 이름, 공간, 칸 목록, 목적)
 *  - POST cancel              → 예약 취소 (본인: 소속+이름 일치 / 관리자: 이름+PIN)
 *  - POST setLocks            → (관리자) 한 공간·학기의 수업 잠금 전체 저장
 *  - POST setPeriods          → (관리자) 교시표 저장
 *  - POST setExceptions       → (관리자) 예외일 저장
 * POST 본문은 Content-Type: text/plain 으로 보낸다(브라우저 사전 요청 회피).
 */

const MAX_ADVANCE_DAYS = 60; // 오늘부터 며칠 앞까지 예약할 수 있는지
const MAX_SLOTS = 120;       // 한 번에 보낼 수 있는 칸 수(반복 예약 포함)
const TZ = 'Asia/Seoul';

const TABLES = {
  rooms:      { name: '공간',     cols: ['공간ID', '이름', '위치', '사용여부'] },
  periods:    { name: '교시',     cols: ['교시', '시작', '종료'] },
  terms:      { name: '학기',     cols: ['학기ID', '이름', '시작일', '종료일'] },
  locks:      { name: '수업잠금', cols: ['공간ID', '학기ID', '요일', '교시', '내용'] },
  users:      { name: '사용자',   cols: ['이름', 'PIN', '권한'] },
  exceptions: { name: '예외일',   cols: ['날짜', '공간ID', '사유'] },
  bookings:   { name: '예약',     cols: ['예약ID', '공간ID', '날짜', '교시', '소속', '예약자', '목적', '상태', '생성시각', '묶음ID'] },
};
// 예약 시트 열 번호(0부터)
const B = { ID: 0, ROOM: 1, DATE: 2, PERIOD: 3, AFFIL: 4, NAME: 5, PURPOSE: 6, STATUS: 7, CREATED: 8, SERIES: 9 };

/** 스프레드시트에 연결된 스크립트에서 한 번 실행: 시트와 머리글, 예시 데이터를 만든다. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABLES).forEach(function (key) {
    const t = TABLES[key];
    let sh = ss.getSheetByName(t.name);
    if (!sh) sh = ss.insertSheet(t.name);
    sh.getRange(1, 1, sh.getMaxRows(), t.cols.length).setNumberFormat('@'); // 날짜·시간·PIN이 자동 변환되지 않게 문자로 고정
    sh.getRange(1, 1, 1, t.cols.length).setValues([t.cols]).setFontWeight('bold').setBackground('#eeeeee');
    sh.setFrozenRows(1);
  });
  seed_('rooms', [['R1', '과학실', '본관 3층', 'Y'], ['R2', '컴퓨터실', '별관 2층', 'Y']]);
  seed_('periods', [
    ['1', '08:50', '09:35'], ['2', '09:45', '10:30'], ['3', '10:40', '11:25'], ['4', '11:35', '12:20'],
    ['5', '13:20', '14:05'], ['6', '14:15', '15:00'], ['7', '15:10', '15:55'],
  ]);
  seed_('terms', [['T1', '2026학년도 2학기', '2026-08-17', '2026-12-31']]);
  seed_('users', [['관리자', '0000', '관리자']]); // PIN은 반드시 바꿀 것
  const dflt = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (dflt && ss.getSheets().length > 1) ss.deleteSheet(dflt);
}

function seed_(key, rows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABLES[key].name);
  if (sh.getLastRow() > 1) return; // 이미 데이터가 있으면 건드리지 않음
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

/* ---------- 시트 읽기/쓰기 ---------- */

function read_(key) {
  const t = TABLES[key];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, t.cols.length).getDisplayValues();
  return values
    .filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); })
    .map(function (r) {
      const o = {};
      t.cols.forEach(function (c, i) { o[c] = String(r[i]).trim(); });
      return o;
    });
}

function replaceAll_(key, rows) {
  const t = TABLES[key];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(t.name);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, t.cols.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, t.cols.length).setValues(rows);
}

/* ---------- 관리자 확인 ---------- */

function findAdmin_(name, pin) {
  name = String(name || '').trim();
  pin = String(pin || '').trim();
  if (!name || !pin) return null;
  const hit = read_('users').filter(function (u) { return u['이름'] === name && u['PIN'] === pin && u['권한'] === '관리자'; })[0];
  return hit ? { name: hit['이름'], role: 'admin' } : null;
}

/* ---------- 요청 처리 ---------- */

function doGet(e) {
  return respond_(handle_((e && e.parameter) || {}));
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) { return respond_({ ok: false, error: '요청 형식이 올바르지 않습니다.' }); }
  return respond_(handle_(req));
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(req) {
  try {
    switch (req.action) {
      case 'config':        return { ok: true, data: getConfig_() };
      case 'login':         return login_(req);
      case 'book':          return book_(req);
      case 'cancel':        return cancel_(req);
      case 'setLocks':      return setLocks_(req);
      case 'setPeriods':    return setPeriods_(req);
      case 'setExceptions': return setExceptions_(req);
      default:              return { ok: true, message: 'room-booking api' };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

function getConfig_() {
  return {
    rooms: read_('rooms').filter(function (r) { return r['사용여부'] !== 'N'; }),
    periods: read_('periods'),
    terms: read_('terms'),
    locks: read_('locks'),
    exceptions: read_('exceptions'),
    bookings: read_('bookings').filter(function (b) { return b['상태'] === '유효' && b['날짜'] >= kstNow_().date; }),
    settings: { maxAdvanceDays: MAX_ADVANCE_DAYS },
  };
}

function login_(req) {
  const u = findAdmin_(req.name, req.pin);
  return u ? { ok: true, user: u } : { ok: false, error: '관리자 이름 또는 PIN이 맞지 않습니다.' };
}

/* ---------- 날짜 도우미 ---------- */

function kstNow_() {
  const n = new Date();
  return { date: Utilities.formatDate(n, TZ, 'yyyy-MM-dd'), hm: Utilities.formatDate(n, TZ, 'HH:mm') };
}

function addDays_(dateStr, n) {
  return new Date(new Date(dateStr + 'T00:00:00Z').getTime() + n * 86400000).toISOString().slice(0, 10);
}

function weekday_(dateStr) { return new Date(dateStr + 'T00:00:00Z').getUTCDay(); } // 0=일 … 6=토

function label_(p) { return /^\d+$/.test(p) ? p + '교시' : p; }

/* ---------- 예약 ---------- */

/**
 * 여러 칸을 한 번에 예약한다. 로그인은 없고 소속·이름을 그대로 기록한다.
 * slots: [{date, period, rep?, series?}]
 *  - rep가 없는 칸(직접 고른 칸)이 하나라도 안 되면 아무것도 저장하지 않는다.
 *  - rep:true 칸(반복으로 늘린 칸)은 안 되는 날만 건너뛰고 결과에 알려 준다.
 */
function book_(req) {
  const affil = String(req.affil || '').trim().slice(0, 20);
  const teacher = String(req.teacher || '').trim().slice(0, 20);
  const roomId = String(req.roomId || '');
  const purpose = String(req.purpose || '').trim().slice(0, 60);
  const slots = req.slots || [];
  if (!affil || !teacher) return { ok: false, error: '소속과 이름을 입력하세요.' };
  if (!roomId || !slots.length) return { ok: false, error: '예약할 칸을 선택하세요.' };
  if (!purpose) return { ok: false, error: '사용 목적을 입력하세요.' };
  if (slots.length > MAX_SLOTS) return { ok: false, error: '한 번에 ' + MAX_SLOTS + '칸까지 예약할 수 있습니다.' };
  if (!read_('rooms').some(function (r) { return r['공간ID'] === roomId && r['사용여부'] !== 'N'; })) return { ok: false, error: '사용할 수 없는 공간입니다.' };

  const periods = read_('periods'), terms = read_('terms'), locks = read_('locks'), exceptions = read_('exceptions');
  const now = kstNow_();
  const lastDay = addDays_(now.date, MAX_ADVANCE_DAYS);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const taken = {};
    read_('bookings').forEach(function (b) { if (b['상태'] === '유효' && b['공간ID'] === roomId) taken[b['날짜'] + '|' + b['교시']] = true; });

    // 칸 하나를 검사해서 문제가 있으면 이유를 돌려준다.
    const problem = function (date, p) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return '날짜 형식이 올바르지 않습니다.';
      const wd = weekday_(date);
      if (wd < 1 || wd > 5) return '평일만 예약할 수 있습니다.';
      const per = periods.filter(function (x) { return x['교시'] === p; })[0];
      if (!per) return '없는 교시입니다.';
      if (date < now.date || (date === now.date && per['종료'] <= now.hm)) return '지난 시간은 예약할 수 없습니다.';
      if (date > lastDay) return '오늘부터 ' + MAX_ADVANCE_DAYS + '일 이내만 예약할 수 있습니다.';
      const term = terms.filter(function (t) { return t['시작일'] <= date && date <= t['종료일']; })[0];
      if (!term) return '학기 기간이 아닙니다.';
      const noClass = exceptions.some(function (e) { return e['날짜'] === date && (!e['공간ID'] || e['공간ID'] === roomId); });
      if (!noClass && locks.some(function (l) { return l['공간ID'] === roomId && l['학기ID'] === term['학기ID'] && l['요일'] === String(wd) && l['교시'] === p; })) {
        return '수업 시간입니다.';
      }
      if (taken[date + '|' + p]) return '이미 예약되어 있습니다.';
      return null;
    };

    const rows = [], skipped = [];
    const groupIds = {}; // 같은 행사로 보낸 칸(같은 라벨)끼리 서버가 만든 묶음ID를 붙인다
    for (let i = 0; i < slots.length; i++) {
      const date = String(slots[i].date), p = String(slots[i].period);
      const label = String(slots[i].series || '').slice(0, 40);
      const series = label ? (groupIds[label] || (groupIds[label] = Utilities.getUuid().slice(0, 8))) : '';
      const tag = date.slice(5).replace('-', '/') + ' ' + label_(p);
      const why = problem(date, p);
      if (why) {
        if (slots[i].rep) { skipped.push(tag + ' (' + why.replace(/\.$/, '') + ')'); continue; }
        return { ok: false, error: tag + ': ' + why };
      }
      taken[date + '|' + p] = true;
      rows.push([Utilities.getUuid().slice(0, 8), roomId, date, p, affil, teacher, purpose, '유효', now.date + ' ' + now.hm, series]);
    }
    if (!rows.length) return { ok: false, error: '예약할 수 있는 칸이 없습니다.' };
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABLES.bookings.name);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, count: rows.length, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 예약을 취소 상태로 바꾼다(행은 지우지 않는다).
 * 본인 확인은 예약할 때 적은 소속·이름이 똑같은지로 한다. 관리자는 이름+PIN을 함께 보내면 모두 취소할 수 있다.
 * scope: 'one'(이 칸만, 기본) / 'day'(같은 행사의 그날 칸 모두) / 'all'(같은 행사의 오늘 이후 칸 모두)
 */
function cancel_(req) {
  const admin = findAdmin_(req.name, req.pin);
  const affil = String(req.affil || '').trim(), teacher = String(req.teacher || '').trim();
  if (!admin && (!affil || !teacher)) return { ok: false, error: '소속과 이름을 입력하세요.' };
  const id = String(req.bookingId || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABLES.bookings.name);
    if (sh.getLastRow() < 2) return { ok: false, error: '예약을 찾을 수 없습니다.' };
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, TABLES.bookings.cols.length).getValues();
    let at = -1;
    for (let i = 0; i < vals.length; i++) { if (String(vals[i][B.ID]) === id) { at = i; break; } }
    if (at < 0) return { ok: false, error: '예약을 찾을 수 없습니다.' };
    if (String(vals[at][B.STATUS]) !== '유효') return { ok: false, error: '이미 취소된 예약입니다.' };
    if (!admin && (String(vals[at][B.AFFIL]) !== affil || String(vals[at][B.NAME]) !== teacher)) {
      return { ok: false, error: '본인이 한 예약만 취소할 수 있습니다. 예약할 때 쓴 소속·이름을 똑같이 입력하세요.' };
    }
    const today = kstNow_().date;
    let targets = [at];
    const sid = String(vals[at][B.SERIES]);
    if ((req.scope === 'all' || req.scope === 'day') && sid) {
      targets = [];
      for (let j = 0; j < vals.length; j++) {
        const mine = String(vals[j][B.AFFIL]) === String(vals[at][B.AFFIL]) && String(vals[j][B.NAME]) === String(vals[at][B.NAME]);
        if (String(vals[j][B.SERIES]) === sid && mine && String(vals[j][B.STATUS]) === '유효' && String(vals[j][B.DATE]) >= today &&
            (req.scope === 'all' || String(vals[j][B.DATE]) === String(vals[at][B.DATE]))) targets.push(j);
      }
    }
    if (!admin) targets = targets.filter(function (j) { return String(vals[j][B.DATE]) >= today; });
    if (!targets.length) return { ok: false, error: '지난 예약은 취소할 수 없습니다.' };
    targets.forEach(function (j) { sh.getRange(j + 2, B.STATUS + 1).setValue('취소'); });
    return { ok: true, count: targets.length };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 관리자 기능 ---------- */

function requireAdmin_(req) {
  const u = findAdmin_(req.name, req.pin);
  return u ? null : { ok: false, error: '관리자 로그인이 필요합니다.' };
}

/** 교시표 전체를 교체한다. 지워진 교시의 수업 잠금·예약 기록은 시트에 그대로 남는다. */
function setPeriods_(req) {
  const denied = requireAdmin_(req); if (denied) return denied;
  const list = (req.periods || []).map(function (p) {
    return [String(p.label || '').trim(), String(p.start || '').trim(), String(p.end || '').trim()];
  });
  const labels = list.map(function (r) { return r[0]; });
  if (!list.length) return { ok: false, error: '교시가 하나 이상 필요합니다.' };
  if (list.some(function (r) { return !r[0] || !/^\d{2}:\d{2}$/.test(r[1]) || !/^\d{2}:\d{2}$/.test(r[2]) || r[2] <= r[1]; })) {
    return { ok: false, error: '교시 이름과 시작·종료 시각(HH:MM)을 확인하세요.' };
  }
  if (labels.some(function (l, i) { return labels.indexOf(l) !== i; })) return { ok: false, error: '같은 이름의 교시가 있습니다.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { replaceAll_('periods', list); } finally { lock.releaseLock(); }
  return { ok: true, count: list.length };
}

/** 수업 없는 날(시험·행사 등) 목록 전체를 교체한다. 공간ID가 비어 있으면 모든 공간에 적용. */
function setExceptions_(req) {
  const denied = requireAdmin_(req); if (denied) return denied;
  const roomIds = read_('rooms').map(function (r) { return r['공간ID']; });
  const list = (req.exceptions || []).map(function (x) {
    return [String(x.date || '').trim(), String(x.roomId || '').trim(), String(x.reason || '').trim().slice(0, 40)];
  });
  if (list.some(function (r) { return !/^\d{4}-\d{2}-\d{2}$/.test(r[0]) || (r[1] && roomIds.indexOf(r[1]) < 0); })) {
    return { ok: false, error: '날짜(YYYY-MM-DD)와 공간을 확인하세요.' };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { replaceAll_('exceptions', list); } finally { lock.releaseLock(); }
  return { ok: true, count: list.length };
}

function setLocks_(req) {
  const denied = requireAdmin_(req); if (denied) return denied;
  const roomId = String(req.roomId || ''), termId = String(req.termId || '');
  if (!roomId || !termId) return { ok: false, error: '공간과 학기를 선택하세요.' };
  const cells = (req.cells || []).filter(function (c) { return String(c.label || '').trim() !== ''; });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const kept = read_('locks')
      .filter(function (r) { return !(r['공간ID'] === roomId && r['학기ID'] === termId); })
      .map(function (r) { return [r['공간ID'], r['학기ID'], r['요일'], r['교시'], r['내용']]; });
    const added = cells.map(function (c) {
      return [roomId, termId, String(c.weekday), String(c.period), String(c.label).trim()];
    });
    replaceAll_('locks', kept.concat(added));
  } finally {
    lock.releaseLock();
  }
  return { ok: true, count: cells.length };
}
