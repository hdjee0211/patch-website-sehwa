/**
 * PATCH 동아리 서버 (Google Apps Script)
 * - 지원서 저장 / 관리자 로그인 / 지원서 목록
 * - 부원 로그인 / 부원 게시판(일정·과제) 보기·등록·수정·삭제
 * - 부원 익명 건의함 (운영진만 열람)
 *
 * 프로젝트 설정 > 스크립트 속성에 아래 5개를 등록하세요.
 *   ADMIN_NAME       관리자 이름
 *   ADMIN_PHONE      관리자 전화번호 (하이픈 없이)
 *   ADMIN_PASSWORD   관리자 비밀번호
 *   MEMBER_ID        부원 아이디   (예: patchsw)
 *   MEMBER_PASSWORD  부원 비밀번호 (예: patchsw)
 * 관리자를 여러 명 두려면 ADMINS 속성에 JSON 배열을 넣어도 됩니다.
 *   [{"name":"홍길동","phone":"01012345678","password":"..."}, {...}]
 */

const SHEET_NAME = '지원서';
const HEADERS = ['ID', '제출일시', '이름', '학번', '전화번호', '희망 전공', '코딩 경험', '자기소개(지원 동기)', '하고 싶은 활동'];
const BOARD_SHEET = '부원게시판';
const BOARD_HEADERS = ['ID', '종류', '날짜', '시간', '제목', '장소', '내용', '중요', '링크', '작성일시'];
const KIND_LABEL = { schedule: '일정', task: '과제', notice: '공지' };
const SUGGEST_SHEET = '건의사항';
const SUGGEST_HEADERS = ['ID', '제출일시', '분류', '내용', '확인'];
const SUGGEST_CATEGORIES = ['활동', '운영', '시설·장소', '기타'];
const ADMIN_TTL_SEC = 60 * 60 * 2;   // 관리자 로그인 유지 2시간
const MEMBER_TTL_SEC = 60 * 60 * 6;  // 부원 로그인 유지 6시간 (최대값)
const MAX_FAILS = 10;                // 10분 동안 로그인 실패 허용 횟수

function doGet() {
  return json({ ok: true, message: '동아리 지원서 API가 정상 동작 중입니다.' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: '잘못된 요청입니다.' });
  }
  try {
    switch (body.action) {
      case 'submit':      return json(submit_(body.data));
      case 'login':       return json(login_(body));
      case 'memberLogin': return json(memberLogin_(body));
      case 'list':        return json(list_(body.token));
      case 'logout':      return json(logout_(body.token));
      case 'boardList':   return json(boardList_(body.token));
      case 'boardSave':   return json(boardSave_(body.token, body.item));
      case 'boardDelete': return json(boardDelete_(body.token, body.id));
      case 'suggestSubmit': return json(suggestSubmit_(body.token, body.item));
      case 'suggestList':   return json(suggestList_(body.token));
      case 'suggestUpdate': return json(suggestUpdate_(body.token, body.id, body.done));
      case 'suggestDelete': return json(suggestDelete_(body.token, body.id));
      default:            return json({ ok: false, error: '알 수 없는 요청입니다. Apps Script를 새 버전으로 배포했는지 확인해 주세요.' });
    }
  } catch (err) {
    return json({ ok: false, error: String(err.message || err), code: err.code });
  }
}

/* ---------------- 지원서 제출 ---------------- */
function submit_(d) {
  if (!d) throw new Error('지원서 데이터가 없습니다.');
  const rec = {
    name: clean_(d.name, 30),
    studentId: clean_(d.studentId, 20),
    phone: digits_(d.phone).slice(0, 11),
    major: clean_(d.major, 50),
    codingExp: d.codingExp === '있음' ? '있음' : '없음',
    motivation: clean_(d.motivation, 1000),
    activities: clean_(d.activities, 1000),
  };
  if (!rec.name || !rec.studentId || !rec.phone || !rec.major || !rec.motivation || !rec.activities) {
    throw new Error('필수 항목이 비어 있습니다.');
  }
  if (!/^01[016789]\d{7,8}$/.test(rec.phone)) throw new Error('전화번호 형식이 올바르지 않습니다.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getSheet_(SHEET_NAME, HEADERS).appendRow([
      Utilities.getUuid().slice(0, 8),
      new Date(),
      text_(rec.name),
      "'" + rec.studentId,   // 앞자리 0 보존
      "'" + rec.phone,
      text_(rec.major),
      rec.codingExp,
      text_(rec.motivation),
      text_(rec.activities),
    ]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/* ---------------- 로그인 ---------------- */
function checkFails_() {
  const fails = Number(CacheService.getScriptCache().get('login_fails') || 0);
  if (fails >= MAX_FAILS) throw new Error('로그인 시도가 너무 많습니다. 10분 후 다시 시도해 주세요.');
  return fails;
}
function addFail_(fails) {
  CacheService.getScriptCache().put('login_fails', String(fails + 1), 600);
  Utilities.sleep(800);
}
function issueToken_(role, name, ttl) {
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('tok_' + token, JSON.stringify({ role: role, name: name }), ttl);
  return token;
}

function login_(b) {
  const fails = checkFails_();
  const name = clean_(b.name, 30);
  const phone = digits_(b.phone);
  const password = String(b.password || '');
  const ok = getAdmins_().some(function (a) {
    return a.name === name && digits_(a.phone) === phone && a.password === password;
  });
  if (!ok) {
    addFail_(fails);
    throw new Error('관리자 정보가 일치하지 않습니다.');
  }
  return { ok: true, role: 'admin', name: name, token: issueToken_('admin', name, ADMIN_TTL_SEC) };
}

function memberLogin_(b) {
  const fails = checkFails_();
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('MEMBER_ID');
  const pw = p.getProperty('MEMBER_PASSWORD');
  if (!id || !pw) throw new Error('서버에 부원 로그인 정보가 설정되지 않았습니다.');
  if (clean_(b.id, 50) !== id || String(b.password || '') !== pw) {
    addFail_(fails);
    throw new Error('아이디 또는 비밀번호가 일치하지 않습니다.');
  }
  return { ok: true, role: 'member', name: '부원', token: issueToken_('member', '부원', MEMBER_TTL_SEC) };
}

function logout_(token) {
  if (token) CacheService.getScriptCache().remove('tok_' + token);
  return { ok: true };
}

function getAdmins_() {
  const p = PropertiesService.getScriptProperties();
  const multi = p.getProperty('ADMINS');
  if (multi) {
    try { return JSON.parse(multi); } catch (e) { throw new Error('ADMINS 속성의 JSON 형식이 잘못되었습니다.'); }
  }
  const name = p.getProperty('ADMIN_NAME');
  const phone = p.getProperty('ADMIN_PHONE');
  const password = p.getProperty('ADMIN_PASSWORD');
  if (!name || !phone || !password) throw new Error('서버에 관리자 정보가 설정되지 않았습니다.');
  return [{ name: name, phone: phone, password: password }];
}

function requireRole_(token, roles) {
  const raw = token ? CacheService.getScriptCache().get('tok_' + token) : null;
  let s = null;
  try { s = raw ? JSON.parse(raw) : null; } catch (e) { s = null; }
  if (!s || roles.indexOf(s.role) === -1) {
    const e = new Error(s ? '권한이 없습니다.' : '로그인이 필요합니다.');
    e.code = 'AUTH';
    throw e;
  }
  return s;
}

/* ---------------- 지원서 목록 (관리자) ---------------- */
function list_(token) {
  requireRole_(token, ['admin']);
  const values = getSheet_(SHEET_NAME, HEADERS).getDataRange().getValues();
  const items = values.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    return {
      id: String(r[0]),
      submittedAt: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
      name: String(r[2]),
      studentId: String(r[3]),
      phone: String(r[4]),
      major: String(r[5]),
      codingExp: String(r[6]),
      motivation: String(r[7]),
      activities: String(r[8]),
    };
  });
  items.reverse(); // 최신순
  return { ok: true, items: items };
}

/* ---------------- 부원 게시판 (일정 · 과제 · 공지) ---------------- */
function kindFromLabel_(v) {
  v = String(v || '').trim();
  for (const k in KIND_LABEL) { if (KIND_LABEL[k] === v || k === v) return k; }
  return '';
}

function boardList_(token) {
  requireRole_(token, ['admin', 'member']);
  const values = getSheet_(BOARD_SHEET, BOARD_HEADERS).getDataRange().getValues();
  const items = [];
  values.slice(1).forEach(function (r) {
    const kind = kindFromLabel_(r[1]);
    if (!r[0] || !kind) return;
    items.push({
      id: String(r[0]),
      kind: kind,
      date: fmtDate_(r[2]),
      time: fmtTime_(r[3]),
      title: String(r[4]),
      place: String(r[5]),
      content: String(r[6]),
      important: r[7] === true || String(r[7]).toUpperCase() === 'TRUE' || r[7] === '중요',
      link: String(r[8] || ''),
    });
  });
  return { ok: true, items: items };
}

function boardSave_(token, item) {
  requireRole_(token, ['admin']);
  if (!item) throw new Error('데이터가 없습니다.');
  const kind = KIND_LABEL[item.kind] ? item.kind : '';
  if (!kind) throw new Error('종류가 올바르지 않습니다.');
  const rec = {
    date: clean_(item.date, 10),
    time: clean_(item.time, 5),
    title: clean_(item.title, 100),
    place: clean_(item.place, 100),
    content: clean_(item.content, 3000),
    link: clean_(item.link, 500),
    important: item.important === true,
  };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(rec.date)) throw new Error('날짜를 입력해 주세요.');
  if (rec.time && !/^\d{2}:\d{2}$/.test(rec.time)) throw new Error('시간 형식이 올바르지 않습니다.');
  if (!rec.title) throw new Error('제목을 입력해 주세요.');
  if (rec.link && !/^https?:\/\//i.test(rec.link)) throw new Error('링크는 http:// 또는 https:// 로 시작해야 합니다.');

  const row = [
    null,
    KIND_LABEL[kind],
    "'" + rec.date,
    rec.time ? "'" + rec.time : '',
    text_(rec.title),
    text_(rec.place),
    text_(rec.content),
    rec.important ? '중요' : '',
    rec.link,
    new Date(),
  ];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(BOARD_SHEET, BOARD_HEADERS);
    const id = clean_(item.id, 20);
    if (id) {
      const r = findRow_(sh, id);
      if (!r) throw new Error('수정할 항목을 찾지 못했습니다. 새로고침 후 다시 시도해 주세요.');
      row[0] = id;
      sh.getRange(r, 1, 1, row.length).setValues([row]);
      return { ok: true, id: id };
    }
    row[0] = Utilities.getUuid().slice(0, 8);
    sh.appendRow(row);
    return { ok: true, id: row[0] };
  } finally {
    lock.releaseLock();
  }
}

function boardDelete_(token, id) {
  requireRole_(token, ['admin']);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(BOARD_SHEET, BOARD_HEADERS);
    const r = findRow_(sh, clean_(id, 20));
    if (r) sh.deleteRow(r);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- 익명 건의함 ---------------- */
// 이름·아이디·IP 등 보낸 사람을 알 수 있는 정보는 저장하지 않습니다.
function suggestSubmit_(token, item) {
  requireRole_(token, ['member', 'admin']);
  const cache = CacheService.getScriptCache();
  const key = 'sg_' + token;
  const n = Number(cache.get(key) || 0);
  if (n >= 5) throw new Error('건의는 10분에 5개까지 보낼 수 있어요. 잠시 후 다시 시도해 주세요.');
  const category = item && SUGGEST_CATEGORIES.indexOf(item.category) >= 0 ? item.category : '기타';
  const content = clean_(item && item.content, 1000);
  if (!content) throw new Error('건의 내용을 입력해 주세요.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getSheet_(SUGGEST_SHEET, SUGGEST_HEADERS).appendRow([Utilities.getUuid().slice(0, 8), new Date(), category, text_(content), '']);
  } finally {
    lock.releaseLock();
  }
  cache.put(key, String(n + 1), 600);
  return { ok: true };
}

function suggestList_(token) {
  requireRole_(token, ['admin']);
  const values = getSheet_(SUGGEST_SHEET, SUGGEST_HEADERS).getDataRange().getValues();
  const items = values.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    return {
      id: String(r[0]),
      createdAt: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
      category: String(r[2]),
      content: String(r[3]),
      done: String(r[4]) === '확인',
    };
  });
  items.reverse(); // 최신순
  return { ok: true, items: items };
}

function suggestUpdate_(token, id, done) {
  requireRole_(token, ['admin']);
  const sh = getSheet_(SUGGEST_SHEET, SUGGEST_HEADERS);
  const r = findRow_(sh, clean_(id, 20));
  if (!r) throw new Error('항목을 찾지 못했습니다. 새로고침해 주세요.');
  sh.getRange(r, 5).setValue(done === true ? '확인' : '');
  return { ok: true };
}

function suggestDelete_(token, id) {
  requireRole_(token, ['admin']);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(SUGGEST_SHEET, SUGGEST_HEADERS);
    const r = findRow_(sh, clean_(id, 20));
    if (r) sh.deleteRow(r);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function findRow_(sh, id) {
  if (!id) return 0;
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return i + 1;
  }
  return 0;
}

/* ---------------- 유틸 ---------------- */
function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8ebf7');
  }
  return sh;
}
function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim();
}
function fmtTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  return String(v || '').trim();
}
function clean_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
function digits_(v) { return String(v == null ? '' : v).replace(/[^0-9]/g, ''); }
// 시트 수식 주입 방지 (=, +, -, @ 로 시작하는 값)
function text_(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 최초 1회 실행: 시트 생성 + 권한 승인용 (이미 있으면 그대로 둡니다) */
function setup() {
  getSheet_(SHEET_NAME, HEADERS);
  getSheet_(BOARD_SHEET, BOARD_HEADERS);
  getSheet_(SUGGEST_SHEET, SUGGEST_HEADERS);
  Logger.log('준비 완료: "' + SHEET_NAME + '", "' + BOARD_SHEET + '", "' + SUGGEST_SHEET + '" 시트가 준비되었습니다.');
}

/** 1회 실행: 2026년 동아리 일정 13개 + 과제 1개를 부원게시판에 넣습니다. (이미 있는 항목은 건너뜀) */
function seedPatchSchedule() {
  const sh = getSheet_(BOARD_SHEET, BOARD_HEADERS);
  const have = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    have[kindFromLabel_(r[1]) + '|' + fmtDate_(r[2]) + '|' + String(r[4])] = true;
  });
  const items = [
    '2026-04-02', '2026-04-16', '2026-05-17', '2026-05-21', '2026-06-11', '2026-06-25',
    '2026-08-27', '2026-08-28', '2026-09-10', '2026-10-22', '2026-10-29', '2026-11-12', '2026-12-03',
  ].map(function (d) { return { kind: 'schedule', date: d, title: '동아리 정기 활동', place: 'IT솔루션룸' }; });
  items.push({ kind: 'task', date: '2026-11-12', title: '내가 만들고 싶은 프로그램 만들기(개인활동)', place: '' });
  const rows = [];
  items.forEach(function (it) {
    if (have[it.kind + '|' + it.date + '|' + it.title]) return;
    rows.push([Utilities.getUuid().slice(0, 8), KIND_LABEL[it.kind], "'" + it.date, '', it.title, it.place, '', '', '', new Date()]);
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log(rows.length + '개 항목을 추가했습니다.');
}
