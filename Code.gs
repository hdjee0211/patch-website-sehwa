/**
 * 동아리 지원서 서버 (Google Apps Script)
 * - 지원서를 구글 시트에 저장
 * - 관리자 로그인 검증 (관리자 정보는 '스크립트 속성'에만 저장 → 웹사이트 소스에 노출되지 않음)
 * - 로그인한 관리자에게만 지원서 목록 제공
 *
 * 스크립트 속성(프로젝트 설정 > 스크립트 속성)에 아래 3개를 등록하세요.
 *   ADMIN_NAME      예) 홍길동
 *   ADMIN_PHONE     예) 01012345678   (하이픈 없이)
 *   ADMIN_PASSWORD  예) 충분히 긴 비밀번호
 * 관리자를 여러 명 두려면 ADMINS 속성에 JSON 배열을 넣어도 됩니다.
 *   [{"name":"홍길동","phone":"01012345678","password":"..."}, {...}]
 */

const SHEET_NAME = '지원서';
const HEADERS = ['ID', '제출일시', '이름', '학번', '전화번호', '희망 전공', '코딩 경험', '자기소개(지원 동기)', '하고 싶은 활동'];
const TOKEN_TTL_SEC = 60 * 60 * 2; // 로그인 유지 2시간
const MAX_FAILS = 10;              // 10분 동안 로그인 실패 허용 횟수

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
      case 'submit': return json(submit_(body.data));
      case 'login':  return json(login_(body));
      case 'list':   return json(list_(body.token));
      case 'logout': return json(logout_(body.token));
      default:       return json({ ok: false, error: '알 수 없는 요청입니다.' });
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
    const sh = getSheet_();
    sh.appendRow([
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

/* ---------------- 관리자 로그인 ---------------- */
function login_(b) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('login_fails') || 0);
  if (fails >= MAX_FAILS) throw new Error('로그인 시도가 너무 많습니다. 10분 후 다시 시도해 주세요.');

  const name = clean_(b.name, 30);
  const phone = digits_(b.phone);
  const password = String(b.password || '');

  const ok = getAdmins_().some(function (a) {
    return a.name === name && digits_(a.phone) === phone && a.password === password;
  });

  if (!ok) {
    cache.put('login_fails', String(fails + 1), 600);
    Utilities.sleep(800);
    throw new Error('관리자 정보가 일치하지 않습니다.');
  }

  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  cache.put('tok_' + token, name, TOKEN_TTL_SEC);
  return { ok: true, token: token };
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

function requireAuth_(token) {
  if (!token || !CacheService.getScriptCache().get('tok_' + token)) {
    const e = new Error('로그인이 필요합니다.');
    e.code = 'AUTH';
    throw e;
  }
}

/* ---------------- 지원서 목록 ---------------- */
function list_(token) {
  requireAuth_(token);
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
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

/* ---------------- 유틸 ---------------- */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#eef0ff');
  }
  return sh;
}
function clean_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
function digits_(v) { return String(v == null ? '' : v).replace(/[^0-9]/g, ''); }
// 시트 수식 주입 방지 (=, +, -, @ 로 시작하는 값)
function text_(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 최초 1회 실행: 시트 생성 + 권한 승인용 */
function setup() {
  getSheet_();
  Logger.log('준비 완료: "' + SHEET_NAME + '" 시트가 생성되었습니다.');
}
