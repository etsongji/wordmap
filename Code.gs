/**
 * 단어 지도(반 운영판) 백엔드 — Google Apps Script
 *
 * [설치]
 *  1. 새 구글 스프레드시트 생성 → 확장 프로그램 → Apps Script
 *  2. 이 코드를 Code.gs에 붙여넣고 아래 TEACHER_PIN을 원하는 값으로 바꾼 뒤 저장
 *  3. 배포 → 새 배포 → 유형: 웹 앱
 *       실행 계정: 나 / 액세스 권한: 모든 사용자
 *  4. 발급된 웹앱 URL을 index.html의 API_URL에 넣기
 *  ※ 코드를 고친 뒤에는 배포 → 배포 관리 → 새 버전으로 다시 배포해야 반영된다.
 *
 * [시트 구조] 시트명 students (자동 생성)
 *   A 번호 | B 이름 | C 데이터(JSON) | D 수정시각
 *   - 학번은 처음 저장한 이름에 묶인다. 다른 이름으로 접속하면 거부된다.
 *     이름을 잘못 넣은 학생은 시트 B열을 직접 고쳐 주면 된다.
 *
 * [API]
 *   GET  ?action=load&num=2103&name=홍길동&callback=cb   → 학생 1명 데이터 (JSONP)
 *   GET  ?action=all&pin=교사비밀번호&callback=cb        → 전체 학생 (JSONP, 교사용)
 *   POST body: {"action":"save","num":"2103","name":"홍길동","data":{...}}
 *
 * [오류 코드] name_mismatch(이름 불일치) / bad_pin(교사 비밀번호 틀림) / pin_not_set(비밀번호 미설정)
 */

const SHEET_NAME = 'students';
const TEACHER_PIN = '';   // ← 교사 비밀번호. 비워 두면 교사 화면 접속이 막힌다. 반드시 설정할 것.

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['번호', '이름', '데이터', '수정시각']);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@'); // 학번 앞자리 0 보존 (예: 0103)
  }
  return sh;
}

// 이름 비교용: 공백 제거 (홍 길동 == 홍길동)
function normName_(s) {
  return String(s || '').replace(/\s+/g, '');
}

function findRow_(sh, num) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(num).trim()) return i + 2;
  }
  return -1;
}

function parseData_(cell) {
  try { return cell ? JSON.parse(cell) : null; } catch (e) { return null; }
}

function loadStudent_(num, name) {
  if (!num) return { ok: false, error: 'num required' };
  const sh = sheet_();
  const r = findRow_(sh, num);
  if (r < 0) return { ok: true, found: false };
  const row = sh.getRange(r, 1, 1, 4).getValues()[0];
  const stored = normName_(row[1]);
  if (name && stored && stored !== normName_(name)) {
    return { ok: false, error: 'name_mismatch' };
  }
  return { ok: true, found: true, num: String(row[0]), name: String(row[1]), data: parseData_(row[2]), updated: String(row[3]) };
}

function allStudents_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 4).getValues();
  const out = [];
  vals.forEach(function (row) {
    if (!row[0]) return;
    out.push({ num: String(row[0]), name: String(row[1]), data: parseData_(row[2]), updated: String(row[3]) });
  });
  return out;
}

function saveStudent_(num, name, data) {
  if (!num) throw new Error('num required');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const json = JSON.stringify(data || {});
    const r = findRow_(sh, num);
    if (r < 0) {
      const newRow = sh.getLastRow() + 1;
      sh.getRange(newRow, 1, 1, 4).setNumberFormat('@').setValues([[String(num), String(name || ''), json, now]]);
    } else {
      const stored = normName_(sh.getRange(r, 2).getValue());
      if (stored && normName_(name) && stored !== normName_(name)) {
        return { ok: false, error: 'name_mismatch' };
      }
      sh.getRange(r, 2, 1, 3).setValues([[String(name || sh.getRange(r, 2).getValue()), json, now]]);
    }
    return { ok: true, updated: now };
  } finally {
    lock.releaseLock();
  }
}

function checkPin_(pin) {
  if (!TEACHER_PIN) return { ok: false, error: 'pin_not_set' };
  if (String(pin || '') !== TEACHER_PIN) return { ok: false, error: 'bad_pin' };
  return null;
}

function jsonOut_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    if (p.action === 'load') out = loadStudent_(p.num, p.name);
    else if (p.action === 'all') out = checkPin_(p.pin) || { ok: true, students: allStudents_() };
    else if (p.action === 'save') out = saveStudent_(p.num, p.name, JSON.parse(p.data || '{}'));
    else out = { ok: true, msg: 'word-map backend alive' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return jsonOut_(out, p.callback);
}

function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'save') out = saveStudent_(body.num, body.name, body.data);
    else out = { ok: false, error: 'unknown action' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return jsonOut_(out);
}
