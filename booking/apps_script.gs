/**
 * 11/28(토) 1:1 진단·학습 컨설팅 선착순 예약 — Google Apps Script 웹 앱
 * 구글 시트에 붙여 쓰는 스크립트다. 설치 방법은 booking/설정방법.md 참고.
 *
 * doGet  : 시간별 예약 수만 돌려준다 (이름·연락처는 절대 내보내지 않음)
 * doPost : 예약 1건 저장. LockService로 같은 시간 동시 예약을 막는다.
 */

const CAP = 1;            // 시간당 정원. 강사 2명이 동시에 상담하면 2로 바꾸고 "새 배포"할 것
const SLOTS = ['10:00', '10:30', '11:00', '11:30', '12:00', '12:30', '13:00', '13:30'];
const SHEET = '예약';
const HEAD = ['접수시각', '시간', '학생 이름', '학교·학년', '학부모 연락처', '메모', '예약번호'];
// 새 예약 알림 메일 받을 주소. 비워 두면 스크립트를 배포한 구글 계정으로 보낸다(공개 저장소에 주소를 적지 않기 위함)
const NOTIFY = '';
// 시트 밖(script.google.com)에서 만든 스크립트면 시트 주소의 /d/ 와 /edit 사이 ID를 넣는다. 시트에서 연 스크립트면 비워 둔다
const SHEET_ID = '';
const ss_ = () => SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();

function sheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange('E:E').setNumberFormat('@');  // 연락처 앞자리 0 유지
  }
  return sh;
}

function counts_(sh) {
  const c = {};
  SLOTS.forEach(t => (c[t] = 0));
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    sh.getRange(2, 2, n, 1).getDisplayValues().forEach(r => {
      if (r[0] in c) c[r[0]]++;
    });
  }
  return c;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 수식으로 해석되지 않게(=, +, -, @로 시작하는 값) 앞에 ' 를 붙이고 길이를 자른다
function clean_(v, max) {
  v = String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doGet() {
  return out_({ ok: true, cap: CAP, counts: counts_(sheet_()) });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, code: 'bad' }); }
  if (d.website) return out_({ ok: true });  // 스팸 봇(숨은 칸을 채움)은 저장하지 않음

  const slot = String(d.slot || '');
  const name = clean_(d.name, 20), school = clean_(d.school, 40), memo = clean_(d.memo, 300);
  const phone = String(d.phone || '').replace(/\D/g, '');
  if (SLOTS.indexOf(slot) < 0 || !name || !school || !/^01\d{8,9}$/.test(phone)) return out_({ ok: false, code: 'bad' });
  const tel = phone.replace(/^(\d{3})(\d{3,4})(\d{4})$/, '$1-$2-$3');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return out_({ ok: false, code: 'busy' });
  try {
    const sh = sheet_();
    const c = counts_(sh);
    if (c[slot] >= CAP) return out_({ ok: false, code: 'full', cap: CAP, counts: c });
    const n = sh.getLastRow() - 1;
    if (n > 0 && sh.getRange(2, 5, n, 1).getDisplayValues().some(r => r[0] === tel)) {
      return out_({ ok: false, code: 'dup', cap: CAP, counts: c });
    }
    const no = '1128-' + slot.replace(':', '') + (CAP > 1 ? '-' + (c[slot] + 1) : '');
    sh.appendRow([new Date(), slot, name, school, tel, memo, no]);
    SpreadsheetApp.flush();
    lock.releaseLock();
    notify_(no, slot, name, school, tel, memo);
    return out_({ ok: true, no: no });
  } finally {
    lock.releaseLock();
  }
}

// 메일이 실패해도 예약 자체는 이미 저장되어 있으므로 오류를 삼킨다
function notify_(no, slot, name, school, tel, memo) {
  try {
    const to = NOTIFY || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail(to, '[11/28 상담 예약] ' + slot + ' ' + name.replace(/^'/, ''),
      '예약번호: ' + no + '\n시간: 11월 28일(토) ' + slot + '\n학생: ' + name.replace(/^'/, '') +
      '\n학교·학년: ' + school.replace(/^'/, '') + '\n연락처: ' + tel + '\n메모: ' + (memo.replace(/^'/, '') || '-') +
      '\n\n시트: ' + ss_().getUrl());
  } catch (err) {}
}

// 처음 한 번 편집기에서 실행하면 메일 권한 승인 창이 뜨고 시험 메일이 온다
function testMail() {
  notify_('1128-0000', '00:00', '시험', '시험', '010-0000-0000', '알림 메일 시험');
}
