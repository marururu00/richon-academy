/**
 * ════════════════════════════════════════════════════════════════════════
 *  리치온 아카데미 · 신청 자동 처리 스크립트 (Google Apps Script)
 *  목적: 구글 폼 신청이 들어오면 → 구글 시트에 자동 적재(기본 기능) →
 *        신청자에게 ZOOM 입장 안내·수강 안내 메일을 "자동" 발송.
 *        → 선생님이 댓글로 일일이 응대할 필요가 없어집니다.
 * ════════════════════════════════════════════════════════════════════════
 *
 *  [설치 방법]
 *  1) 구글 폼을 만들고, 응답을 구글 시트로 연결합니다.
 *     (폼 편집 화면 → '응답' 탭 → 시트 아이콘 → 스프레드시트 연결)
 *     폼 질문 예시: 이름 / 이메일 / 연락처 / 신청 과정 / 입금자명 /
 *                  현금영수증 유형(소득공제용·지출증빙용·미발급) / 현금영수증 번호
 *  2) 연결된 "스프레드시트" 열기 → 확장 프로그램 → Apps Script.
 *  3) 이 코드를 붙여넣고 저장.
 *  4) 상단 함수 목록에서 setupTrigger 를 한 번 실행(권한 허용).
 *     → 이후 신청이 들어올 때마다 onFormSubmit 이 자동 실행됩니다.
 *
 *  [주의] 아래 대문자 설정값과 질문 제목(QCOL)을 실제 폼에 맞게 바꾸세요.
 */

// ── 설정 ────────────────────────────────────────────────────────────────
var FROM_NAME = '리치온 아카데미';
var ACCOUNT   = '국민 123456-78-901234 (예금주: 홍길동)'; // 무통장입금 계좌
// 과정명 → 안내에 들어갈 ZOOM/수강 링크 (실제 값으로 교체)
var ZOOM_BY_COURSE = {
  '무료 브리핑'   : 'https://zoom.us/j/000000',
  'Pre리치온'     : 'https://zoom.us/j/111111',
  '리치온 스터디' : 'https://zoom.us/j/222222',
  '재개발 중급반' : 'https://zoom.us/j/333333',
  '리치온 인테리어': 'https://zoom.us/j/444444',
  '웰컴 클래스'   : 'https://zoom.us/j/555555'
};
var DEFAULT_ZOOM = 'https://zoom.us/j/000000';
// 폼 질문 "제목" 그대로 입력 (폼에서 쓴 질문 텍스트와 일치해야 함)
var QCOL = {
  email : '이메일',
  name  : '이름',
  course: '신청 과정',
  cash  : '현금영수증 유형'
};
// ────────────────────────────────────────────────────────────────────────

function setupTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 중복 방지: 기존 onFormSubmit 트리거 제거 후 재등록
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onFormSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  Logger.log('트리거 등록 완료');
}

function onFormSubmit(e) {
  var v = e.namedValues || {};
  var get = function (k) { return (v[k] && v[k][0]) ? v[k][0].trim() : ''; };

  var email  = get(QCOL.email);
  var name   = get(QCOL.name) || '신청자';
  var course = get(QCOL.course) || '신청하신 과정';
  var cash   = get(QCOL.cash);
  if (!email) return; // 이메일 없으면 발송 안 함

  var zoom = ZOOM_BY_COURSE[course] || DEFAULT_ZOOM;

  var subject = '[리치온 아카데미] ' + course + ' 신청이 접수되었습니다';
  var body =
    name + '님, 안녕하세요. 리치온 아카데미입니다.\n\n' +
    '▶ 신청 과정: ' + course + '\n' +
    '▶ ZOOM 입장 주소: ' + zoom + '\n\n' +
    '아직 결제 전이시라면 아래 계좌로 입금해 주세요.\n' +
    '▶ 입금 계좌: ' + ACCOUNT + '\n' +
    (cash ? ('▶ 현금영수증: ' + cash + ' (신청서에 입력하신 정보로 발급해 드립니다)\n') : '') +
    '\n입금이 확인되면 수강이 확정됩니다. 강의에서 뵙겠습니다!\n— ' + FROM_NAME;

  MailApp.sendEmail({ to: email, name: FROM_NAME, subject: subject, body: body });

  // 시트 마지막 열에 발송 표시(선택)
  try {
    var sh = e.range.getSheet();
    sh.getRange(e.range.getRow(), sh.getLastColumn() + 1).setValue('안내메일 발송 ' + new Date());
  } catch (err) {}
}
