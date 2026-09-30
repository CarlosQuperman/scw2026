/**
 * 형성평가 제출 → 교사 구글 시트 기록 (바이브 코딩 2단계 예시)
 *
 * 기록 열: 학번 | 이름 | 1번~6번(O/X) | 점수 | 제출시간(초까지) | 제출 회차
 *
 * 설치: 구글 시트 → 확장 프로그램 → Apps Script
 *       → 이 파일을 Code.gs에, 수정한 HTML을 index(파일명 index.html)로 붙여넣기
 *       → setup 한 번 실행(권한 승인) → 웹 앱으로 배포
 */

// 시트에 붙어 있는(바인딩) 스크립트면 비워 둔다. 독립 스크립트일 때만 시트 ID 입력.
const SHEET_ID = '';

// 정답은 서버에 둔다: 학생이 브라우저에서 점수를 조작해도 여기서 다시 채점한다.
// 인덱스는 0부터 (0=①, 1=②, 2=③, 3=④). HTML의 Q[].c 값과 같아야 한다.
const QUIZZES = {
  crane: { title: '이동식 크레인', answers: [1, 1, 2, 0, 1, 2] },
  // 다음 지문을 추가할 때: quantum: { title: '양자점', answers: [ ... ] },
};

const TZ = 'Asia/Seoul';

/** 웹 앱 첫 화면: index.html을 학생에게 보여 준다 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('이동식 크레인 — 지문으로 배우는 인양의 역학')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1') // 휴대폰 화면 대응(필수)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // 구글 사이트 등에 삽입 허용
}

/** HTML을 앱스크립트 밖(제미나이 공유 등)에서 열었을 때 fetch로 들어오는 제출 */
function doPost(e) {
  let result;
  try {
    result = submitQuiz(JSON.parse(e.postData.contents));
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 제출 처리: 검증 → 서버 채점 → 시트 한 줄 기록 */
function submitQuiz(p) {
  p = p || {};
  const quiz = QUIZZES[p.quizId];
  if (!quiz) throw new Error('등록되지 않은 평가다: ' + p.quizId);

  const id = String(p.studentId || '').trim();
  if (!/^\d{4,6}$/.test(id)) throw new Error('학번 형식이 올바르지 않다.');

  const name = String(p.name || '').trim().slice(0, 20);
  if (!name) throw new Error('이름이 비어 있다.');

  const ans = p.answers;
  if (!Array.isArray(ans) || ans.length !== quiz.answers.length) {
    throw new Error('답안 개수가 맞지 않는다.');
  }

  const marks = quiz.answers.map((c, i) => (Number(ans[i]) === c ? 'O' : 'X'));
  const score = marks.filter(m => m === 'O').length;
  const now = new Date();

  // 40명이 동시에 눌러도 줄이 겹치거나 회차가 꼬이지 않도록 잠금
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(quiz);
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues().flat() : [];
    const attempt = ids.filter(v => v === id).length + 1;

    sh.appendRow([id, safe_(name), ...marks, score, now, attempt]);
    SpreadsheetApp.flush();

    return {
      ok: true,
      score: score,
      total: quiz.answers.length,
      marks: marks,
      attempt: attempt,
      time: Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'),
    };
  } finally {
    lock.releaseLock();
  }
}

/** 평가별 시트를 찾고, 없으면 머리글·서식까지 만든다 */
function getSheet_(quiz) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const name = quiz.title + ' 제출';
  let sh = ss.getSheetByName(name);
  if (sh) return sh;

  sh = ss.insertSheet(name);
  const n = quiz.answers.length;
  const header = ['학번', '이름']
    .concat(quiz.answers.map((_, i) => (i + 1) + '번'))
    .concat(['점수(/' + n + ')', '제출시간', '제출 회차']);

  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#17202B').setFontColor('#F5B700')
    .setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('@');                                   // 학번: 텍스트
  sh.getRange(2, 3, sh.getMaxRows() - 1, n + 1).setHorizontalAlignment('center'); // O/X·점수 가운데
  sh.getRange(2, n + 4, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'); // 초까지

  // O는 초록, X는 빨강
  const oxRange = sh.getRange(2, 3, sh.getMaxRows() - 1, n);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('O')
      .setBackground('#E7F5EE').setFontColor('#1E6B48').setRanges([oxRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('X')
      .setBackground('#FBE7E4').setFontColor('#9E2414').setRanges([oxRange]).build(),
  ]);
  return sh;
}

/** 이름이 =, +, -, @ 로 시작하면 수식으로 해석되지 않게 막는다 */
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

/** 처음 한 번 편집기에서 실행: 권한 승인 + 시간대 설정 + 제출 시트 생성 */
function setup() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(QUIZZES).forEach(k => getSheet_(QUIZZES[k]));
}

/** 배포 전 확인용: 가짜 학생 한 명을 기록해 본다 (확인 후 그 줄은 지우기) */
function testSubmit() {
  Logger.log(submitQuiz({ quizId: 'crane', studentId: '30199', name: '테스트', answers: [1, 1, 2, 0, 0, 0] }));
}
