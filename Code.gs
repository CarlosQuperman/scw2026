/**
 * 바이브 코딩 3단계 예시 — 수시 지원 상담카드
 *  · 방대한 자료(27,117행)는 HTML에도, 시트에도 넣지 않는다.
 *  · 교사 드라이브 폴더에 JSON으로 두고, 필요한 대학 파일 하나만 불러온다.
 *  · 학생이 '교사에게 제출'을 누르면 교사 시트 두 탭에 누적한다.
 *
 * 드라이브 폴더 구조 (zip을 풀어 폴더째 업로드)
 *   수시자료/
 *     sched.json          ← 대학별 고사일 (시작할 때 한 번)
 *     uni/가천대.json ... ← 대학별 전형 자료 179개 (대학을 고를 때 하나씩)
 */

const DATA_FOLDER_ID = '';           // ← '수시자료' 폴더 주소의 folders/ 뒤 문자열
const TZ = 'Asia/Seoul';
const SHEET_STUDENT = '상담카드_학생';
const SHEET_APPLY = '상담카드_지원';

/** 웹 앱 첫 화면 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('2027 수시 지원 상담카드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 처음 한 번, 그리고 JSON 파일을 새로 올릴 때마다 실행.
 * 폴더를 훑어 '대학명 → 파일 ID' 목록을 스크립트 속성에 저장하고, 제출 시트를 만든다.
 */
function setup() {
  if (!DATA_FOLDER_ID) throw new Error('DATA_FOLDER_ID를 먼저 입력해 줘.');
  const props = PropertiesService.getScriptProperties();
  props.getKeys().filter(k => k.indexOf('U:') === 0).forEach(k => props.deleteProperty(k));

  const map = {};
  let schedId = '';
  (function walk(folder) {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName();
      if (!/\.json$/i.test(name)) continue;
      if (name === 'sched.json') schedId = f.getId();
      else map['U:' + name.replace(/\.json$/i, '')] = f.getId();
    }
    const subs = folder.getFolders();
    while (subs.hasNext()) walk(subs.next());
  })(DriveApp.getFolderById(DATA_FOLDER_ID));

  const unis = Object.keys(map).map(k => k.slice(2));
  props.setProperties(Object.assign(map, { SCHED_ID: schedId, UNIS: JSON.stringify(unis) }));
  CacheService.getScriptCache().removeAll(unis.map(u => 'U:' + u));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  studentSheet_(); applySheet_();
  Logger.log('대학 파일 ' + unis.length + '개, 고사일 파일 ' + (schedId ? '있음' : '없음'));
}

/** 시작할 때: 대학 목록 + 고사일 */
function getBoot() {
  const props = PropertiesService.getScriptProperties();
  const unis = JSON.parse(props.getProperty('UNIS') || '[]');
  if (!unis.length) throw new Error('자료 목록이 비어 있다. 선생님이 setup을 실행해야 한다.');
  const schedId = props.getProperty('SCHED_ID');
  const sched = schedId ? DriveApp.getFileById(schedId).getBlob().getDataAsString('UTF-8') : '[]';
  return { unis: unis, sched: sched };
}

/** 대학을 고를 때: 그 대학 JSON 하나만 (작은 파일은 6시간 캐시) */
function getUni(name) {
  const key = 'U:' + String(name || '').trim();
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return hit;
  const id = PropertiesService.getScriptProperties().getProperty(key);
  if (!id) throw new Error('자료에 없는 대학이다: ' + name);
  const text = DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8');
  if (Utilities.newBlob(text).getBytes().length < 95000) cache.put(key, text, 21600);
  return text;
}

/** 제출: 학생 한 줄 + 지원 건마다 한 줄 */
function submitCard(p) {
  p = p || {};
  const id = String(p.studentId || '').trim();
  if (!/^\d{4,6}$/.test(id)) throw new Error('학번 형식이 올바르지 않다.');
  const name = String(p.name || '').trim().slice(0, 20);
  if (!name) throw new Error('이름이 비어 있다.');
  const rows = Array.isArray(p.rows) ? p.rows.slice(0, 12) : [];
  if (!rows.length) throw new Error('지원 대학이 없다.');
  const h = p.header || {};
  const now = new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = studentSheet_();
    const last = ss.getLastRow();
    const ids = last > 1 ? ss.getRange(2, 1, last - 1, 1).getDisplayValues().flat() : [];
    const attempt = ids.filter(v => v === id).length + 1;

    ss.appendRow([id, safe_(name), safe_(h.naeshin), safe_(h.jinro), safe_(h.hgwa),
      h.m_kor || '', h.m_math || '', h.m_eng || '', h.m_t1 || '', h.m_t2 || '', h.m_hist || '',
      rows.length, now, attempt]);

    const as = applySheet_();
    const out = rows.map(r => [id, safe_(name), attempt, r.no,
      safe_(r.uni), safe_(r.major), safe_(r.typeKind), safe_(r.typeName),
      r.quota || '', r.comp || '', safe_(r.avg), safe_(r.conv), r.myGrade === '' ? '' : r.myGrade,
      safe_(r.verdict), safe_(r.minreq), safe_(r.minOk), safe_(r.strategy), safe_(r.exam), now]);
    as.getRange(as.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
    SpreadsheetApp.flush();

    return { ok: true, count: rows.length, attempt: attempt,
      time: Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss') };
  } finally {
    lock.releaseLock();
  }
}

function studentSheet_() {
  return sheet_(SHEET_STUDENT, ['학번', '이름', '내신성적', '희망진로', '희망학과',
    '모의 국어', '모의 수학', '모의 영어', '모의 탐구1', '모의 탐구2', '모의 한국사',
    '지원 수', '제출시간', '제출 회차'], 13);
}
function applySheet_() {
  return sheet_(SHEET_APPLY, ['학번', '이름', '제출 회차', '순', '대학명', '학과', '전형유형', '전형명',
    '모집인원', '전년도 경쟁률', '전년도 입결', '대학별 환산등급', '비교 등급', '입결 판정(추정)',
    '수능최저', '최저 충족(추정)', '지원전략(학생 선택)', '대학별 고사일', '제출시간'], 19);
}
function sheet_(name, header, timeCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#17456b').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange(2, timeCol, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  return sh;
}
function safe_(v) {
  const s = v === null || v === undefined ? '' : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** 확인용: 편집기에서 실행해 자료가 잘 읽히는지 본다 */
function testRead() {
  const b = getBoot();
  Logger.log('대학 ' + b.unis.length + '개 / 고사일 ' + JSON.parse(b.sched).length + '건');
  const o = JSON.parse(getUni(b.unis[0]));
  Logger.log(b.unis[0] + ': ' + o.rows.length + '행');
}
