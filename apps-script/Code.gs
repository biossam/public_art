/**
 * 예술가 퀘스트 · 작품 갤러리 (Google Apps Script)
 * 학생이 올린 작품(예술가 이름, 아이디어, 사진)을 구글 시트와 드라이브에 모으고,
 * 선생님이 '공개'에 체크한 작품만 갤러리에 보여 줍니다.
 *
 * 설치 방법은 apps-script/설치방법.md 를 보세요.
 */

// true: 선생님이 시트에서 '공개'를 체크해야 갤러리에 보임 (권장)
// false: 올리자마자 바로 보임 (선생님은 '공개' 체크를 풀어서 숨길 수 있음)
const APPROVAL_REQUIRED = false;
// 신고가 들어온 작품은 바로 숨김 ('공개' 체크가 풀리고 '신고' 칸에 표시됨)

const SHEET_NAME = '작품';
const FOLDER_NAME = '예술가 퀘스트 작품 사진';
const MISSIONS = ['동시대 공공미술가', '공공미술가', '사회참여 공공미술가', '재사용 공공미술가'];
const MAX_TEXT = 500;
const MAX_IMAGE_BYTES = 1.5 * 1024 * 1024;

// 이름·아이디어에 이런 말이 들어 있으면 저장하지 않음 (필요하면 더 넣으세요)
const BLOCK_WORDS = ['시발', '씨발', 'ㅅㅂ', 'ㅆㅂ', '병신', 'ㅂㅅ', '개새', '좆', '존나', 'ㅈㄴ', '닥쳐', '꺼져', '미친놈', '미친년', '애미', '느금', '섹스', 'fuck', 'shit'];
function hasBadWord_(text) {
  const t = String(text).toLowerCase().replace(/[\s.\-_*~!@#^]/g, '');
  return BLOCK_WORDS.some((w) => t.indexOf(w) >= 0);
}

// 시트에서 연 스크립트면 그 시트를, 따로 만든 스크립트면 '예술가 퀘스트 작품' 시트를 만들어 씀
function book_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* 지워졌으면 새로 만듦 */ } }
  const ss = SpreadsheetApp.create('예술가 퀘스트 작품');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function sheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['올린 시각', '미션', '예술가 이름', '작품 아이디어', '사진 ID', '공개', '사진 미리보기', '신고', '학급 코드', '좋아요']);
    sh.setFrozenRows(1);
  }
  // 예전에 만든 시트에도 새 칸 제목을 채움
  const heads = ['사진 미리보기', '신고', '학급 코드', '좋아요'];
  const cur = sh.getRange(1, 7, 1, 4).getValues()[0];
  if (cur.join('|') !== heads.join('|')) sh.getRange(1, 7, 1, 4).setValues([heads]);
  return sh;
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 갤러리 목록: 공개된 작품만, 최근 것부터
// 학급 코드: 영문·숫자·한글 20자까지, 대문자로 맞춤
function cls_(v) { return String(v || '').trim().toUpperCase().replace(/[^0-9A-Z가-힣_-]/g, '').slice(0, 20); }

function doGet(e) {
  const mission = e && e.parameter && e.parameter.mission;
  const cls = cls_(e && e.parameter && e.parameter.class);
  const rows = sheet_().getDataRange().getValues().slice(1);
  const items = rows
    .map((r, i) => ({ row: i + 2, time: r[0], mission: r[1], name: r[2], idea: r[3], photo: r[4], shown: r[5] === true, cls: cls_(r[8]), likes: Number(r[9]) || 0 }))
    .filter((x) => x.time && x.shown && x.cls === cls && (!mission || x.mission === mission))
    .reverse()
    .slice(0, 200)
    .map((x) => ({
      id: x.row,
      mission: x.mission,
      name: x.name,
      idea: x.idea,
      image: x.photo ? 'https://drive.google.com/thumbnail?id=' + x.photo + '&sz=w800' : '',
      time: Utilities.formatDate(new Date(x.time), 'Asia/Seoul', 'M월 d일'),
      likes: x.likes,
    }));
  return json_({ ok: true, items: items, approval: APPROVAL_REQUIRED });
}

// 신고: 해당 작품을 바로 숨기고 신고 시각을 남김
function report_(row) {
  const sh = sheet_();
  if (!(row >= 2 && row <= sh.getLastRow())) return json_({ ok: false, error: '작품을 찾지 못했어요.' });
  sh.getRange(row, 6).setValue(false);
  sh.getRange(row, 8).setValue('신고됨 ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'M/d HH:mm'));
  return json_({ ok: true });
}

// 좋아요: 숫자 하나 올림 (같은 기기에서 한 번만 누르도록 사이트가 막음)
function like_(row) {
  const sh = sheet_();
  if (!(row >= 2 && row <= sh.getLastRow()) || sh.getRange(row, 6).getValue() !== true) return json_({ ok: false });
  const c = sh.getRange(row, 10);
  const n = (Number(c.getValue()) || 0) + 1;
  c.setValue(n);
  return json_({ ok: true, likes: n });
}

// 작품 올리기
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'report') return report_(Number(d.id));
    if (d.action === 'like') return like_(Number(d.id));
    const mission = String(d.mission || '');
    const name = String(d.name || '').trim().slice(0, 30);
    const idea = String(d.idea || '').trim().slice(0, MAX_TEXT);
    if (MISSIONS.indexOf(mission) < 0 || !name || !idea) return json_({ ok: false, error: '빈 칸이 있어요.' });
    if (hasBadWord_(name) || hasBadWord_(idea)) return json_({ ok: false, error: '고운 말로 다시 써 주세요.' });

    let photoId = '';
    if (d.image) {
      const m = String(d.image).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
      if (!m) return json_({ ok: false, error: '사진 형식이 올바르지 않아요.' });
      const bytes = Utilities.base64Decode(m[2]);
      if (bytes.length > MAX_IMAGE_BYTES) return json_({ ok: false, error: '사진이 너무 커요.' });
      const file = folder_().createFile(Utilities.newBlob(bytes, m[1], mission + '_' + name + '.jpg'));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoId = file.getId();
    }
    // A열(올린 시각)이 빈 첫 줄에 씀 (공개 칸 체크박스만 있는 줄은 빈 줄로 봄)
    const sh = sheet_();
    const times = sh.getRange('A:A').getValues();
    let row = times.length + 1;
    for (let i = 1; i < times.length; i++) { if (times[i][0] === '') { row = i + 1; break; } }
    sh.getRange(row, 1, 1, 6).setValues([[new Date(), mission, name, idea, photoId, !APPROVAL_REQUIRED]]);
    sh.getRange(row, 9, 1, 2).setValues([[cls_(d.class), 0]]);
    sh.getRange(row, 6).insertCheckboxes();
    if (photoId) {
      sh.getRange(row, 7).setFormula('=IMAGE("https://lh3.googleusercontent.com/d/' + photoId + '=w400")');
      sh.setRowHeight(row, 120);
    }
    if (!APPROVAL_REQUIRED) sh.getRange(row, 6).setValue(true);
    return json_({ ok: true, approval: APPROVAL_REQUIRED });
  } catch (err) {
    return json_({ ok: false, error: '저장하지 못했어요. 잠시 뒤 다시 해 주세요.' });
  } finally {
    lock.releaseLock();
  }
}
