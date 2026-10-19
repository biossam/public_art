/**
 * 예술가 퀘스트 · 작품 갤러리 (Google Apps Script)
 * 학생이 올린 작품(예술가 이름, 아이디어, 사진)을 구글 시트와 드라이브에 모으고,
 * 선생님이 '공개'에 체크한 작품만 갤러리에 보여 줍니다.
 *
 * 설치 방법은 apps-script/설치방법.md 를 보세요.
 */

// true: 선생님이 시트에서 '공개'를 체크해야 갤러리에 보임 (권장)
// false: 올리자마자 바로 보임 (선생님은 '공개' 체크를 풀어서 숨길 수 있음)
const APPROVAL_REQUIRED = true;

const SHEET_NAME = '작품';
const FOLDER_NAME = '예술가 퀘스트 작품 사진';
const MISSIONS = ['동시대 공공미술가', '공공미술가', '사회참여 공공미술가', '재사용 공공미술가'];
const MAX_TEXT = 500;
const MAX_IMAGE_BYTES = 1.5 * 1024 * 1024;

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
    sh.appendRow(['올린 시각', '미션', '예술가 이름', '작품 아이디어', '사진 ID', '공개']);
    sh.setFrozenRows(1);
  }
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
function doGet(e) {
  const mission = e && e.parameter && e.parameter.mission;
  const rows = sheet_().getDataRange().getValues().slice(1);
  const items = rows
    .map((r, i) => ({ row: i + 2, time: r[0], mission: r[1], name: r[2], idea: r[3], photo: r[4], shown: r[5] === true }))
    .filter((x) => x.time && x.shown && (!mission || x.mission === mission))
    .reverse()
    .slice(0, 200)
    .map((x) => ({
      id: x.row,
      mission: x.mission,
      name: x.name,
      idea: x.idea,
      image: x.photo ? 'https://drive.google.com/thumbnail?id=' + x.photo + '&sz=w800' : '',
      time: Utilities.formatDate(new Date(x.time), 'Asia/Seoul', 'M월 d일'),
    }));
  return json_({ ok: true, items: items, approval: APPROVAL_REQUIRED });
}

// 작품 올리기
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const mission = String(d.mission || '');
    const name = String(d.name || '').trim().slice(0, 30);
    const idea = String(d.idea || '').trim().slice(0, MAX_TEXT);
    if (MISSIONS.indexOf(mission) < 0 || !name || !idea) return json_({ ok: false, error: '빈 칸이 있어요.' });

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
    sh.getRange(row, 6).insertCheckboxes();
    if (!APPROVAL_REQUIRED) sh.getRange(row, 6).setValue(true);
    return json_({ ok: true, approval: APPROVAL_REQUIRED });
  } catch (err) {
    return json_({ ok: false, error: '저장하지 못했어요. 잠시 뒤 다시 해 주세요.' });
  } finally {
    lock.releaseLock();
  }
}
