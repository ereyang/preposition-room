// 전치사 3D 방 사진 저장소 (index.html의 '📸 Snap 찰칵'이 fetch로 사용)
// POST {a:'snap', grade, cls, num, name, sentences:[...], img:'<base64 JPEG>', thumb:'<작은 base64 JPEG>'}
//   → 내 드라이브 '전치사 방 앨범/2학년 3반/' 폴더에 사진 저장 (설명란에 학생이 쓴 문장)
//   → 같은 폴더의 '전치사 방 제출 목록' 시트에 한 줄 기록 (앱 앨범용 작은 사진 포함)
// POST {a:'setpw', pw}               → 선생님 암호가 없을 때만 처음 정하기
// POST {a:'album', pw, grade, cls}   → 반 목록 + 그 반의 제출 (선생님만)
// POST {a:'photo', pw, id}           → 원본 사진 한 장 (선생님만)
// 처음 한 번: 편집기에서 setup() 실행 → 권한 허용 → 실행 로그에 나온 주소가 반별 앨범 폴더
var ROOT_NAME = '전치사 방 앨범';
var SHEET_NAME = '전치사 방 제출 목록';

function props() { return PropertiesService.getScriptProperties(); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hash(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'preproom:' + s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function teacherOk(pw) { var h = props().getProperty('pw'); return !!h && hash(String(pw || '')) === h; }
function clampInt(v, lo, hi) { var n = parseInt(v, 10); return n >= lo && n <= hi ? n : 0; }

function rootFolder() {
  var id = props().getProperty('root');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder(ROOT_NAME);
  props().setProperty('root', f.getId());
  return f;
}
function subFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function sheet() {
  var id = props().getProperty('sheet');
  if (id) {
    try {
      var old = SpreadsheetApp.openById(id).getSheets()[0];
      if (!old.getRange(1, 9).getValue()) old.getRange(1, 9, 1, 2).setValues([['작은 사진(앱용)', '파일 ID']]);
      return old;
    } catch (e) {}
  }
  var ss = SpreadsheetApp.create(SHEET_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(rootFolder());
  var sh = ss.getSheets()[0];
  sh.appendRow(['보낸 시각', '학년', '반', '번호', '이름', '문장 수', '문장', '사진', '작은 사진(앱용)', '파일 ID']);
  sh.setFrozenRows(1);
  props().setProperty('sheet', ss.getId());
  return sh;
}

function setup() {
  var root = rootFolder();
  sheet();
  Logger.log('반별 앨범 폴더: ' + root.getUrl());
}

function doGet() { return json({ ok: true, app: 'prep-room' }); }

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return json({ ok: false, err: 'bad_request' }); }
  if (!b) return json({ ok: false, err: 'bad_request' });
  if (b.a === 'snap') return snap(b);
  if (b.a === 'setpw') {
    if (props().getProperty('pw')) return json({ ok: false, err: 'exists' });
    if (String(b.pw || '').length < 4) return json({ ok: false, err: 'short' });
    props().setProperty('pw', hash(String(b.pw)));
    return json({ ok: true });
  }
  if (b.a === 'album' || b.a === 'photo') {
    if (!props().getProperty('pw')) return json({ ok: false, err: 'nopw' });
    if (!teacherOk(b.pw)) return json({ ok: false, err: 'pw' });
    return json(b.a === 'album' ? album(b) : photo(b));
  }
  return json({ ok: false, err: 'unknown_action' });
}

function album(b) {
  var rows = sheet().getDataRange().getValues().slice(1);
  var counts = {};
  rows.forEach(function (r) {
    var k = r[1] + '-' + r[2];
    counts[k] = counts[k] || {};
    counts[k][r[3] + '|' + r[4]] = 1;
  });
  var classes = Object.keys(counts).map(function (k) {
    var p = k.split('-');
    return { grade: Number(p[0]), cls: Number(p[1]), n: Object.keys(counts[k]).length };
  }).sort(function (x, y) { return x.grade - y.grade || x.cls - y.cls; });
  var grade = Number(b.grade) || 0, cls = Number(b.cls) || 0;
  if (!classes.some(function (c) { return c.grade === grade && c.cls === cls; })) {
    grade = classes.length ? classes[0].grade : 0; cls = classes.length ? classes[0].cls : 0;
  }
  var out = rows.filter(function (r) { return Number(r[1]) === grade && Number(r[2]) === cls; }).map(function (r) {
    return { t: new Date(r[0]).getTime(), num: Number(r[3]), name: String(r[4]),
      sents: String(r[6] || '').split('\n').filter(String), thumb: String(r[8] || ''), id: String(r[9] || '') };
  });
  return { ok: true, classes: classes, grade: grade, cls: cls, rows: out };
}

function photo(b) {
  try {
    var f = DriveApp.getFileById(String(b.id || ''));
    if (f.getMimeType() !== 'image/jpeg') return { ok: false, err: 'not_found' };
    return { ok: true, img: Utilities.base64Encode(f.getBlob().getBytes()) };
  } catch (e) { return { ok: false, err: 'not_found' }; }
}

function snap(b) {
  var grade = clampInt(b.grade, 1, 3), cls = clampInt(b.cls, 1, 20), num = clampInt(b.num, 1, 60);
  var name = String(b.name || '').replace(/[\\\/:*?"<>|\r\n\t]/g, '').trim().slice(0, 20);
  if (!grade || !cls || !num || !name) return json({ ok: false, err: 'who' });
  var img = String(b.img || '');
  if (!img || img.length > 3000000) return json({ ok: false, err: 'img' });
  var thumb = String(b.thumb || '');
  if (thumb.length > 45000 || !/^[A-Za-z0-9+\/=]*$/.test(thumb)) thumb = '';
  var sents = (Array.isArray(b.sentences) ? b.sentences : []).slice(0, 40).map(function (s) { return String(s).slice(0, 200); });
  var bytes;
  try { bytes = Utilities.base64Decode(img); } catch (x) { return json({ ok: false, err: 'img' }); }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var folder = subFolder(rootFolder(), grade + '학년 ' + cls + '반');
    var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd HHmm');
    var file = folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', num + '번 ' + name + ' ' + stamp + '.jpg'));
    file.setDescription(sents.join('\n'));
    sheet().appendRow([new Date(), grade, cls, num, name, sents.length, sents.join('\n'), file.getUrl(), thumb, file.getId()]);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}
