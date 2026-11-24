/**
 * Learner Drop-off backend
 * Paste into Extensions > Apps Script inside your Google Sheet.
 * Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 */
const FOLDER_NAME = 'Learner Drop-off';   // Drive folder created automatically
const SHEET_NAME  = 'Submissions';
const MAX_BYTES   = 10 * 1024 * 1024;     // 10 MB
const ALLOWED     = ['pdf', 'doc', 'docx', 'jpg', 'jpeg', 'png', 'gif', 'webp'];

function doGet() {
  return json_({ ok: true, message: 'Drop-off service is running.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);

    const name = clean_(d.name), section = clean_(d.section), activity = clean_(d.activity);
    if (!name || !activity || !d.data || !d.filename) throw new Error('Missing required fields.');

    const ext = String(d.filename).split('.').pop().toLowerCase();
    if (ALLOWED.indexOf(ext) < 0) throw new Error('File type not allowed.');

    const bytes = Utilities.base64Decode(d.data);
    if (bytes.length > MAX_BYTES) throw new Error('File is larger than 10 MB.');

    // Folder: root > Section > Activity
    let folder = getFolder_(DriveApp.getRootFolder(), FOLDER_NAME);
    if (section) folder = getFolder_(folder, section);
    folder = getFolder_(folder, activity);

    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    const savedName = stamp + '_' + name.replace(/[^\w\- ]/g, '') + '_' + d.filename;
    const file = folder.createFile(Utilities.newBlob(bytes, d.mimeType || 'application/octet-stream', savedName));

    sheet_().appendRow([new Date(), name, section, activity, d.filename, ext.toUpperCase(),
      Math.round(bytes.length / 1024) + ' KB', file.getUrl(), clean_(d.note)]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Timestamp', 'Learner', 'Section', 'Activity', 'File name', 'Type', 'Size', 'Drive link', 'Note']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold');
  }
  return sh;
}

function getFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function clean_(s) { return String(s || '').replace(/[\\\/:*?"<>|]/g, '-').trim().slice(0, 100); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
