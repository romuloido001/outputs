// Learner Output Drop-off — Google Sheet backend
// Single-file version: paste this ALONE into Extensions > Apps Script.

var SHEET_NAME = 'Submissions';
var FOLDER_NAME = 'Learner Drop-off Files';
var ALLOWED = ['pdf','doc','docx','jpg','jpeg','png','gif','webp'];
var MAX_BYTES = 20 * 1024 * 1024;

function doGet() {
  return HtmlService.createHtmlOutput(PAGE_HTML)
    .setTitle('Output Drop-off')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Date & time received', 'Date of upload', 'Learner name', 'Grade', 'Section', 'File name', 'Type', 'Size (KB)', 'File link']);
    sh.setFrozenRows(1);
    sh.getRange('A1:I1').setFontWeight('bold');
  }
  return sh;
}

// Called from the page via google.script.run
function uploadFile(learnerName, grade, section, uploadDate, fileName, mimeType, base64) {
  learnerName = String(learnerName || '').trim().slice(0, 100);
  grade = String(grade || '').trim().slice(0, 20);
  section = String(section || '').trim().slice(0, 50);
  uploadDate = String(uploadDate || '').trim().slice(0, 10);
  if (!grade || !section) throw new Error('Grade and section are required.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(uploadDate)) throw new Error('Date of upload is required.');
  fileName = String(fileName || '').trim();
  if (!learnerName) throw new Error('Name is required.');
  var ext = (fileName.split('.').pop() || '').toLowerCase();
  if (ALLOWED.indexOf(ext) === -1) throw new Error('File type not allowed.');

  var bytes = Utilities.base64Decode(base64);
  if (bytes.length > MAX_BYTES) throw new Error('File is over 20 MB.');

  var safeName = learnerName.replace(/[\\\/:*?"<>|]/g, '_');
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream',
                               safeName + '_G' + grade.replace(/\D/g,'') + '-' + section.replace(/[\\\/:*?"<>|]/g,'_') + '_' + stamp + '_' + fileName);
  var file = getFolder_().createFile(blob);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    getSheet_().appendRow([
      new Date(), uploadDate, learnerName, grade, section, fileName, ext.toUpperCase(),
      Math.max(1, Math.round(bytes.length / 1024)), file.getUrl()
    ]);
  } finally { lock.releaseLock(); }
  return { ok: true, name: fileName };
}


// ---- ZIP download (teacher use, from the Google Sheet menu) ----

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Drop-off')
    .addItem('Download ALL files as ZIP', 'zipAll')
    .addItem('Download ZIP by Grade & Section…', 'zipBySection')
    .addToUi();
}

function zipAll() { makeZip_(null, null); }

function zipBySection() {
  var ui = SpreadsheetApp.getUi();
  var g = ui.prompt('Grade (e.g. 7)', ui.ButtonSet.OK_CANCEL);
  if (g.getSelectedButton() !== ui.Button.OK) return;
  var s = ui.prompt('Section (e.g. Rizal)', ui.ButtonSet.OK_CANCEL);
  if (s.getSelectedButton() !== ui.Button.OK) return;
  makeZip_(g.getResponseText().trim(), s.getResponseText().trim());
}

function makeZip_(grade, section) {
  var ui = SpreadsheetApp.getUi();
  var rows = getSheet_().getDataRange().getValues().slice(1);
  var blobs = [], total = 0;
  rows.forEach(function (r) {
    // columns: 0 received, 1 upload date, 2 name, 3 grade, 4 section, 5 file, 6 type, 7 KB, 8 link
    if (grade && String(r[3]).toLowerCase() !== grade.toLowerCase()) return;
    if (section && String(r[4]).toLowerCase() !== section.toLowerCase()) return;
    var m = String(r[8]).match(/[-\w]{25,}/);
    if (!m) return;
    try {
      var b = DriveApp.getFileById(m[0]).getBlob();
      total += b.getBytes().length;
      blobs.push(b);
    } catch (e) {}
  });
  if (!blobs.length) { ui.alert('No files found.'); return; }
  if (total > 45 * 1024 * 1024) {
    ui.alert('These files total ' + Math.round(total / 1048576) + ' MB, too large to zip here. ' +
      'Use the "by Grade & Section" option, or open the Drive folder "' + FOLDER_NAME +
      '", select all, right-click ▸ Download.');
    return;
  }
  var label = (grade ? 'G' + grade + '-' + section : 'ALL') + '_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');
  var zip = getFolder_().createFile(Utilities.zip(blobs, 'Submissions_' + label + '.zip'));
  var html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif">' + blobs.length + ' file(s) zipped.</p>' +
    '<p><a target="_blank" href="' + zip.getDownloadUrl() + '" ' +
    'style="font:16px sans-serif;padding:10px 16px;background:#2f5bea;color:#fff;border-radius:8px;text-decoration:none">⬇ Download ZIP</a></p>' +
    '<p style="font:13px sans-serif;color:#666">A copy is also saved in your Drive folder.</p>'
  ).setWidth(340).setHeight(170);
  ui.showModalDialog(html, 'ZIP ready');
}

// ---- Page (built in, so no separate Index file is needed) ----
var PAGE_HTML = "<!DOCTYPE html>\n<html>\n<head>\n<base target=\"_top\">\n<meta charset=\"utf-8\">\n<style>\n  :root{--bg:#f6f7fb;--card:#fff;--ink:#1d2433;--mute:#64708a;--line:#dfe3ee;--brand:#2f5bea;--ok:#12805c;--err:#c0352b}\n  @media (prefers-color-scheme:dark){:root{--bg:#12151d;--card:#1b2030;--ink:#eceff7;--mute:#9aa5bf;--line:#2c3348;--brand:#6b8cff;--ok:#4cc79a;--err:#ff7b70}}\n  *{box-sizing:border-box}\n  body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.5 system-ui,Segoe UI,Roboto,sans-serif}\n  main{max-width:640px;margin:0 auto;padding:24px 16px 48px}\n  h1{font-size:1.6rem;margin:0 0 4px}\n  .sub{color:var(--mute);margin:0 0 20px}\n  .card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px}\n  label{display:block;font-weight:600;margin-bottom:6px}\n  input[type=text]{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:10px;background:var(--bg);color:var(--ink);font:inherit}\n  .drop{margin-top:14px;border:2px dashed var(--line);border-radius:14px;padding:28px 12px;text-align:center;cursor:pointer;color:var(--mute)}\n  .drop.over{border-color:var(--brand);color:var(--ink)}\n  .drop strong{color:var(--brand)}\n  select,input[type=date]{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:10px;background:var(--bg);color:var(--ink);font:inherit}\n  .grid{display:grid;grid-template-columns:110px 1fr;gap:12px;margin-top:12px}\n  .dtrow{display:flex;gap:8px}\n  button{font:inherit;border:0;border-radius:10px;padding:9px 14px;cursor:pointer;background:var(--brand);color:#fff}\n  ul{list-style:none;padding:0;margin:14px 0 0}\n  li{padding:8px 0;border-top:1px solid var(--line);font-size:.95rem;display:flex;gap:8px;justify-content:space-between}\n  li span:first-child{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}\n  .ok{color:var(--ok)}.err{color:var(--err)}.wait{color:var(--mute)}\n</style>\n</head>\n<body>\n<main>\n  <h1>\ud83d\udce5 Output Drop-off</h1>\n  <p class=\"sub\">Upload your computer output: PDF, Word, or picture files (max 20 MB each).</p>\n  <div class=\"card\">\n    <label for=\"nm\">Your full name</label>\n    <input id=\"nm\" type=\"text\" placeholder=\"e.g. Maria Santos\" autocomplete=\"name\">\n    <div class=\"grid\">\n      <div><label for=\"gr\">Grade</label>\n        <select id=\"gr\"><option>7</option><option>8</option><option>9</option><option>10</option><option>11</option><option>12</option></select></div>\n      <div><label for=\"sec\">Section</label>\n        <input id=\"sec\" type=\"text\" placeholder=\"e.g. Rizal\" autocomplete=\"off\"></div>\n    </div>\n    <label for=\"dt\" style=\"margin-top:12px\">Date of upload</label>\n    <div class=\"dtrow\"><input id=\"dt\" type=\"date\"><button type=\"button\" id=\"today\">Today</button></div>\n    <div id=\"drop\" class=\"drop\" tabindex=\"0\">\n      <div><strong>Click to choose files</strong> or drag them here</div>\n      <div style=\"font-size:.85rem;margin-top:4px\">PDF \u00b7 Word \u00b7 JPG \u00b7 PNG \u00b7 GIF \u00b7 WEBP</div>\n    </div>\n    <input id=\"pick\" type=\"file\" multiple hidden\n      accept=\".pdf,.doc,.docx,.jpg,.jpeg,.png,.gif,.webp\">\n    <ul id=\"list\"></ul>\n  </div>\n</main>\n<script>\n  var ALLOWED=['pdf','doc','docx','jpg','jpeg','png','gif','webp'];\n  var MAX=20*1024*1024;\n  var drop=document.getElementById('drop'),pick=document.getElementById('pick'),\n      nm=document.getElementById('nm'),list=document.getElementById('list');\n\n  var gr=document.getElementById('gr'),sec=document.getElementById('sec'),dt=document.getElementById('dt');\n  function today(){var d=new Date();return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2)}\n  dt.value=today();\n  document.getElementById('today').onclick=function(){dt.value=today()};\n  drop.onclick=function(){pick.click()};\n  drop.onkeydown=function(e){if(e.key==='Enter'||e.key===' '){e.preventDefault();pick.click()}};\n  pick.onchange=function(){handle(Array.from(pick.files));pick.value=''};\n  drop.ondragover=function(e){e.preventDefault();drop.classList.add('over')};\n  drop.ondragleave=function(){drop.classList.remove('over')};\n  drop.ondrop=function(e){e.preventDefault();drop.classList.remove('over');handle(Array.from(e.dataTransfer.files))};\n\n  function addRow(name){\n    var li=document.createElement('li'),a=document.createElement('span'),b=document.createElement('span');\n    a.textContent=name;b.textContent='Waiting\u2026';b.className='wait';\n    li.appendChild(a);li.appendChild(b);list.insertBefore(li,list.firstChild);\n    return function(t,c){b.textContent=t;b.className=c};\n  }\n  function toBase64(file){\n    return new Promise(function(res,rej){\n      var r=new FileReader();\n      r.onload=function(){res(String(r.result).split(',')[1])};\n      r.onerror=function(){rej(new Error('Could not read file'))};\n      r.readAsDataURL(file);\n    });\n  }\n  function send(name,file,b64){\n    var g=gr.value,s=sec.value.trim(),d=dt.value;\n    return new Promise(function(res,rej){\n      google.script.run.withSuccessHandler(res).withFailureHandler(function(e){rej(new Error(e.message||'Upload failed'))})\n        .uploadFile(name,g,s,d,file.name,file.type,b64);\n    });\n  }\n  async function handle(files){\n    var name=nm.value.trim();\n    if(!name){alert('Please type your name first.');nm.focus();return}\n    if(!sec.value.trim()){alert('Please type your section.');sec.focus();return}\n    if(!dt.value){alert('Please choose the date of upload.');dt.focus();return}\n    for(var i=0;i<files.length;i++){\n      var f=files[i],set=addRow(f.name),ext=(f.name.split('.').pop()||'').toLowerCase();\n      if(ALLOWED.indexOf(ext)<0){set('Type not allowed','err');continue}\n      if(f.size>MAX){set('Over 20 MB','err');continue}\n      try{\n        set('Uploading\u2026','wait');\n        await send(name,f,await toBase64(f));\n        set('Submitted \u2714','ok');\n      }catch(e){set(e.message,'err')}\n    }\n  }\n</script>\n</body>\n</html>\n";
