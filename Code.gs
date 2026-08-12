// ============================================================
// CID Personnel Progress Tracker — Apps Script backend
// Deploy as a Web App (Extensions > Apps Script in any Google
// account). No Cloud Console or OAuth setup required.
//
// Everything is stored inside ONE Drive folder (set FOLDER_ID
// below): the live JSON data file, the mirrored Google Sheet,
// and any files personnel upload as evidence of submission.
// ============================================================

var FOLDER_ID = '13kZNIGoUAOr97mVfwoDQ7arorsUJXQXr'; // <-- your shared folder
var DATA_FILE_NAME = 'cid_prm_tracker_data.json';
var SHEET_NAME = 'CID PRM Tracker Data';

function getFolder_() {
  return DriveApp.getFolderById(FOLDER_ID);
}

function getDataFile_() {
  var folder = getFolder_();
  var files = folder.getFilesByName(DATA_FILE_NAME);
  if (files.hasNext()) return files.next();
  var blank = JSON.stringify({ people: [], records: {}, timeline: [] });
  return folder.createFile(DATA_FILE_NAME, blank, MimeType.PLAIN_TEXT);
}

function getSpreadsheet_() {
  var folder = getFolder_();
  var files = folder.getFilesByName(SHEET_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  var ss = SpreadsheetApp.create(SHEET_NAME);
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file); // move it out of My Drive root, into the folder
  return ss;
}

function doGet(e) {
  var file = getDataFile_();
  var data = JSON.parse(file.getBlob().getDataAsString());
  var ss = getSpreadsheet_();
  data._sheetUrl = ss.getUrl();
  data._folderUrl = getFolder_().getUrl();
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);

  if (payload.type === 'upload') {
    return handleUpload_(payload);
  }

  // default: save the tracker data
  var file = getDataFile_();
  file.setContent(JSON.stringify({
    people: payload.people,
    records: payload.records,
    timeline: payload.timeline
  }));
  var ss = writeToSheet_(payload);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, sheetUrl: ss.getUrl() }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleUpload_(payload) {
  var folder = getFolder_();
  var bytes = Utilities.base64Decode(payload.base64);
  var blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream', payload.fileName || 'upload');
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return ContentService.createTextOutput(JSON.stringify({
    ok: true, url: file.getUrl(), id: file.getId(), name: file.getName()
  })).setMimeType(ContentService.MimeType.JSON);
}

function writeToSheet_(data) {
  var ss = getSpreadsheet_();
  var people = data.people || [];
  var records = data.records || {};
  var timeline = data.timeline || [];

  // --- Personnel Progress tab ---
  var sheet = ss.getSheetByName('Personnel Progress') || ss.insertSheet('Personnel Progress');
  sheet.clear();
  sheet.appendRow([
    'Name', 'Position',
    'IPCRF/OPCRF Req', 'IPCRF/OPCRF Done',
    'QMS-QCP Req', 'QMS-QCP Done',
    'DEDP-AIP Req', 'DEDP-AIP Done',
    '% Done', 'Status', 'Last Submitted', 'Remarks', 'Attachments'
  ]);
  people.forEach(function (p) {
    var r = records[p.id] || {};
    var docs = r.docs || {};
    var ipcrf = docs.ipcrf || { req: 0, done: 0 };
    var qcp = docs.qcp || { req: 0, done: 0 };
    var aip = docs.aip || { req: 0, done: 0 };
    var totalReq = (Number(ipcrf.req) || 0) + (Number(qcp.req) || 0) + (Number(aip.req) || 0);
    var totalDone = (Number(ipcrf.done) || 0) + (Number(qcp.done) || 0) + (Number(aip.done) || 0);
    var pct = totalReq > 0 ? Math.round((totalDone / totalReq) * 1000) / 10 : 0;
    var attachments = (r.attachments || []).map(function (a) { return a.url; }).join('\n');
    sheet.appendRow([
      p.name, p.role || '',
      ipcrf.req || 0, ipcrf.done || 0,
      qcp.req || 0, qcp.done || 0,
      aip.req || 0, aip.done || 0,
      pct + '%',
      r.submitted ? 'Submitted' : 'Not yet submitted',
      r.lastSubmittedAt || '',
      r.remarks || '',
      attachments
    ]);
  });
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 13);

  // --- Timeline tab ---
  var tSheet = ss.getSheetByName('Timeline') || ss.insertSheet('Timeline');
  tSheet.clear();
  tSheet.appendRow(['Timestamp', 'Person', 'Event']);
  var sorted = timeline.slice().sort(function (a, b) { return (b.ts || '').localeCompare(a.ts || ''); });
  sorted.forEach(function (t) {
    tSheet.appendRow([t.ts, t.personName, t.event]);
  });
  tSheet.setFrozenRows(1);
  tSheet.autoResizeColumns(1, 3);

  return ss;
}
