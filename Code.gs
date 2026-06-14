/**
 * RBS Evaluation System - Google Apps Script backend
 *
 * Updated: 1 row = 1 job, split columns A-Q
 * Columns: Timestamp | Phone | Truck | InstDate | Name | Customer | CustType | JobType | Location | Region | Accuracy | Timeliness | Professionalism | Cleanliness | Communication | Comments | AvgScore
 */

var USERS_SHEET_ID = '1HkcfuUZ27AWZXi8OBXoEW7C5n1dInFMWsSTSCCCxhng';
var SUBMISSIONS_SHEET_ID = '1HkcfuUZ27AWZXi8OBXoEW7C5n1dInFMWsSTSCCCxhng';
var TOKEN_TTL_MS = 8 * 60 * 60 * 1000; // 8 hours

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var action = body.action;

    if (action === 'adminLogin') {
      return handleAdminLogin(body);
    }

    if (action === 'addUser' || action === 'deleteUser' ||
        action === 'getSubmissionDetail' || action === 'deleteSubmissionFromSheet') {
      if (!verifyToken(body.token)) {
        return json({ success: false, error: 'Unauthorized' });
      }
      if (action === 'addUser')                   return handleAddUser(body);
      if (action === 'deleteUser')                return handleDeleteUser(body);
      if (action === 'getSubmissionDetail')       return handleGetSubmissionDetail(body);
      if (action === 'deleteSubmissionFromSheet') {
        // Support both timestamp and id parameters
        if (body.timestamp) {
          return handleDeleteSubmissionByTimestamp(body);
        } else {
          return handleDeleteSubmissionFromSheet(body);
        }
      }
    }

    return handleSubmission(body);

  } catch (err) {
    return json({ status: 'error', success: false, message: String(err) });
  }
}

function handleAdminLogin(body) {
  var props = PropertiesService.getScriptProperties();
  var expectedId = props.getProperty('ADMIN_ID');
  var expectedPw = props.getProperty('ADMIN_PASSWORD');

  if (!expectedId || !expectedPw) {
    return json({ success: false, error: 'Admin credentials not configured' });
  }

  if (String(body.adminId) === expectedId && String(body.password) === expectedPw) {
    return json({ success: true, token: issueToken('admin'), name: 'Admin' });
  }
  return json({ success: false, error: 'Invalid credentials' });
}

function issueToken(subject) {
  var secret = PropertiesService.getScriptProperties().getProperty('TOKEN_SECRET') || 'change-me';
  var payload = subject + '|' + (Date.now() + TOKEN_TTL_MS);
  var sig = sign(payload, secret);
  return Utilities.base64EncodeWebSafe(payload + '|' + sig);
}

function verifyToken(token) {
  if (!token) return false;
  try {
    var secret = PropertiesService.getScriptProperties().getProperty('TOKEN_SECRET') || 'change-me';
    var decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(token)).getDataAsString();
    var parts = decoded.split('|');
    if (parts.length !== 3) return false;
    var payload = parts[0] + '|' + parts[1];
    var expiry = parseInt(parts[1], 10);
    if (isNaN(expiry) || Date.now() > expiry) return false;
    return sign(payload, secret) === parts[2];
  } catch (e) {
    return false;
  }
}

function sign(data, secret) {
  var raw = Utilities.computeHmacSha256Signature(data, secret);
  return Utilities.base64EncodeWebSafe(raw);
}

function handleAddUser(body) {
  var spreadsheet = SpreadsheetApp.openById(USERS_SHEET_ID);
  var sheet = spreadsheet.getSheetByName('Users');
  if (!sheet) {
    sheet = spreadsheet.getSheets()[1]; // Fallback to 2nd sheet if Users sheet not found
  }
  sheet.appendRow([body.name || '', String(body.phone || ''), body.timestamp || new Date()]);
  return json({ success: true });
}

function handleDeleteUser(body) {
  var spreadsheet = SpreadsheetApp.openById(USERS_SHEET_ID);
  var sheet = spreadsheet.getSheetByName('Users');
  if (!sheet) {
    sheet = spreadsheet.getSheets()[1]; // Fallback to 2nd sheet if Users sheet not found
  }
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(body.phone)) {
      sheet.deleteRow(i + 1);
    }
  }
  return json({ success: true });
}

function handleSubmission(body) {
  var sheet = SpreadsheetApp.openById(SUBMISSIONS_SHEET_ID).getSheets()[0];

  // Force column B to Text format
  sheet.getRange('B:B').setNumberFormat('@');

  var jobs = body.jobs || [];
  var baseId = body.id || Date.now();

  // Format phone: ensure 0 prefix
  var phoneStr = String(body.phone || '').trim();
  if (phoneStr && !phoneStr.startsWith('0')) {
    phoneStr = '0' + phoneStr;
  }

  for (var i = 0; i < jobs.length; i++) {
    var job = jobs[i];
    var jobId = baseId + '-' + i;

    var scores = job.scores || {};
    var accuracy = scores.accuracy || 0;
    var timeliness = scores.timeliness || 0;
    var professionalism = scores.professionalism || 0;
    var cleanliness = scores.cleanliness || 0;
    var communication = scores.communication || 0;

    var scoreValues = [accuracy, timeliness, professionalism, cleanliness, communication];
    var avgScore = scoreValues.length > 0 ? (scoreValues.reduce((a, b) => a + b, 0) / scoreValues.length).toFixed(1) : 0;

    // Ensure region is included in the submission
    var region = job.region || '';

    sheet.appendRow([
      body.timestamp || new Date(),           // A: Timestamp
      phoneStr,                               // B: Phone (text with 0 prefix)
      body.truckNumber || '',                 // C: Truck Number
      body.installationDate || '',            // D: Installation Date
      body.name || '',                        // E: Name
      job.customer || '',                     // F: Customer
      job.customerType || '',                 // G: Customer Type
      job.type || '',                         // H: Job Type
      job.location || '',                     // I: Location
      region,                                 // J: Region
      accuracy,                               // K: Accuracy
      timeliness,                             // L: Timeliness
      professionalism,                        // M: Professionalism
      cleanliness,                            // N: Cleanliness
      communication,                          // O: Communication
      job.comments || '',                     // P: Comments
      avgScore                                // Q: Average Score
    ]);
  }

  return json({ status: 'success', success: true, id: baseId });
}

function handleGetSubmissionDetail(body) {
  var sheet = SpreadsheetApp.openById(SUBMISSIONS_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(body.id)) {  // Match by phone (column B)
      var detail = {
        timestamp: data[i][0],      // A
        phone: data[i][1],          // B
        truckNumber: data[i][2],    // C
        installationDate: data[i][3], // D
        name: data[i][4],           // E
        customer: data[i][5],       // F
        customerType: data[i][6],   // G
        type: data[i][7],           // H
        location: data[i][8],       // I
        region: data[i][9],         // J
        scores: {
          accuracy: data[i][10],    // K
          timeliness: data[i][11],  // L
          professionalism: data[i][12], // M
          cleanliness: data[i][13], // N
          communication: data[i][14] // O
        },
        comments: data[i][15],      // P
        avgScore: data[i][16]       // Q
      };
      return json({ success: true, submission: detail });
    }
  }
  return json({ success: false, error: 'Not found' });
}

function handleDeleteSubmissionFromSheet(body) {
  var sheet = SpreadsheetApp.openById(SUBMISSIONS_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var baseId = String(body.id).split('-')[0]; // Extract baseId from ID
  var deletedCount = 0;

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).includes(baseId)) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }

  return deletedCount > 0
    ? json({ success: true, message: 'Deleted ' + deletedCount + ' row(s)' })
    : json({ success: false, error: 'Not found' });
}

function handleDeleteSubmissionByTimestamp(body) {
  var sheet = SpreadsheetApp.openById(SUBMISSIONS_SHEET_ID).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var timestamp = String(body.timestamp);
  var deletedCount = 0;

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === timestamp) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }

  return deletedCount > 0
    ? json({ success: true, message: 'Deleted ' + deletedCount + ' row(s)' })
    : json({ success: false, error: 'Not found' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
