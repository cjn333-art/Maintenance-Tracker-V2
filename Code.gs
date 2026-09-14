/**
 * Stacor Maintenance — Apps Script backend
 *
 * SETUP
 * 1. Create a new Google Sheet (any name, e.g. "Stacor Maintenance Data").
 * 2. In the Sheet, go to Extensions > Apps Script.
 * 3. Delete any starter code and paste this whole file in.
 * 4. Click Deploy > New deployment.
 *      - Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone (or "Anyone with the link")
 * 5. Click Deploy, authorize the script when prompted, and copy the Web App URL.
 * 6. Paste that URL into the API_URL constant near the top of the <script>
 *    block in BOTH maintenance-time-log.html and task-request.html.
 *
 * This script creates its own "Entries" and "Pending" sheet tabs
 * automatically the first time it runs — you don't need to set them up
 * by hand. You can still open them any time to see the raw data.
 */

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'entries') return jsonResponse(getEntries());
  if (action === 'pending') return jsonResponse(getPending());
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ success: false, error: 'Bad request body' });
  }

  const action = body.action;
  let result;
  if (action === 'addEntry') result = addEntry(body.entry);
  else if (action === 'updateEntry') result = updateEntry(body.id, body.updates);
  else if (action === 'addPending') result = addPending(body.task);
  else if (action === 'deletePending') result = deletePending(body.id);
  else result = { success: false, error: 'Unknown action' };

  return jsonResponse(result);
}

/* ---------- Entries (Job Log) ---------- */
// Columns: id | department | crew (JSON) | description | blocks (JSON) | completed | manHours

const ENTRIES_HEADER = ['id', 'department', 'crew', 'description', 'blocks', 'completed', 'manHours'];

function getEntries() {
  const sheet = getSheet('Entries');
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  return rows
    .filter(r => r[0])
    .map(r => ({
      id: String(r[0]),
      department: r[1],
      crew: r[2] ? JSON.parse(r[2]) : [],
      description: r[3],
      blocks: r[4] ? JSON.parse(r[4]) : [],
      completed: r[5] === true || r[5] === 'TRUE',
      manHours: r[6] === '' || r[6] === null || r[6] === undefined ? null : Number(r[6])
    }));
}

function addEntry(entry) {
  const sheet = getSheet('Entries');
  if (sheet.getLastRow() === 0) sheet.appendRow(ENTRIES_HEADER);
  sheet.appendRow([
    entry.id,
    entry.department,
    JSON.stringify(entry.crew || []),
    entry.description,
    JSON.stringify(entry.blocks || []),
    entry.completed === true,
    entry.manHours != null ? entry.manHours : ''
  ]);
  return { success: true };
}

function updateEntry(id, updates) {
  const sheet = getSheet('Entries');
  if (sheet.getLastRow() < 2) return { success: false, error: 'Not found' };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const row = i + 1;
      if (updates.blocks !== undefined) sheet.getRange(row, 5).setValue(JSON.stringify(updates.blocks));
      if (updates.completed !== undefined) sheet.getRange(row, 6).setValue(updates.completed === true);
      if (updates.manHours !== undefined) sheet.getRange(row, 7).setValue(updates.manHours != null ? updates.manHours : '');
      return { success: true };
    }
  }
  return { success: false, error: 'Not found' };
}

/* ---------- Pending (Tasks awaiting assignment) ---------- */
// Columns: id | department | requestedBy | type | tankEmpty | otherDetail | description | urgent | createdAt

const PENDING_HEADER = ['id', 'department', 'requestedBy', 'type', 'tankEmpty', 'otherDetail', 'description', 'urgent', 'createdAt'];

function getPending() {
  const sheet = getSheet('Pending');
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  return rows
    .filter(r => r[0])
    .map(r => ({
      id: String(r[0]),
      department: r[1],
      requestedBy: r[2],
      type: r[3],
      tankEmpty: r[4] || null,
      otherDetail: r[5] || null,
      description: r[6] || '',
      urgent: r[7] === true || r[7] === 'TRUE',
      createdAt: r[8]
    }));
}

function addPending(task) {
  const sheet = getSheet('Pending');
  if (sheet.getLastRow() === 0) sheet.appendRow(PENDING_HEADER);
  sheet.appendRow([
    task.id,
    task.department,
    task.requestedBy,
    task.type,
    task.tankEmpty || '',
    task.otherDetail || '',
    task.description || '',
    task.urgent === true,
    task.createdAt
  ]);
  return { success: true };
}

function deletePending(id) {
  const sheet = getSheet('Pending');
  if (sheet.getLastRow() < 2) return { success: false, error: 'Not found' };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Not found' };
}
