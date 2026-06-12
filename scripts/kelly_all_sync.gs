// ════════════════════════════════════════════════════════════════════════════
// KELLY ALL SYNC v7.2
// Google Sheet → Firebase → Client Portal
// ════════════════════════════════════════════════════════════════════════════

var FIREBASE         = "https://kelly-deliverables-default-rtdb.firebaseio.com";
var SHEET_ID         = "1PvtRF8jVLn5vA1AmdXDMBkbMdZ2un91n1zDwCKzIQck";
var FB_PATH          = "kelly_271_v7";
var AGENDA_FOLDER_ID = "1EBFJdEU7peCgDoIMVc478BSL8_GE5Mb6";

// Checklist tabs → Firebase key prefix
var CHECKLIST_TABS = {
  'PRELIMINARY DR': 'PDR',
  'FINAL DR':       'FDR',
  'STATE PERMIT':   'STATE',
  'LOCAL PERMIT':   'LOCAL'
};

// Column indices (0-based)
// A=0  B=1  C=2  D=3  E=4  F=5  G=6
// DRC SECTION | SHEET # | DUE DATE | DONE | DATE | CHECKLIST ITEM | NOTES
var COL_SECTION = 0;
var COL_DONE    = 3;
var COL_DATE    = 4;
var COL_ITEM    = 5;

// ── WEB APP ───────────────────────────────────────────────────────────────────

function doGet(e) {
  var data, output;
  try {
    data = getLatestAgendaFromDrive();
  } catch (err) {
    data = { error: err.toString() };
  }
  output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function getLatestAgendaFromDrive() {
  var folder = DriveApp.getFolderById(AGENDA_FOLDER_ID);
  var files  = folder.getFilesByType(MimeType.MICROSOFT_WORD);
  var latest = null, latestDate = null;

  while (files.hasNext()) {
    var f = files.next();
    var d = f.getLastUpdated();
    if (!latestDate || d > latestDate) { latestDate = d; latest = f; }
  }
  if (!latest) return {};
  var raw = latest.getBlob().getDataAsString('UTF-8');
  return parseAgendaDocx(raw, latest.getName());
}

function parseAgendaDocx(text, filename) {
  var lines    = text.split('\n').map(function(l){ return l.trim(); }).filter(Boolean);
  var sections = [];
  var current  = null;

  lines.forEach(function(line) {
    if (/^\d{2}\s[·\-–]/.test(line) || /^[A-Z][A-Z\s\/]+$/.test(line)) {
      if (current) sections.push(current);
      current = { title: line, items: [] };
    } else if (current && line.length > 3) {
      current.items.push({ text: line });
    }
  });
  if (current) sections.push(current);

  return {
    filename:  filename,
    date:      agendaDateFromFilename(filename),
    attendees: 'Toby Long AIA · André Mandel · Chandra Baerg · Abby Wittman · Mai Thor',
    sections:  sections
  };
}

function agendaDateFromFilename(name) {
  var m = name.match(/^(\d{2})(\d{2})(\d{2})/);
  if (!m) return '';
  var mo = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return mo[parseInt(m[2]) - 1] + ' ' + parseInt(m[3]) + ', 20' + m[1];
}

// ── ON EDIT HANDLER ───────────────────────────────────────────────────────────

function onEditHandler(e) {
  try {
    var sheet   = e.range.getSheet();
    var tabName = sheet.getName();
    var prefix  = CHECKLIST_TABS[tabName];
    if (!prefix) return;

    var editedCol = e.range.getColumn();
    if (editedCol !== COL_DONE + 1) return;

    var data     = sheet.getDataRange().getValues();
    var rowIndex = e.range.getRow() - 1;
    var key      = buildChecklistKey(data, rowIndex, prefix);
    if (!key) return;

    var val = (e.value === 'TRUE' || e.value === true);
    var url = FIREBASE + '/' + FB_PATH + '/checklists/' + key + '.json';
    UrlFetchApp.fetch(url, {
      method: 'put',
      contentType: 'application/json',
      payload: JSON.stringify(val),
      muteHttpExceptions: true
    });
    Logger.log('onEditHandler pushed: ' + key + ' = ' + val);
  } catch (err) {
    Logger.log('onEditHandler error: ' + err);
  }
}

function buildChecklistKey(data, targetRow, prefix) {
  var sectionIndex = -1;
  var itemIndex    = 0;

  for (var r = 1; r < data.length; r++) {
    var cellA       = String(data[r][COL_SECTION] || '').trim();
    var doneVal     = data[r][COL_DONE];
    var itemText    = String(data[r][COL_ITEM]    || '').trim();
    var isDoneBoolean = (doneVal === true || doneVal === false);
    var isSection     = cellA !== '' && !isDoneBoolean;

    if (isSection) { sectionIndex++; itemIndex = 0; continue; }
    if (sectionIndex >= 0 && itemText && isDoneBoolean) {
      if (r === targetRow) return prefix + '_' + sectionIndex + '_' + itemIndex;
      itemIndex++;
    }
  }
  return null;
}

// ── CHECKLIST READER ──────────────────────────────────────────────────────────

function getChecklistsFromSheet(ss) {
  var result = {};

  Object.keys(CHECKLIST_TABS).forEach(function(tabName) {
    var prefix = CHECKLIST_TABS[tabName];
    var sheet  = ss.getSheetByName(tabName);
    if (!sheet) { Logger.log('Tab not found: ' + tabName); return; }

    var data         = sheet.getDataRange().getValues();
    var sectionIndex = -1;
    var itemIndex    = 0;

    for (var r = 1; r < data.length; r++) {
      var cellA       = String(data[r][COL_SECTION] || '').trim();
      var doneVal     = data[r][COL_DONE];
      var itemText    = String(data[r][COL_ITEM]    || '').trim();
      var isDoneBoolean = (doneVal === true || doneVal === false);
      var isSection     = cellA !== '' && !isDoneBoolean;

      if (isSection) { sectionIndex++; itemIndex = 0; continue; }
      if (sectionIndex >= 0 && itemText && isDoneBoolean) {
        result[prefix + '_' + sectionIndex + '_' + itemIndex] = (doneVal === true);
        itemIndex++;
      }
    }
  });

  Logger.log('getChecklistsFromSheet: ' + Object.keys(result).length + ' keys');
  return result;
}

// ── SCHEDULE READER ───────────────────────────────────────────────────────────

function getScheduleFromSheet(ss) {
  // — READ SCHEDULE FROM SHEET TAB —
  var sheet = ss.getSheetByName('SCHEDULE');
  if (!sheet) throw new Error("Tab 'SCHEDULE' not found");

  var data      = sheet.getDataRange().getValues();
  var headerRow = -1;

  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0] || '').trim().toUpperCase() === 'PHASE') {
      headerRow = r; break;
    }
  }
  if (headerRow < 0) throw new Error("'PHASE' header row not found");

  var headers = data[headerRow].map(function(h) { return String(h || '').trim(); });
  var result  = {};

  for (var r = headerRow + 1; r < data.length; r++) {
    var phase = String(data[r][0] || '').trim();
    if (!phase) continue;
    result[phase] = {};
    for (var c = 1; c < headers.length; c++) {
      if (headers[c]) result[phase][headers[c]] = String(data[r][c] || '').trim();
    }
  }

  Logger.log('getScheduleFromSheet: ' + Object.keys(result).length + ' phases');
  return result;
}

// ── MILESTONES READER ─────────────────────────────────────────────────────────

function getMilestonesFromSheet(ss) {
  var sheet = ss.getSheetByName('MILESTONES');
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var out  = [];

  for (var r = 3; r < data.length; r++) {  // rows 0-2 are title/sub/headers
    var label = String(data[r][1] || '').trim();
    if (!label) continue;
    out.push({
      date:   String(data[r][0] || '').trim(),
      label:  label,
      sub:    String(data[r][2] || '').trim(),
      status: String(data[r][3] || 'upcoming').trim().toLowerCase(),
      notes:  String(data[r][4] || '').trim()
    });
  }

  Logger.log('getMilestonesFromSheet: ' + out.length);
  return out;
}

// ── PHASES READER ─────────────────────────────────────────────────────────────
// Writes to two Firebase paths that the portal's existing phase JS reads:
//   kelly_271_v7/phase_status  → { 0: {status:"complete"}, 1: {status:"active"}, ... }
//   kelly_271_v7/phase_fields  → { phase0_status:"text", phase0_date:"date", ... }
//   kelly_271_v7/phase_meta    → { 0: {num:"Phase 01", title:"PDR...", progress:100}, ... }

function getPhasesFromSheet(ss) {
  var sheet = ss.getSheetByName('PHASES');
  if (!sheet) return null;

  var data       = sheet.getDataRange().getValues();
  var phaseStatus = {};
  var phaseFields = {};
  var phaseMeta   = {};

  // Rows 0-2: title/subtitle/headers — data starts row 3 (index 3)
  for (var r = 3; r < data.length; r++) {
    var num      = String(data[r][0] || '').trim();
    var title    = String(data[r][1] || '').trim();
    var status   = String(data[r][2] || 'upcoming').trim().toLowerCase();
    var text     = String(data[r][3] || '').trim();
    var date     = String(data[r][4] || '').trim();
    var progress = parseInt(data[r][5]) || 0;
    if (!title) continue;

    var idx = r - 3;  // 0-based phase index

    // Path 1: phase_status (drives pill selector and card data-status)
    phaseStatus[idx] = { status: status };

    // Path 2: phase_fields (drives contenteditable text + date fields)
    phaseFields['phase' + idx + '_status'] = text;
    phaseFields['phase' + idx + '_date']   = date;

    // Path 3: phase_meta (title, num, progress — for display)
    phaseMeta[idx] = { num: num, title: title, progress: progress };
  }

  Logger.log('getPhasesFromSheet: ' + Object.keys(phaseStatus).length + ' phases');
  return { phase_status: phaseStatus, phase_fields: phaseFields, phase_meta: phaseMeta };
}

// ── FULL SYNC ─────────────────────────────────────────────────────────────────

function pushAllToFirebase() {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var payload = {};

  payload['checklists'] = getChecklistsFromSheet(ss);
  payload['schedule']   = getScheduleFromSheet(ss);

  var milestones = getMilestonesFromSheet(ss);
  if (milestones) payload['milestones'] = milestones;

  var phases = getPhasesFromSheet(ss);
  if (phases) {
    payload['phase_status'] = phases.phase_status;
    payload['phase_fields'] = phases.phase_fields;
    payload['phase_meta']   = phases.phase_meta;
  }

  var url      = FIREBASE + '/' + FB_PATH + '.json';
  var response = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('pushAllToFirebase: HTTP ' + response.getResponseCode());
}

// ── TRIGGER SETUP ─────────────────────────────────────────────────────────────
// Run ONCE. Deletes existing triggers first to prevent duplicates.

function setupAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('pushAllToFirebase')
    .timeBased().everyMinutes(5).create();

  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(SpreadsheetApp.openById(SHEET_ID))
    .onEdit().create();

  Logger.log('Triggers set: pushAllToFirebase (5 min) + onEditHandler');
}
