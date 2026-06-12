// ════════════════════════════════════════════════════════════════════════════
// KELLY ALL SYNC — complete bidirectional sync engine
// Google Sheet → Firebase → Client Portal
// Version: 7.1  |  Updated: June 2026
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

// Column indices (0-based) — applies to all checklist tabs
// A=0  B=1  C=2  D=3  E=4  F=5  G=6
// DRC SECTION | SHEET # | DUE DATE | DONE | DATE | CHECKLIST ITEM | NOTES
var COL_SECTION = 0;  // A: section header text (all-caps) or blank for item rows
var COL_SHEET   = 1;  // B: sheet reference
var COL_DONE    = 3;  // D: checkbox (boolean true/false)
var COL_DATE    = 4;  // E: date
var COL_ITEM    = 5;  // F: checklist item text
var COL_COUNTER = 6;  // G: counter formula (G1) — skip on data rows

// ── WEB APP ───────────────────────────────────────────────────────────────────
// doGet: serves agenda JSON to the client portal fetch call.
// Deploy as: Execute as Me, Anyone (even anonymous) can access.

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

  var raw  = latest.getBlob().getDataAsString('UTF-8');
  return parseAgendaDocx(raw, latest.getName());
}

function parseAgendaDocx(text, filename) {
  var lines    = text.split('\n').map(function(l){ return l.trim(); }).filter(Boolean);
  var sections = [];
  var current  = null;

  lines.forEach(function(line) {
    // Section headers: start with two digits and a separator, or all-caps label
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
  // e.g. 260613_KELLY_Weekly_Agenda.docx → Jun 13, 2026
  var m = name.match(/^(\d{2})(\d{2})(\d{2})/);
  if (!m) return '';
  var mo = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return mo[parseInt(m[2]) - 1] + ' ' + parseInt(m[3]) + ', 20' + m[1];
}

// ── ON EDIT HANDLER ───────────────────────────────────────────────────────────
// Fires on every cell edit. Filters to checklist checkbox column only,
// then pushes the single changed key to Firebase immediately.

function onEditHandler(e) {
  try {
    var sheet   = e.range.getSheet();
    var tabName = sheet.getName();
    var prefix  = CHECKLIST_TABS[tabName];
    if (!prefix) return;  // not a checklist tab, ignore

    var editedCol = e.range.getColumn();
    if (editedCol !== COL_DONE + 1) return;  // only care about DONE column (1-indexed)

    var data = sheet.getDataRange().getValues();
    var rowIndex = e.range.getRow() - 1;  // 0-indexed
    var key = buildChecklistKey(data, rowIndex, prefix);
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

// ── CHECKLIST KEY BUILDER ─────────────────────────────────────────────────────
// Scans data from the top to find which section/item the given row belongs to.
// Returns e.g. "PDR_3_12" for prefix "PDR", section index 3, item index 12.

function buildChecklistKey(data, targetRow, prefix) {
  var sectionIndex = -1;
  var itemIndex    = 0;

  for (var r = 1; r < data.length; r++) {  // skip row 0 (header/title)
    var cellA    = String(data[r][COL_SECTION] || '').trim();
    var doneVal  = data[r][COL_DONE];
    var itemText = String(data[r][COL_ITEM]    || '').trim();

    // Section header: col A has content AND col D is not a boolean checkbox value
    var isDoneBoolean = (doneVal === true || doneVal === false);
    var isSection = cellA !== '' && !isDoneBoolean;

    if (isSection) {
      sectionIndex++;
      itemIndex = 0;
      continue;
    }

    if (sectionIndex >= 0 && itemText && isDoneBoolean) {
      if (r === targetRow) return prefix + '_' + sectionIndex + '_' + itemIndex;
      itemIndex++;
    }
  }
  return null;
}

// ── CHECKLIST READER ──────────────────────────────────────────────────────────
// Returns flat object: { PDR_0_0: true, PDR_0_1: false, FDR_0_0: true, ... }

function getChecklistsFromSheet(ss) {
  var result = {};

  Object.keys(CHECKLIST_TABS).forEach(function(tabName) {
    var prefix = CHECKLIST_TABS[tabName];
    var sheet  = ss.getSheetByName(tabName);
    if (!sheet) {
      Logger.log('getChecklistsFromSheet: tab not found — ' + tabName);
      return;
    }

    var data         = sheet.getDataRange().getValues();
    var sectionIndex = -1;
    var itemIndex    = 0;

    for (var r = 1; r < data.length; r++) {
      var cellA    = String(data[r][COL_SECTION] || '').trim();
      var doneVal  = data[r][COL_DONE];
      var itemText = String(data[r][COL_ITEM]    || '').trim();

      var isDoneBoolean = (doneVal === true || doneVal === false);
      var isSection     = cellA !== '' && !isDoneBoolean;

      if (isSection) {
        sectionIndex++;
        itemIndex = 0;
        continue;
      }

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
// Reads SCHEDULE tab. Looks for row where col A = "PHASE" as the header row,
// then builds { phase_name: { week_label: cell_value, ... }, ... }.

function getScheduleFromSheet(ss) {
  // — READ SCHEDULE FROM SHEET TAB —
  var sheet = ss.getSheetByName('SCHEDULE');
  if (!sheet) throw new Error("Tab 'SCHEDULE' not found");

  var data      = sheet.getDataRange().getValues();
  var headerRow = -1;

  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0] || '').trim().toUpperCase() === 'PHASE') {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) throw new Error("'PHASE' header row not found in SCHEDULE tab");

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
// Reads MILESTONES tab (rows 4+). Returns array of milestone objects.
// Status values: done | active | upcoming

function getMilestonesFromSheet(ss) {
  var sheet = ss.getSheetByName('MILESTONES');
  if (!sheet) return null;

  var data       = sheet.getDataRange().getValues();
  var milestones = [];

  // Rows 0-2: title, subtitle, headers — data starts at row 3 (index 3)
  for (var r = 3; r < data.length; r++) {
    var date   = String(data[r][0] || '').trim();
    var label  = String(data[r][1] || '').trim();
    var sub    = String(data[r][2] || '').trim();
    var status = String(data[r][3] || 'upcoming').trim().toLowerCase();
    var notes  = String(data[r][4] || '').trim();
    if (!label) continue;
    milestones.push({ date: date, label: label, sub: sub, status: status, notes: notes });
  }

  Logger.log('getMilestonesFromSheet: ' + milestones.length + ' milestones');
  return milestones;
}

// ── PHASES READER ─────────────────────────────────────────────────────────────
// Reads PHASES tab (rows 4+). Returns array of phase card objects.
// Status values: complete | active | upcoming
// Progress: 0–100 (integer)

function getPhasesFromSheet(ss) {
  var sheet = ss.getSheetByName('PHASES');
  if (!sheet) return null;

  var data   = sheet.getDataRange().getValues();
  var phases = [];

  // Rows 0-2: title, subtitle, headers — data starts at row 3 (index 3)
  for (var r = 3; r < data.length; r++) {
    var num      = String(data[r][0] || '').trim();
    var title    = String(data[r][1] || '').trim();
    var status   = String(data[r][2] || 'upcoming').trim().toLowerCase();
    var text     = String(data[r][3] || '').trim();
    var date     = String(data[r][4] || '').trim();
    var progress = parseInt(data[r][5]) || 0;
    if (!title) continue;
    phases.push({ num: num, title: title, status: status, text: text, date: date, progress: progress });
  }

  Logger.log('getPhasesFromSheet: ' + phases.length + ' phases');
  return phases;
}

// ── FULL SYNC ─────────────────────────────────────────────────────────────────
// Reads all sheet tabs and writes the complete payload to Firebase in one PUT.
// Called by the 5-minute time-driven trigger, and manually when needed.

function pushAllToFirebase() {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var payload = {};

  payload['checklists'] = getChecklistsFromSheet(ss);
  payload['schedule']   = getScheduleFromSheet(ss);

  var milestones = getMilestonesFromSheet(ss);
  if (milestones) payload['milestones'] = milestones;

  var phases = getPhasesFromSheet(ss);
  if (phases) payload['phases'] = phases;

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
// Run setupAllTriggers() ONCE to install both triggers.
// Check Extensions → Apps Script → Triggers before running — re-running
// this function first deletes existing triggers to avoid duplicates.

function setupAllTriggers() {
  // Remove all existing triggers first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // 1. Time-driven: full sync every 5 minutes
  ScriptApp.newTrigger('pushAllToFirebase')
    .timeBased()
    .everyMinutes(5)
    .create();

  // 2. On-edit: immediate single-key push on checkbox toggle
  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(SpreadsheetApp.openById(SHEET_ID))
    .onEdit()
    .create();

  Logger.log('Triggers installed: pushAllToFirebase (5 min) + onEditHandler (on edit)');
}
