// ════════════════════════════════════════════════════════════════════════════
// KELLY ALL SYNC v7.3
// Google Sheet → Firebase → Client Portal
// Tabs: CONFIG · MILESTONES · PHASES · PRELIMINARY DR · FINAL DR ·
//        STATE PERMIT · LOCAL PERMIT · SCHEDULE
// ════════════════════════════════════════════════════════════════════════════

var FIREBASE         = "https://kelly-deliverables-default-rtdb.firebaseio.com";
var SHEET_ID         = "1PvtRF8jVLn5vA1AmdXDMBkbMdZ2un91n1zDwCKzIQck";
var FB_PATH          = "kelly_271_v7";
var AGENDA_FOLDER_ID = "1EBFJdEU7peCgDoIMVc478BSL8_GE5Mb6";

var CHECKLIST_TABS = {
  'PRELIMINARY DR': 'PDR',
  'FINAL DR':       'FDR',
  'STATE PERMIT':   'STATE',
  'LOCAL PERMIT':   'LOCAL'
};

var COL_SECTION = 0;
var COL_DONE    = 3;
var COL_ITEM    = 5;

// ── KEY SANITIZER ────────────────────────────────────────────────────────────
// Firebase keys cannot contain . # $ / [ ] or be empty.
// Replace all invalid chars with _ so writes never 400.
function sanitizeKey(k) {
  return String(k || 'blank').replace(/[.#$\/\[\]]/g, '_').trim() || 'blank';
}

// ── WEB APP ───────────────────────────────────────────────────────────────────

function doGet(e) {
  var data;
  try { data = getLatestAgendaFromDrive(); }
  catch (err) { data = { error: err.toString() }; }
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function getLatestAgendaFromDrive() {
  var folder = DriveApp.getFolderById(AGENDA_FOLDER_ID);
  var files  = folder.getFilesByType(MimeType.MICROSOFT_WORD);
  var latest = null, latestDate = null;
  while (files.hasNext()) {
    var f = files.next(), d = f.getLastUpdated();
    if (!latestDate || d > latestDate) { latestDate = d; latest = f; }
  }
  if (!latest) return {};
  return parseAgendaDocx(latest.getBlob().getDataAsString('UTF-8'), latest.getName());
}

function parseAgendaDocx(text, filename) {
  var lines = text.split('\n').map(function(l){ return l.trim(); }).filter(Boolean);
  var sections = [], current = null;
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
    if (e.range.getColumn() !== COL_DONE + 1) return;
    var data = sheet.getDataRange().getValues();
    var key  = buildChecklistKey(data, e.range.getRow() - 1, prefix);
    if (!key) return;
    var val = (e.value === 'TRUE' || e.value === true);
    UrlFetchApp.fetch(FIREBASE + '/' + FB_PATH + '/checklists/' + key + '.json', {
      method: 'put', contentType: 'application/json',
      payload: JSON.stringify(val), muteHttpExceptions: true
    });
    Logger.log('onEditHandler: ' + key + ' = ' + val);
  } catch (err) { Logger.log('onEditHandler error: ' + err); }
}

function buildChecklistKey(data, targetRow, prefix) {
  var sectionIndex = -1, itemIndex = 0;
  for (var r = 1; r < data.length; r++) {
    var cellA = String(data[r][COL_SECTION] || '').trim();
    var done  = data[r][COL_DONE];
    var item  = String(data[r][COL_ITEM]    || '').trim();
    var isBool    = (done === true || done === false);
    var isSection = cellA !== '' && !isBool;
    if (isSection) { sectionIndex++; itemIndex = 0; continue; }
    if (sectionIndex >= 0 && item && isBool) {
      if (r === targetRow) return prefix + '_' + sectionIndex + '_' + itemIndex;
      itemIndex++;
    }
  }
  return null;
}

// ── CONFIG READER ─────────────────────────────────────────────────────────────
// Reads CONFIG tab. Rows 0-2 = title/subtitle/headers. Data starts row 3.
// Section header rows (grey, no value in col B) are skipped.
// Returns { current_phase: "PDR Submitted", plans_pdf_url: "...", ... }

function getConfigFromSheet(ss) {
  var sheet = ss.getSheetByName('CONFIG');
  if (!sheet) return null;
  var data   = sheet.getDataRange().getValues();
  var config = {};
  for (var r = 3; r < data.length; r++) {
    var key = String(data[r][0] || '').trim();
    var val = String(data[r][1] || '').trim();
    // Skip section headers (start with —) and blank rows
    if (!key || key.startsWith('—') || !val) continue;
    config[sanitizeKey(key)] = val;
  }
  Logger.log('getConfigFromSheet: ' + Object.keys(config).length + ' keys');
  return config;
}

// ── CHECKLIST READER ──────────────────────────────────────────────────────────

function getChecklistsFromSheet(ss) {
  var result = {};
  Object.keys(CHECKLIST_TABS).forEach(function(tabName) {
    var prefix = CHECKLIST_TABS[tabName];
    var sheet  = ss.getSheetByName(tabName);
    if (!sheet) { Logger.log('Tab not found: ' + tabName); return; }
    var data = sheet.getDataRange().getValues();
    var si = -1, ii = 0;
    for (var r = 1; r < data.length; r++) {
      var cellA = String(data[r][COL_SECTION] || '').trim();
      var done  = data[r][COL_DONE];
      var item  = String(data[r][COL_ITEM]    || '').trim();
      var isBool    = (done === true || done === false);
      var isSection = cellA !== '' && !isBool;
      if (isSection) { si++; ii = 0; continue; }
      if (si >= 0 && item && isBool) { result[prefix+'_'+si+'_'+ii] = (done === true); ii++; }
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
  var data = sheet.getDataRange().getValues();
  var headerRow = -1;
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0] || '').trim().toUpperCase() === 'PHASE') { headerRow = r; break; }
  }
  if (headerRow < 0) throw new Error("'PHASE' header row not found");
  var headers = data[headerRow].map(function(h){ return String(h||'').trim(); });
  var result  = {};
  for (var r = headerRow + 1; r < data.length; r++) {
    var phase = sanitizeKey(String(data[r][0] || '').trim());
    if (!phase || phase === 'blank') continue;
    result[phase] = {};
    for (var c = 1; c < headers.length; c++) {
      var hk = sanitizeKey(headers[c]);
      if (hk && hk !== 'blank') result[phase][hk] = String(data[r][c] || '').trim();
    }
  }
  Logger.log('getScheduleFromSheet: ' + Object.keys(result).length + ' phases');
  return result;
}

// ── MILESTONES READER ─────────────────────────────────────────────────────────

function getMilestonesFromSheet(ss) {
  var sheet = ss.getSheetByName('MILESTONES');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues(), out = [];
  for (var r = 3; r < data.length; r++) {
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
// Writes to paths the portal's existing phase JS reads from:
//   phase_status: { 0: {status:"complete"}, 1: {status:"active"}, ... }
//   phase_fields: { phase0_status: "text", phase0_date: "date", ... }
//   phase_meta:   { 0: {num:"Phase 01", title:"PDR...", progress:100}, ... }

function getPhasesFromSheet(ss) {
  var sheet = ss.getSheetByName('PHASES');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var phaseStatus = {}, phaseFields = {}, phaseMeta = {};
  for (var r = 3; r < data.length; r++) {
    var num      = String(data[r][0] || '').trim();
    var title    = String(data[r][1] || '').trim();
    var status   = String(data[r][2] || 'upcoming').trim().toLowerCase();
    var text     = String(data[r][3] || '').trim();
    var date     = String(data[r][4] || '').trim();
    var progress = parseInt(data[r][5]) || 0;
    if (!title) continue;
    var idx = r - 3;
    phaseStatus[idx] = { status: status };
    phaseFields['phase' + idx + '_status'] = text;
    phaseFields['phase' + idx + '_date']   = date;
    phaseMeta[idx]   = { num: num, title: title, progress: progress };
  }
  Logger.log('getPhasesFromSheet: ' + Object.keys(phaseStatus).length + ' phases');
  return { phase_status: phaseStatus, phase_fields: phaseFields, phase_meta: phaseMeta };
}

// ── FULL SYNC ─────────────────────────────────────────────────────────────────

function pushAllToFirebase() {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var payload = {};

  // Config
  var config = getConfigFromSheet(ss);
  if (config) payload['config'] = config;

  // Checklists
  payload['checklists'] = getChecklistsFromSheet(ss);

  // Schedule (Gantt / timeline)
  payload['schedule'] = getScheduleFromSheet(ss);

  // Milestones
  var milestones = getMilestonesFromSheet(ss);
  if (milestones) payload['milestones'] = milestones;

  // Phase cards
  var phases = getPhasesFromSheet(ss);
  if (phases) {
    payload['phase_status'] = phases.phase_status;
    payload['phase_fields'] = phases.phase_fields;
    payload['phase_meta']   = phases.phase_meta;
  }

  var response = UrlFetchApp.fetch(FIREBASE + '/' + FB_PATH + '.json', {
    method: 'put', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  Logger.log('pushAllToFirebase: HTTP ' + response.getResponseCode());
}

// ── TRIGGER SETUP ─────────────────────────────────────────────────────────────
// Run ONCE. Deletes existing triggers first to prevent duplicates.

function setupAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('pushAllToFirebase').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(SpreadsheetApp.openById(SHEET_ID)).onEdit().create();
  Logger.log('Triggers set: pushAllToFirebase (5 min) + onEditHandler');
}
/**
 * createConfigTab()
 * Creates a CONFIG sheet as the first tab in the Kelly matrix.
 * Staff edit this tab to control key portal elements from the spreadsheet.
 * Safe to re-run — clears and re-seeds if tab already exists.
 */
function createConfigTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var existing = ss.getSheetByName("CONFIG");
  if (existing) ss.deleteSheet(existing);

  var sh = ss.insertSheet("CONFIG", 0);  // first tab

  // ── Column widths ──────────────────────────────────────────────────────────
  sh.setColumnWidth(1, 220);  // A: Field label
  sh.setColumnWidth(2, 500);  // B: Value
  sh.setColumnWidth(3, 320);  // C: Notes

  // ── Row 1: Title ───────────────────────────────────────────────────────────
  sh.setRowHeight(1, 44);
  var titleRange = sh.getRange("A1:C1");
  titleRange.merge();
  titleRange.setValue("KELLY RESIDENCE — PORTAL CONFIG");
  titleRange.setBackground("#F4DDC8");
  titleRange.setFontFamily("Inter");
  titleRange.setFontSize(13);
  titleRange.setFontWeight("bold");
  titleRange.setFontColor("#1A1A1A");
  titleRange.setVerticalAlignment("middle");

  // ── Row 2: Subtitle ────────────────────────────────────────────────────────
  sh.setRowHeight(2, 22);
  var sub = sh.getRange("A2:C2");
  sub.merge();
  sub.setValue("Edit values in column B to update the client portal. Changes sync within 5 minutes.");
  sub.setBackground("#F4DDC8");
  sub.setFontFamily("Inter");
  sub.setFontSize(9);
  sub.setFontStyle("italic");
  sub.setFontColor("#4A4A4A");
  sub.setVerticalAlignment("middle");

  // ── Row 3: Column headers ──────────────────────────────────────────────────
  sh.setRowHeight(3, 26);
  var hRange = sh.getRange(3, 1, 1, 3);
  hRange.setValues([["FIELD", "VALUE", "NOTES"]]);
  hRange.setBackground("#D49A6A");
  hRange.setFontFamily("Inter");
  hRange.setFontSize(10);
  hRange.setFontWeight("bold");
  hRange.setFontColor("#FFFFFF");
  hRange.setVerticalAlignment("middle");

  // ── Data rows ──────────────────────────────────────────────────────────────
  var rows = [
    // SECTION: HERO
    ["— HERO SECTION —",        "",    ""],
    ["current_phase",           "PDR Submitted",         "Shown in hero as current phase badge. Examples: PDR Submitted / FDR In Progress / Permits Issued"],
    ["plans_pdf_url",           "https://www.dropbox.com/scl/fi/z32u4g9axbnonhj44b81c/KELLY-CURRENT-HQ-PLANS.pdf?rlkey=zttw4dgygj52plluj640qm9dy&dl=1",
                                                          "Dropbox share link to current plans PDF. Replace when plans update."],
    ["plans_pdf_label",         "Current Plans PDF",     "Button text on hero. Examples: Current Plans PDF / 50% State Permit Set / FDR Package"],

    // SECTION: PROJECT INFO
    ["— PROJECT INFO —",        "",    ""],
    ["project_name",            "Kelly Residence",       "Project display name"],
    ["project_address",         "9095 Horned Lark Court", "Address line 1"],
    ["project_location",        "Lot 271 · Schaffer's Mill · Truckee, CA", "Address line 2"],

    // SECTION: TEAM
    ["— TEAM —",                "",    ""],
    ["architect",               "tobylongdesign / ch×tld", "Shown in portal footer or about section"],
    ["contractor",              "Timberline Construction",  "GC name"],
    ["factory",                 "Method Homes",             "Module manufacturer"],
  ];

  var dataRange = sh.getRange(4, 1, rows.length, 3);
  dataRange.setValues(rows);
  dataRange.setFontFamily("Inter");
  dataRange.setFontSize(10);
  dataRange.setFontColor("#2B2B2B");
  dataRange.setVerticalAlignment("middle");
  dataRange.setWrap(true);
  dataRange.setBackground("#FFFFFF");

  // Section header rows — grey background
  [4, 9, 13].forEach(function(r) {
    sh.getRange(r, 1, 1, 3).setBackground("#E8E8E8").setFontWeight("bold").setFontColor("#555555");
  });

  // Row heights
  for (var i = 4; i < 4 + rows.length; i++) sh.setRowHeight(i, 36);

  // ── Freeze + border ────────────────────────────────────────────────────────
  sh.setFrozenRows(3);
  sh.getRange(3, 1, 1, 3).setBorder(null, null, true, null, null, null,
    "#D49A6A", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    "CONFIG tab created as the first sheet.\n\n" +
    "Edit the VALUE column (B) to control portal content.\n" +
    "Changes sync to Firebase within 5 minutes, or run pushAllToFirebase manually."
  );
}
