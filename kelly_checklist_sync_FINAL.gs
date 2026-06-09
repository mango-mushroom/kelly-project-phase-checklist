// ════════════════════════════════════════════════════════════════════
// KELLY CHECKLIST SYNC — v5  (Bidirectional: Sheet ↔ Firebase ↔ Site)
// ════════════════════════════════════════════════════════════════════
//
// SETUP (run once):
//   1. Paste this file into Apps Script (Extensions > Apps Script)
//   2. Run setupTriggers() — creates onEdit + 5-min pull
//   3. Run forceSyncChecklistsToFirebase() — seeds Firebase from sheet
//
// ONGOING (automatic):
//   • onEdit     → sheet change → Firebase (site updates in real-time)
//   • Every 5min → Firebase → sheet (site checkbox clicks propagate back)
//
// SHEET COLUMN MAP (0-indexed):
//   Col A (0) = DRC SECTION header / blank
//   Col B (1) = SHEET # 
//   Col C (2) = ✓  (TRUE/FALSE checkbox)  ← DONE_COL
//   Col D (3) = DUE DATE
//   Col E (4) = CHECKLIST ITEM label      ← LABEL_COL
//   Col F (5) = NOTES
//
// TAB → Firebase path prefix:
//   "PDR CHECKLIST"  → PDR
//   "FDR CHECKLIST"  → FDR
//   "STATE PERMIT"   → STATE
//   "LOCAL PERMIT"   → LOCAL
// ════════════════════════════════════════════════════════════════════

var FIREBASE_URL = "https://kelly-deliverables-default-rtdb.firebaseio.com";
var FIREBASE_PATH = "kelly_271_v7/checklists";

var TAB_MAP = {
  "PDR CHECKLIST":  "PDR",
  "FDR CHECKLIST":  "FDR",
  "STATE PERMIT":   "STATE",
  "LOCAL PERMIT":   "LOCAL"
};

var DONE_COL  = 2;  // Col C (0-indexed)
var LABEL_COL = 4;  // Col E (0-indexed)
var HEADER_MARKER = "DRC SECTION";  // text in Col A that marks data start

// ─── Firebase helpers ─────────────────────────────────────────────

function fbGet(path) {
  var res = UrlFetchApp.fetch(FIREBASE_URL + "/" + path + ".json", {
    method: "get", muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return null;
  return JSON.parse(res.getContentText());
}

function fbSet(path, data) {
  UrlFetchApp.fetch(FIREBASE_URL + "/" + path + ".json", {
    method: "put",
    contentType: "application/json",
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  });
}

function fbPatch(path, data) {
  UrlFetchApp.fetch(FIREBASE_URL + "/" + path + ".json", {
    method: "patch",
    contentType: "application/json",
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  });
}

// ─── Row parser — returns array of {gi, ii, done, label} ─────────

function parseTab(sheet) {
  var data = sheet.getDataRange().getValues();
  var items = [];
  var gi = -1, ii = 0;
  var dataStarted = false;

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var colA = String(row[0] || "").trim();
    var label = String(row[LABEL_COL] || "").trim();
    var done  = row[DONE_COL];

    // Wait for the header sentinel row
    if (!dataStarted) {
      if (colA === HEADER_MARKER) dataStarted = true;
      continue;
    }

    // Section header: col A has text AND label is empty
    if (colA !== "" && label === "") {
      gi++;
      ii = 0;
      continue;
    }

    // Item row: has a label
    if (label !== "" && gi >= 0) {
      items.push({
        gi: gi,
        ii: ii,
        done: (done === true || done === "TRUE" || done === "true"),
        label: label
      });
      ii++;
    }
  }
  return items;
}

// ─── PUSH: Sheet → Firebase (full sync) ──────────────────────────

function forceSyncChecklistsToFirebase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var payload = {};
  var total = 0, checked = 0;

  for (var tabName in TAB_MAP) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) { Logger.log("Sheet not found: " + tabName); continue; }
    var prefix = TAB_MAP[tabName];
    var items = parseTab(sheet);

    items.forEach(function(item) {
      var key = prefix + "_" + item.gi + "_" + item.ii;
      payload[key] = item.done;
      total++;
      if (item.done) checked++;
    });

    Logger.log(tabName + " → " + prefix + ": " + items.length + " items parsed");
  }

  fbSet(FIREBASE_PATH, payload);
  Logger.log("Sync complete: " + checked + "/" + total + " checked → Firebase");
}

// ─── PUSH: Single row → Firebase (called by onEdit) ──────────────

function pushRowToFirebase(sheet, row) {
  var tabName = sheet.getName();
  var prefix  = TAB_MAP[tabName];
  if (!prefix) return;

  var data = sheet.getDataRange().getValues();
  var gi = -1, ii = 0;
  var dataStarted = false;

  for (var r = 0; r < data.length; r++) {
    var colA  = String(data[r][0] || "").trim();
    var label = String(data[r][LABEL_COL] || "").trim();
    var done  = data[r][DONE_COL];

    if (!dataStarted) {
      if (colA === HEADER_MARKER) dataStarted = true;
      continue;
    }
    if (colA !== "" && label === "") { gi++; ii = 0; continue; }
    if (label !== "" && gi >= 0) {
      if (r === row - 1) {  // found our row (1-indexed)
        var key = prefix + "_" + gi + "_" + ii;
        var val = (done === true || done === "TRUE" || done === "true");
        var patch = {};
        patch[key] = val;
        fbPatch(FIREBASE_PATH, patch);
        Logger.log("onEdit push: " + key + " = " + val);
        return;
      }
      ii++;
    }
  }
}

// ─── PULL: Firebase → Sheet (site edits → sheet) ─────────────────

function pullChecklistsFromFirebase() {
  var fbData = fbGet(FIREBASE_PATH);
  if (!fbData) { Logger.log("No Firebase data"); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var updates = 0;

  for (var tabName in TAB_MAP) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) continue;
    var prefix = TAB_MAP[tabName];
    var data = sheet.getDataRange().getValues();
    var gi = -1, ii = 0;
    var dataStarted = false;

    for (var r = 0; r < data.length; r++) {
      var colA  = String(data[r][0] || "").trim();
      var label = String(data[r][LABEL_COL] || "").trim();

      if (!dataStarted) {
        if (colA === HEADER_MARKER) dataStarted = true;
        continue;
      }
      if (colA !== "" && label === "") { gi++; ii = 0; continue; }
      if (label !== "" && gi >= 0) {
        var key = prefix + "_" + gi + "_" + ii;
        if (fbData.hasOwnProperty(key)) {
          var fbVal = fbData[key] === true;
          var sheetVal = data[r][DONE_COL] === true;
          if (fbVal !== sheetVal) {
            sheet.getRange(r + 1, DONE_COL + 1).setValue(fbVal);
            updates++;
          }
        }
        ii++;
      }
    }
  }
  Logger.log("Pull complete: " + updates + " cells updated from Firebase");
}

// ─── TRIGGERS ────────────────────────────────────────────────────

function onEdit(e) {
  if (!e) return;
  var sheet = e.source.getActiveSheet();
  var tabName = sheet.getName();
  if (!TAB_MAP[tabName]) return;
  var col = e.range.getColumn() - 1;  // 0-indexed
  if (col !== DONE_COL) return;
  pushRowToFirebase(sheet, e.range.getRow());
}

function setupTriggers() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });

  // onEdit trigger
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  // 5-minute pull trigger
  ScriptApp.newTrigger("pullChecklistsFromFirebase")
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log("Triggers created: onEdit + 5-min pull");
}

