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
