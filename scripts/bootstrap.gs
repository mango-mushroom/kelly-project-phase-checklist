// ════════════════════════════════════════════════════════════════════════════
// KELLY TIER 2 BOOTSTRAP v2 — self-updating via Apps Script API
// This file lives alongside kelly_all_sync.gs in the same project.
// ════════════════════════════════════════════════════════════════════════════

var GITHUB_RAW_SYNC = "https://raw.githubusercontent.com/mango-mushroom/kelly-project-phase-checklist/main/scripts/kelly_all_sync.gs";

// ── STEP 1 ────────────────────────────────────────────────────────────────────
// Run getScriptId(). The ID appears instantly in the Execution Log below.
// Copy it — you'll need it to find your Cloud project in Step 2.

function getScriptId() {
  var id = ScriptApp.getScriptId();
  Logger.log("=================================");
  Logger.log("SCRIPT ID: " + id);
  Logger.log("=================================");
  Logger.log("Copy the ID above, then follow Step 2 in the comments.");
  // This function finishes in under 1 second.
  // Check the Execution Log panel below for the ID.
}

// ── STEP 2 ────────────────────────────────────────────────────────────────────
// Enable Apps Script API + add OAuth scope. Do this ONCE:
//
//  A. Go to: console.cloud.google.com
//     Select the project linked to your script:
//     (Apps Script editor → Project Settings → Google Cloud Platform project → link)
//
//  B. Click "APIs & Services" → "+ Enable APIs and Services"
//     Search: "Apps Script API" → Enable it
//
//  C. Back in Apps Script editor:
//     Project Settings → check "Show appsscript.json manifest file in editor"
//
//  D. Click appsscript.json in the Files list, replace everything with:
//
//     {
//       "timeZone": "America/Los_Angeles",
//       "dependencies": {},
//       "exceptionLogging": "STACKDRIVER",
//       "runtimeVersion": "V8",
//       "oauthScopes": [
//         "https://www.googleapis.com/auth/spreadsheets",
//         "https://www.googleapis.com/auth/drive",
//         "https://www.googleapis.com/auth/script.external_request",
//         "https://www.googleapis.com/auth/script.projects"
//       ]
//     }
//
//  E. Save (Cmd+S). The new scope will be requested on next Run.

// ── STEP 3 ────────────────────────────────────────────────────────────────────
// Run updateScriptViaAPI(). When prompted, click "Review Permissions" and Allow.
// The function completes in a few seconds and logs SUCCESS or an error message.
// Then close + reopen the Apps Script editor to see the updated kelly_all_sync.gs.

function updateScriptViaAPI() {
  var scriptId = ScriptApp.getScriptId();
  Logger.log("Script ID: " + scriptId);

  // 1. Fetch latest sync script from GitHub
  Logger.log("Fetching from GitHub...");
  var ghResp = UrlFetchApp.fetch(GITHUB_RAW_SYNC, { muteHttpExceptions: true });
  var httpCode = ghResp.getResponseCode();
  Logger.log("GitHub response: " + httpCode);

  if (httpCode !== 200) {
    Logger.log("ERROR: GitHub fetch failed with HTTP " + httpCode);
    return;
  }

  var newCode = ghResp.getContentText();
  Logger.log("Fetched " + newCode.split('\n').length + " lines from GitHub.");

  // 2. Push to Apps Script API
  Logger.log("Calling Apps Script API...");
  var token   = ScriptApp.getOAuthToken();
  var apiUrl  = "https://script.googleapis.com/v1/projects/" + scriptId + "/content";

  var apiResp = UrlFetchApp.fetch(apiUrl, {
    method: "put",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + token },
    payload: JSON.stringify({
      files: [
        { name: "kelly_all_sync", type: "SERVER_JS", source: newCode }
      ]
    }),
    muteHttpExceptions: true
  });

  var code    = apiResp.getResponseCode();
  var body    = apiResp.getContentText();
  Logger.log("Apps Script API response: " + code);

  if (code === 200) {
    Logger.log("SUCCESS — kelly_all_sync.gs updated from GitHub.");
    Logger.log("Close and reopen the Apps Script editor to see the new code.");
    Logger.log("Running pushAllToFirebase now from kelly_all_sync.gs...");
    // Note: the new code is in the project but this execution still uses old memory.
    // To run pushAllToFirebase from the new code: close editor, reopen, run it.

  } else if (code === 401 || code === 403) {
    Logger.log("SCOPE ERROR (HTTP " + code + ") — Apps Script API scope not yet granted.");
    Logger.log("Complete Step 2: add script.projects to appsscript.json oauthScopes.");
    Logger.log("Full error: " + body.slice(0, 400));

  } else {
    Logger.log("API ERROR HTTP " + code + ": " + body.slice(0, 400));
  }
}
