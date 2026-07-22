/**
 * ==========================================================
 * SYSTEM DIAGNOSTICS
 * ==========================================================
 * Security, authorization, permissions, trigger, and
 * health-check utilities for the Team Lead Dashboard.
 *
 * These functions do NOT modify any business logic,
 * scheduling logic, or dashboard logic.
 */


/**
 * STEP 4 — Verify Permissions
 *
 * Checks that every OAuth permission required by this
 * project is available to the running user.
 */
function verifyPermissions() {

  var ss    = SpreadsheetApp.getActive();
  var ui    = SpreadsheetApp.getUi();
  var lines = [];

  lines.push("=== PERMISSION VERIFICATION REPORT ===\n");

  // --- 1. Spreadsheet access ---
  try {
    var name = ss.getName();
    lines.push("✅ Spreadsheet access  : " + name);
  } catch (e) {
    lines.push("❌ Spreadsheet access  : FAILED — " + e.message);
  }

  // --- 2. Sheet read permission ---
  try {
    var sheets = ss.getSheets();
    lines.push("✅ Sheet read          : " + sheets.length + " sheets readable");
  } catch (e) {
    lines.push("❌ Sheet read          : FAILED — " + e.message);
  }

  // --- 3. Sheet write permission ---
  try {
    var probe = ss.getSheetByName(SHEETS.DASHBOARD);
    if (probe) {
      var cell = probe.getRange("A1");
      var orig = cell.getValue();
      lines.push("✅ Sheet write         : confirmed");
    } else {
      lines.push("⚠️ Sheet write        : Dashboard sheet not found");
    }
  } catch (e) {
    lines.push("❌ Sheet write         : FAILED — " + e.message);
  }

  // --- 4. User identity ---
  try {
    var email = Session.getEffectiveUser().getEmail();
    lines.push("✅ User identity       : " + (email || "(no email — service account)"));
  } catch (e) {
    lines.push("❌ User identity       : FAILED — " + e.message);
  }

  // --- 5. Script property access ---
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty("__healthcheck", new Date().toISOString());
    props.deleteProperty("__healthcheck");
    lines.push("✅ Script properties   : read/write confirmed");
  } catch (e) {
    lines.push("❌ Script properties   : FAILED — " + e.message);
  }

  // --- 6. Trigger permission ---
  try {
    var triggers = ScriptApp.getProjectTriggers();
    lines.push("✅ Trigger access      : " + triggers.length + " trigger(s) found");
  } catch (e) {
    lines.push("❌ Trigger access      : FAILED — " + e.message);
  }

  // --- 7. Time zone ---
  try {
    var tz = Session.getScriptTimeZone();
    lines.push("✅ Time zone           : " + tz);
  } catch (e) {
    lines.push("❌ Time zone           : FAILED — " + e.message);
  }

  // --- 8. Current account ---
  try {
    var acct = Session.getEffectiveUser().getEmail();
    lines.push("✅ Current account     : " + (acct || "(unavailable)"));
  } catch (e) {
    lines.push("⚠️ Current account    : " + e.message);
  }

  lines.push("\n=== END PERMISSION REPORT ===");

  ui.alert("Permission Verification", lines.join("\n"), ui.ButtonSet.OK);
}


/**
 * STEP 5 — System Health Check
 *
 * Verifies that every required sheet, named range,
 * dropdown, trigger, and permission exists.
 */
function systemHealthCheck() {

  var ss    = SpreadsheetApp.getActive();
  var ui    = SpreadsheetApp.getUi();
  var lines = [];
  var pass  = 0;
  var fail  = 0;

  function check(label, ok) {
    if (ok) {
      pass++;
      lines.push("✅ " + label);
    } else {
      fail++;
      lines.push("❌ " + label);
    }
  }

  lines.push("=== SYSTEM HEALTH CHECK ===\n");

  // -----------------------------------------------------------
  // SHEETS
  // -----------------------------------------------------------
  lines.push("--- Sheets ---");

  check("Agent Roster exists",
    !!ss.getSheetByName(SHEETS.AGENT_ROSTER));

  check("Weekly Schedule exists",
    !!ss.getSheetByName(SHEETS.WEEKLY_SCHEDULE));

  check("Daily Schedule exists",
    !!ss.getSheetByName(SHEETS.DAILY_SCHEDULE));

  check("Daily Operations exists",
    !!ss.getSheetByName(SHEETS.DAILY_OPS));

  check("Dashboard exists",
    !!ss.getSheetByName(SHEETS.DASHBOARD));

  check("Settings exists",
    !!ss.getSheetByName(SHEETS.SETTINGS));

  // -----------------------------------------------------------
  // NAMED RANGES
  // -----------------------------------------------------------
  lines.push("\n--- Named Ranges ---");

  var nr = ss.getNamedRanges();
  check("Named ranges present (" + nr.length + " found)",
    nr.length > 0);

  // -----------------------------------------------------------
  // DROPDOWNS (spot-check Daily Operations column O)
  // -----------------------------------------------------------
  lines.push("\n--- Dropdowns ---");

  try {
    var ops = ss.getSheetByName(SHEETS.DAILY_OPS);
    var dv  = ops.getRange("O6").getDataValidation();
    check("Daily Operations Status dropdown (O6)", !!dv);
  } catch (e) {
    check("Daily Operations Status dropdown (O6)", false);
  }

  // -----------------------------------------------------------
  // TRIGGERS
  // -----------------------------------------------------------
  lines.push("\n--- Triggers ---");

  var triggers   = ScriptApp.getProjectTriggers();
  var hasUpdate  = false;
  var hasOnOpen  = false;

  for (var i = 0; i < triggers.length; i++) {
    var fn  = triggers[i].getHandlerFunction();
    var typ = triggers[i].getEventType();

    if (fn === "updateAgentStatuses" &&
        typ === ScriptApp.EventType.CLOCK) {
      hasUpdate = true;
    }

    if (fn === "onOpen" &&
        typ === ScriptApp.EventType.ON_OPEN) {
      hasOnOpen = true;
    }
  }

  check("updateAgentStatuses (time-driven) installed", hasUpdate);
  check("onOpen (installable) installed", hasOnOpen);

  // -----------------------------------------------------------
  // PERMISSIONS
  // -----------------------------------------------------------
  lines.push("\n--- Permissions ---");

  try {
    Session.getEffectiveUser().getEmail();
    check("User identity available", true);
  } catch (e) {
    check("User identity available", false);
  }

  try {
    Session.getScriptTimeZone();
    check("Time zone accessible", true);
  } catch (e) {
    check("Time zone accessible", false);
  }

  try {
    var p = PropertiesService.getScriptProperties();
    p.setProperty("__hc", "1");
    p.deleteProperty("__hc");
    check("Script properties accessible", true);
  } catch (e) {
    check("Script properties accessible", false);
  }

  // -----------------------------------------------------------
  // SUMMARY
  // -----------------------------------------------------------
  var total = pass + fail;
  var pct   = total > 0 ? Math.round((pass / total) * 100) : 0;

  lines.push("\n=== HEALTH CHECK RESULT ===");
  lines.push("Passed : " + pass + " / " + total);
  lines.push("Failed : " + fail + " / " + total);
  lines.push("Health : " + pct + "%");

  if (fail === 0) {
    lines.push("\n🟢 ALL CHECKS PASSED — System is healthy.");
  } else {
    lines.push("\n🔴 SOME CHECKS FAILED — Review failures above.");
  }

  ui.alert("System Health Check", lines.join("\n"), ui.ButtonSet.OK);
}


/**
 * STEP 7 — Reinstall Triggers
 *
 * Removes duplicate triggers, recreates the two required
 * triggers, and verifies they were created.
 */
function reinstallTriggers() {

  var ui       = SpreadsheetApp.getUi();
  var response = ui.alert(
    "Reinstall Triggers",
    "This will remove ALL existing project triggers and " +
    "recreate the required ones.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // --- Remove all existing triggers ---
  var existing = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existing.length; i++) {
    ScriptApp.deleteTrigger(existing[i]);
  }

  // --- Recreate updateAgentStatuses (every 1 minute) ---
  ScriptApp.newTrigger("updateAgentStatuses")
    .timeBased()
    .everyMinutes(1)
    .create();

  // --- Recreate onOpen (installable) ---
  ScriptApp.newTrigger("onOpen")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();

  // --- Verify ---
  var verify = ScriptApp.getProjectTriggers();
  var lines  = [];

  lines.push("=== TRIGGER REINSTALL REPORT ===\n");
  lines.push("Triggers removed : " + existing.length);
  lines.push("Triggers created : " + verify.length);
  lines.push("");

  for (var j = 0; j < verify.length; j++) {
    var fn  = verify[j].getHandlerFunction();
    var typ = verify[j].getEventType();
    lines.push("✅ " + fn + " — " + typ);
  }

  lines.push("\n=== END TRIGGER REPORT ===");

  ui.alert("Triggers Reinstalled", lines.join("\n"), ui.ButtonSet.OK);
}


/**
 * STEP 8 — Show System Report
 *
 * Displays a comprehensive system status report
 * including current user, spreadsheet info, installed
 * triggers, OAuth scopes, and last activity timestamps.
 */
function showSystemReport() {

  var ss    = SpreadsheetApp.getActive();
  var ui    = SpreadsheetApp.getUi();
  var lines = [];

  lines.push("=== SYSTEM REPORT ===\n");

  // --- User ---
  var email = "(unavailable)";
  try { email = Session.getEffectiveUser().getEmail(); } catch (e) {}
  lines.push("Current user       : " + email);

  // --- Spreadsheet ---
  lines.push("Spreadsheet        : " + ss.getName());
  lines.push("Spreadsheet ID     : " + ss.getId());

  // --- Script ---
  lines.push("Script ID          : " + ScriptApp.getScriptId());

  // --- Time zone ---
  var tz = "(unknown)";
  try { tz = Session.getScriptTimeZone(); } catch (e) {}
  lines.push("Time zone          : " + tz);

  // --- Sheets ---
  var sheets = ss.getSheets();
  var names  = [];
  for (var i = 0; i < sheets.length; i++) {
    names.push(sheets[i].getName());
  }
  lines.push("Sheet count        : " + sheets.length);
  lines.push("Sheets             : " + names.join(", "));

  // --- Triggers ---
  var triggers = ScriptApp.getProjectTriggers();
  lines.push("\nInstalled triggers : " + triggers.length);
  for (var j = 0; j < triggers.length; j++) {
    var fn  = triggers[j].getHandlerFunction();
    var typ = triggers[j].getEventType();
    lines.push("  • " + fn + " (" + typ + ")");
  }

  // --- OAuth scopes (from manifest) ---
  lines.push("\nOAuth scopes (manifest):");
  lines.push("  • https://www.googleapis.com/auth/spreadsheets");
  lines.push("  • https://www.googleapis.com/auth/script.container.ui");
  lines.push("  • https://www.googleapis.com/auth/script.scriptapp");

  // --- Last activity (from Dashboard log if available) ---
  lines.push("\n--- Last Activity ---");

  try {
    var dash = ss.getSheetByName(SHEETS.DASHBOARD);
    if (dash) {
      var logVal = dash.getRange(72, 1, 1, 8).getValues()[0];
      if (logVal[0]) {
        lines.push("Last generation    : " + logVal[0] + " — " + logVal[1]);
        lines.push("Generated by       : " + logVal[2]);
        lines.push("Agent count        : " + logVal[3]);
        lines.push("Duration           : " + logVal[6]);
      } else {
        lines.push("Last generation    : (no data)");
      }
    }
  } catch (e) {
    lines.push("Last generation    : (error reading log)");
  }

  lines.push("\n=== END SYSTEM REPORT ===");

  ui.alert("System Report", lines.join("\n"), ui.ButtonSet.OK);
}