/**
 * ==========================================================
 * AGENT ROSTER — SHEET-DRIVEN
 * ==========================================================
 * The Agent Roster sheet is the source of truth for agents.
 * Add an agent by typing a new row in the sheet. Remove one
 * by deleting its row OR by setting Employment Status to
 * "Inactive" — no code change is needed either way, and
 * generation now skips (rather than crashes on) agents that
 * are no longer on the Roster.
 *
 * buildRoster() is deliberately NON-DESTRUCTIVE: it only
 * creates the sheet/headers/validation if they are missing,
 * and it only seeds the current 14 agents when the sheet is
 * completely empty. Re-running INSTALL must never wipe
 * agents, emails or phones you typed into the sheet.
 * (The old version cleared the whole sheet and rewrote a
 * hardcoded list on every build — that is what forced
 * add/remove to be done in code.)
 */

/**
 * Roster headers, in sheet order (A–J).
 */
var ROSTER_HEADERS = [
  "Agent Name",
  "Center",
  "Supervisor",
  "Phone",
  "Email",
  "Shift",
  "Employment Status",
  "Queue Preference",
  "Break Group",
  "Remarks"
];

/**
 * The seed list, used ONLY when the Roster sheet is empty
 * (first install). After that, the sheet itself is the list.
 */
var ROSTER_SEED_AGENTS = [
  ["Yimaumuaju Joshua", "", "", "", "", "Morning", "Active", "Auto", "", ""],
  ["Shehu Usman Daniya", "", "", "", "", "Afternoon", "Active", "Auto", "", ""],
  ["Mariam Ajikeola Abdulazeez", "", "", "", "", "Afternoon", "Active", "Auto", "", ""],
  ["Comfort Ijeoma Okereke", "", "", "", "", "Night", "Active", "Auto", "", ""],
  ["Thomas Goodeness Dilah", "", "", "", "", "Morning", "Active", "Auto", "", ""],
  ["Subuhanallahi Abiodun Adeleye", "", "", "", "", "Morning", "Active", "Auto", "", ""],
  ["Emmanuel Okahena Owobu", "", "", "", "", "Night", "Active", "Auto", "", ""],
  ["Zainab Yusuf", "", "", "", "", "Afternoon", "Active", "Auto", "", ""],
  ["Cynthia Nkechi Nwadukwe", "", "", "", "", "Morning", "Active", "Auto", "", ""],
  ["Daniel Zongo", "", "", "", "", "OFF", "Active", "Auto", "", ""],
  ["Nelson Okeah", "", "", "", "", "Morning", "Active", "Auto", "", ""],
  ["Isaac Idemeto", "", "", "", "", "Morning", "Active", "Auto", "", ""],
  ["Michael Asapuo", "", "", "", "", "Morning", "Active", "Auto", "", ""],
  ["Franklin Anyanwu", "", "", "", "", "Morning", "Active", "Auto", "", ""]
];

/**
 * Creates the Agent Roster sheet if missing, ensures its
 * headers/widths/validation, and seeds it ONLY if empty.
 * Existing rows are never touched.
 */
function buildRoster() {

  var ss = SpreadsheetApp.getActive();

  var sh = ss.getSheetByName(SHEETS.AGENT_ROSTER);

  if (!sh) {
    sh = ss.insertSheet(SHEETS.AGENT_ROSTER);
  }

  // Headers (write only if row 1 is not already the roster header).
  var firstRow = sh.getRange(1, 1, 1, ROSTER_HEADERS.length).getValues()[0];
  if (String(firstRow[0] || "").trim() !== "Agent Name") {
    sh.getRange(1, 1, 1, ROSTER_HEADERS.length)
      .setValues([ROSTER_HEADERS])
      .setBackground(COLORS.HEADER_BG)
      .setFontColor(COLORS.HEADER_TEXT)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
  }

  // Seed only when there are no agent rows at all.
  var existing = getRosterAgents();
  if (existing.length === 0) {
    sh.getRange(2, 1, ROSTER_SEED_AGENTS.length, ROSTER_HEADERS.length)
      .setValues(ROSTER_SEED_AGENTS);
  }

  // Column widths.
  var widths = [250, 100, 180, 140, 220, 100, 140, 140, 120, 250];
  widths.forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });

  sh.setFrozenRows(1);
}

/**
 * Reads every agent row from the Agent Roster sheet.
 * Blank-name rows are ignored. No filtering on status here —
 * callers decide whether to include Inactive agents.
 *
 * @returns {Array<Object>} each with name, center, supervisor,
 *   phone, email, shift, status, queue, group, remarks, row
 */
function getRosterAgents() {

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.AGENT_ROSTER);
  if (!sh) return [];

  var data = sh.getRange(
    ROSTER_DATA_START_ROW,
    1,
    MAX_ROSTER_ROWS,
    ROSTER_HEADERS.length
  ).getValues();

  var agents = [];
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][ROSTER_COL.NAME] || "").trim();
    if (!name) continue;
    agents.push({
      name: name,
      center: data[i][ROSTER_COL.CENTER],
      supervisor: data[i][ROSTER_COL.SUPERVISOR],
      phone: String(data[i][ROSTER_COL.PHONE] || "").trim(),
      email: String(data[i][ROSTER_COL.EMAIL] || "").trim(),
      shift: String(data[i][ROSTER_COL.SHIFT] || "").trim(),
      status: String(data[i][ROSTER_COL.AGENT_STATUS] || "").trim(),
      queue: String(data[i][ROSTER_COL.QUEUE] || "").trim(),
      group: data[i][ROSTER_COL.GROUP],
      remarks: data[i][ROSTER_COL.REMARKS],
      row: ROSTER_DATA_START_ROW + i
    });
  }
  return agents;
}

/**
 * Active agents only — the set that can be scheduled and notified.
 * An agent is active when Employment Status is blank (legacy rows)
 * or "Active". "Inactive" and "Leave" are excluded.
 */
function getActiveRosterAgents() {
  return getRosterAgents().filter(function(a) {
    return a.status === "" || a.status === "Active";
  });
}

/**
 * Name -> agent map for the whole roster (any status).
 */
function getRosterMap() {
  var map = {};
  getRosterAgents().forEach(function(a) {
    map[a.name] = a;
  });
  return map;
}

/**
 * Simple email sanity check (format only, not deliverability).
 */
function isValidEmailAddress(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}

/**
 * Finds roster problems before they break a schedule or a send:
 * duplicate names, active agents with no shift/queue, and
 * malformed emails. Returns an array of human-readable lines
 * (empty = clean).
 */
function validateRoster() {
  var agents = getRosterAgents();
  var problems = [];
  var seen = {};

  agents.forEach(function(a) {
    if (seen[a.name]) {
      problems.push("Duplicate roster name: " + a.name + " (rows " + seen[a.name] + " and " + a.row + ").");
    } else {
      seen[a.name] = a.row;
    }
    if (a.status === "Active") {
      if (!a.shift) problems.push(a.name + ": no Shift set.");
      if (!a.queue) problems.push(a.name + ": no Queue Preference set.");
    }
    if (a.email && !isValidEmailAddress(a.email)) {
      problems.push(a.name + ": email looks invalid (" + a.email + "). Fix it in column E before sending notifications.");
    }
  });

  return problems;
}
