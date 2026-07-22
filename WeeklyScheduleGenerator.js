/**
 * ==========================================================
 * WEEKLY SCHEDULE GENERATOR
 * ==========================================================
 * Generates a daily schedule from the Weekly Schedule sheet
 * instead of reading directly from the Agent Roster.
 *
 * Flow:
 *   1. Read today's (or tomorrow's) shift column from Weekly Schedule
 *   2. Filter out OFF / Leave / blank agents
 *   3. Look up each agent in Agent Roster for supervisor, centre, queue
 *   4. Run the existing Scheduler break-slot and queue logic
 *   5. Write to Daily Schedule and Daily Operations
 *   6. Run updateAgentStatuses() and refreshDashboard()
 */

/**
 * Generates today's schedule from the Weekly Schedule sheet.
 * Called from menu: "Generate Today's Schedule".
 */
function generateTodayFromWeekly() {
  generateFromWeeklySchedule(0);
}

/**
 * Generates tomorrow's schedule from the Weekly Schedule sheet.
 * Called from menu: "Generate Tomorrow's Schedule".
 */
function generateTomorrowFromWeekly() {
  generateFromWeeklySchedule(1);
}

/**
 * Syncs Agent Roster shift values from the Weekly Schedule for the target day.
 * Leave is mirrored as OFF in the roster shift column because the roster shift
 * validation only supports Morning, Afternoon, Night, and OFF.
 *
 * @param {number} dayOffset
 * @returns {{targetDayName: string, updatedCount: number}}
 */
function syncRosterFromWeekly(dayOffset) {

  var ss = SpreadsheetApp.getActive();

  var ws = getSheetOrThrow(ss, SHEETS.WEEKLY_SCHEDULE);
  var roster = getSheetOrThrow(ss, SHEETS.AGENT_ROSTER);

  var targetDate = new Date();
  targetDate.setDate(targetDate.getDate() + dayOffset);

  var dayNames = [
    "Sunday", "Monday", "Tuesday",
    "Wednesday", "Thursday", "Friday", "Saturday"
  ];

  var colMap = {
    "Monday": 2,
    "Tuesday": 3,
    "Wednesday": 4,
    "Thursday": 5,
    "Friday": 6,
    "Saturday": 7,
    "Sunday": 8
  };

  var targetDayName = dayNames[targetDate.getDay()];
  var targetCol = colMap[targetDayName];

  if (!targetCol) {
    throw new Error("Could not determine weekday column for: " + targetDayName);
  }

  var wsData = ws.getRange(
    WEEKLY_DATA_START_ROW,
    1,
    MAX_WEEKLY_ROWS,
    8
  ).getValues();

  var weeklyShiftMap = {};

  for (var i = 0; i < wsData.length; i++) {
    var agentName = String(wsData[i][0]).trim();
    if (!agentName) continue;

    var weeklyShift = String(wsData[i][targetCol - 1]).trim();
    if (!weeklyShift) continue;

    weeklyShiftMap[agentName] = weeklyShift === "Leave" ? "OFF" : weeklyShift;
  }

  var rosterRange = roster.getRange(
    ROSTER_DATA_START_ROW,
    1,
    MAX_ROSTER_ROWS,
    10
  );
  var rosterData = rosterRange.getValues();
  var updatedCount = 0;

  for (var r = 0; r < rosterData.length; r++) {
    var rosterAgentName = String(rosterData[r][ROSTER_COL.NAME]).trim();
    if (!rosterAgentName) continue;

    if (!Object.prototype.hasOwnProperty.call(weeklyShiftMap, rosterAgentName)) {
      continue;
    }

    var syncedShift = weeklyShiftMap[rosterAgentName];

    if (rosterData[r][ROSTER_COL.SHIFT] !== syncedShift) {
      rosterData[r][ROSTER_COL.SHIFT] = syncedShift;
      updatedCount++;
    }
  }

  rosterRange.setValues(rosterData);

  return {
    targetDayName: targetDayName,
    updatedCount: updatedCount
  };
}


/**
 * Core function that reads the Weekly Schedule, resolves roster data,
 * and runs the Scheduler to produce Daily Schedule + Daily Operations.
 *
 * @param {number} dayOffset — 0 = today, 1 = tomorrow
 */
function generateFromWeeklySchedule(dayOffset) {

  var startTime = new Date().getTime();

  syncRosterFromWeekly(dayOffset);

  var ss = SpreadsheetApp.getActive();

  var ws = getSheetOrThrow(ss, SHEETS.WEEKLY_SCHEDULE);
  var roster = getSheetOrThrow(ss, SHEETS.AGENT_ROSTER);
  var scheduleSheet = getSheetOrThrow(ss, SHEETS.DAILY_SCHEDULE);
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  // =========================================================
  // STEP 1 — Determine target weekday column
  // =========================================================
  var targetDate = new Date();
  targetDate.setDate(targetDate.getDate() + dayOffset);

  var dayIndex = targetDate.getDay();

  // JavaScript: 0=Sun, 1=Mon ... 6=Sat
  // Weekly Schedule columns: A=Agent, B=Mon(2), C=Tue(3), D=Wed(4),
  //                          E=Thu(5), F=Fri(6), G=Sat(7), H=Sun(8)
  var dayNames = [
    "Sunday", "Monday", "Tuesday",
    "Wednesday", "Thursday", "Friday", "Saturday"
  ];
  var targetDayName = dayNames[dayIndex];

  // Column mapping: Monday=2, Tuesday=3 ... Sunday=8
  var colMap = {
    "Monday": 2,
    "Tuesday": 3,
    "Wednesday": 4,
    "Thursday": 5,
    "Friday": 6,
    "Saturday": 7,
    "Sunday": 8
  };

  var targetCol = colMap[targetDayName];

  if (!targetCol) {
    throw new Error("Could not determine weekday column for: " + targetDayName);
  }

  // =========================================================
  // STEP 2 — Read Weekly Schedule data
  // =========================================================
  var wsData = ws.getRange(
    WEEKLY_DATA_START_ROW,
    1,
    MAX_WEEKLY_ROWS,
    8
  ).getValues();

  // =========================================================
  // STEP 3 — Read Agent Roster into lookup map
  // =========================================================
  var rosterData = roster.getRange(
    ROSTER_DATA_START_ROW,
    1,
    MAX_ROSTER_ROWS,
    10
  ).getValues();

  var rosterMap = {};

  for (var r = 0; r < rosterData.length; r++) {

    var rName = String(rosterData[r][ROSTER_COL.NAME]).trim();
    if (!rName) continue;

    rosterMap[rName] = {
      name: rName,
      center: rosterData[r][ROSTER_COL.CENTER],
      supervisor: rosterData[r][ROSTER_COL.SUPERVISOR],
      queue: String(rosterData[r][ROSTER_COL.QUEUE]).trim(),
      status: String(rosterData[r][ROSTER_COL.AGENT_STATUS]).trim()
    };
  }

  // =========================================================
  // STEP 4 — Build active agent list from Weekly Schedule
  // =========================================================
  var activeAgents = [];
  var missingAgents = [];

  var scheduledCount = 0;
  var offCount = 0;
  var leaveCount = 0;

  for (var i = 0; i < wsData.length; i++) {

    var agentName = String(wsData[i][0]).trim();
    if (!agentName) continue;

    var shift = String(wsData[i][targetCol - 1]).trim(); // 0-based

    // Skip OFF, Leave, blank
    if (!shift || shift === "OFF") {
      offCount++;
      continue;
    }

    if (shift === "Leave") {
      leaveCount++;
      continue;
    }

    // Validate shift value
    if (!BREAK_SLOTS[shift]) {
      throw new Error(
        "Invalid shift value '" + shift + "' for agent " + agentName +
        " on " + targetDayName + ". Expected: Morning, Afternoon, Night, OFF, or Leave."
      );
    }

    // Look up roster data
    var rosterEntry = rosterMap[agentName];

    if (!rosterEntry) {
      missingAgents.push(agentName);
      continue;
    }

    // Skip inactive agents
    if (rosterEntry.status !== "Active") {
      continue;
    }

    scheduledCount++;

    activeAgents.push({
      name: agentName,
      center: rosterEntry.center,
      supervisor: rosterEntry.supervisor,
      shift: shift,
      preferredQueue: rosterEntry.queue
    });
  }

  // =========================================================
  // ERROR HANDLING — Missing agents
  // =========================================================
  if (missingAgents.length > 0) {
    throw new Error(
      "The following agents exist in the Weekly Schedule but NOT " +
      "in the Agent Roster:\n\n" +
      missingAgents.join("\n") +
      "\n\nPlease add them to the Agent Roster before generating."
    );
  }

  // Prioritize fixed roster queues before Auto so flexible agents
  // do not block fixed-queue assignments.
  activeAgents.sort(function(a, b) {
    var aIsAuto = String(a.preferredQueue || "").trim() === "Auto";
    var bIsAuto = String(b.preferredQueue || "").trim() === "Auto";

    if (aIsAuto === bIsAuto) return 0;
    return aIsAuto ? 1 : -1;
  });

  // =========================================================
  // STEP 5 — Run Scheduler logic (same rules as generateDailySchedule)
  // =========================================================
  scheduleSheet.getRange("A2:J500").clearContent();
  ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).clearContent();

  var scheduleRows = [];
  var operationRows = [];
  var scheduledBreaks = [];

  var scheduleDate = new Date(targetDate);
  scheduleDate.setHours(0, 0, 0, 0);

  // Track break slot usage per shift
  var slotAssignments = {};

  Object.keys(BREAK_SLOTS).forEach(function(shift) {
    slotAssignments[shift] = {};
    BREAK_SLOTS[shift].forEach(function(slot) {
      slotAssignments[shift][slot.slot] = [];
    });
  });

  // Shift counts for logging
  var morningCount = 0;
  var afternoonCount = 0;
  var nightCount = 0;

  activeAgents.forEach(function(agent) {

    var shift = agent.shift;

    // Count shifts
    if (shift === "Morning") morningCount++;
    else if (shift === "Afternoon") afternoonCount++;
    else if (shift === "Night") nightCount++;

    // =========================================================
    // BREAK SLOT ASSIGNMENT (same rules as Scheduler.js)
    // =========================================================
    var availableSlots = BREAK_SLOTS[shift];
    var assignment = findBreakAssignment({
      agent: agent.name,
      shift: shift,
      preferredQueue: agent.preferredQueue,
      availableSlots: availableSlots,
      scheduleDate: scheduleDate,
      slotAssignments: slotAssignments,
      scheduledBreaks: scheduledBreaks
    });

    if (!assignment) {
      throw new Error(
        "No valid break slot and queue combination available for " + agent.name +
        " during the " + shift + " shift under the current coverage rules."
      );
    }

    var selectedSlot = assignment.slot;
    var selectedOut = assignment.scheduledOut;
    var selectedBack = assignment.scheduledBack;
    var queue = assignment.queue;

    if (!slotAssignments[shift][selectedSlot.slot]) {
      slotAssignments[shift][selectedSlot.slot] = [];
    }

    slotAssignments[shift][selectedSlot.slot].push({
      agent: agent.name,
      queue: queue
    });

    scheduledBreaks.push({
      agent: agent.name,
      queue: queue,
      shift: shift,
      slot: selectedSlot.slot,
      start: selectedOut,
      end: selectedBack
    });

    // =========================================================
    // BUILD SCHEDULE ROW
    // =========================================================
    var scheduledOut = selectedOut;
    var scheduledBack = selectedBack;

    if (!scheduledOut || !scheduledBack) {
      throw new Error(
        "Invalid break slot time for " + agent.name +
        " (" + selectedSlot.start + " - " + selectedSlot.end + ")"
      );
    }

    scheduleRows.push([
      scheduleRows.length + 1,
      agent.name,
      agent.center,
      agent.supervisor,
      shift,
      queue,
      selectedSlot.slot,
      scheduledOut,
      scheduledBack,
      "Scheduled"
    ]);

    // =========================================================
    // BUILD OPERATIONS ROW
    // =========================================================
    operationRows.push([
      new Date(scheduleDate),    // A: Date
      agent.name,                // B
      agent.center,              // C
      agent.supervisor,          // D
      shift,                     // E
      queue,                     // F
      selectedSlot.slot,         // G
      scheduledOut,              // H
      scheduledBack,             // I
      "",                        // J: Login
      "",                        // K: Actual Out
      "",                        // L: Actual Back
      "",                        // M: Break Used
      "",                        // N: Variance
      STATUS.ON_QUEUE,           // O: Status
      "",                        // P: Remarks
      "",                        // Q: Override
      ""                         // R: Override Time
    ]);
  });

  validateGeneratedBreakAssignments(scheduledBreaks);

  // =========================================================
  // STEP 6 & 7 — WRITE TO SHEETS (BATCH WRITE)
  // =========================================================
  if (scheduleRows.length > 0) {
    scheduleSheet
      .getRange(2, 1, scheduleRows.length, 10)
      .setValues(scheduleRows);
  }

  if (operationRows.length > 0) {
    ops
      .getRange(OPS_DATA_START_ROW, 1, operationRows.length, OPS_COL_COUNT)
      .setValues(operationRows);
  }

  // Regeneration clears data rows, so reapply calculated formulas.
  applyFormulas();

  SpreadsheetApp.flush();

  // =========================================================
  // STEP 8 — Run Status Engine
  // =========================================================
  updateAgentStatuses();

  // =========================================================
  // STEP 9 — Refresh Dashboard
  // =========================================================
  refreshDashboard();

  // =========================================================
  // LOGGING
  // =========================================================
  var endTime = new Date().getTime();
  var durationMs = endTime - startTime;
  var durationSec = Math.round(durationMs / 1000);

  var logSheet = ss.getSheetByName(SHEETS.DASHBOARD);

  if (logSheet) {
    var logRow = 72; // Row below SYSTEM LOG header
    logSheet.getRange(logRow, 1, 1, 8).setValues([[
      targetDayName + (dayOffset === 0 ? " (Today)" : " (Tomorrow)"),
      targetDate.toLocaleDateString(),
      Session.getEffectiveUser().getEmail(),
      scheduledCount + " scheduled (" +
        morningCount + "M, " +
        afternoonCount + "A, " +
        nightCount + "N)",
      offCount + " OFF",
      leaveCount + " Leave",
      durationSec + "s",
      "Generated at " + new Date().toLocaleTimeString()
    ]]);
  }

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    targetDayName + " Schedule Generated\n\n" +
    "Agents scheduled: " + scheduledCount + "\n" +
    "  Morning: " + morningCount + "\n" +
    "  Afternoon: " + afternoonCount + "\n" +
    "  Night: " + nightCount + "\n" +
    "OFF: " + offCount + "\n" +
    "Leave: " + leaveCount + "\n\n" +
    "Duration: " + durationSec + " seconds"
  );
}