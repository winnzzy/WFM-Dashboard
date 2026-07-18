/**
 * ==========================================================
 * SCHEDULER
 * ==========================================================
 * Generates a daily break schedule by auto-assigning
 * staggered break slots per agent.
 *
 * Business logic:
 *   - Morning shift: 5 fixed break slots (M1–M5)
 *   - Afternoon shift: 5 fixed break slots (A1–A5)
 *   - Night shift: 1 fixed break slot (N1)
 *   - Queue assignment: Round Robin (Night = Auto)
 *   - Break slots: Max 2 agents per slot
 *   - No duplicate queue inside same slot (whenever possible)
 */

/**
 * Break slot definitions per shift.
 * Times are stored as strings for sheet display.
 */
var BREAK_SLOTS = {

  Morning: [
    { slot: "M1", start: "12:00 PM", end: "1:00 PM" },
    { slot: "M2", start: "12:15 PM", end: "1:15 PM" },
    { slot: "M3", start: "12:30 PM", end: "1:30 PM" },
    { slot: "M4", start: "12:45 PM", end: "1:45 PM" },
    { slot: "M5", start: "1:00 PM",  end: "2:00 PM" }
  ],

  Afternoon: [
    { slot: "A1", start: "3:00 PM", end: "4:00 PM" },
    { slot: "A2", start: "3:15 PM", end: "4:15 PM" },
    { slot: "A3", start: "3:30 PM", end: "4:30 PM" },
    { slot: "A4", start: "3:45 PM", end: "4:45 PM" },
    { slot: "A5", start: "4:00 PM", end: "5:00 PM" }
  ],

  Night: [
    { slot: "N1", start: "3:00 AM", end: "6:00 AM" }
  ]
};


/**
 * Generates a daily break schedule for all roster agents.
 * Writes to both Daily Schedule and Daily Operations sheets.
 */
function generateDailySchedule() {

  var ss = SpreadsheetApp.getActive();

  var roster = getSheetOrThrow(ss, SHEETS.AGENT_ROSTER);
  var scheduleSheet = getSheetOrThrow(ss, SHEETS.DAILY_SCHEDULE);
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  // Read roster data
  var rosterData = roster.getRange(
    ROSTER_DATA_START_ROW,
    1,
    MAX_ROSTER_ROWS,
    9
  ).getValues();

  // Clear old data
  scheduleSheet.getRange("A2:J500").clearContent();
  ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).clearContent();

  // Round Robin queue pointer
  var queuePointer = 0;

  var scheduleRows = [];
  var operationRows = [];

  // Track break slot usage per shift
  var slotAssignments = {};

  Object.keys(BREAK_SLOTS).forEach(function(shift) {

    slotAssignments[shift] = {};

    BREAK_SLOTS[shift].forEach(function(slot) {
      slotAssignments[shift][slot.slot] = [];
    });
  });

  // Process every active agent
  rosterData.forEach(function(row) {

    var agent = String(row[ROSTER_COL.NAME]).trim();
    if (!agent) return;

    var center = row[ROSTER_COL.CENTER];
    var supervisor = row[ROSTER_COL.SUPERVISOR];
    var shift = String(row[ROSTER_COL.SHIFT]).trim();
    var agentStatus = String(row[ROSTER_COL.AGENT_STATUS]).trim();
    var preferredQueue = String(row[ROSTER_COL.QUEUE]).trim();

    // Skip inactive and off-duty agents
    if (agentStatus !== "Active") return;
    if (shift === "OFF") return;

    if (!BREAK_SLOTS[shift]) {
      throw new Error("Invalid shift: " + shift + " (" + agent + ")");
    }

    // =========================================================
    // QUEUE ASSIGNMENT
    // =========================================================
    var queue;

    if (shift === "Night") {

      // Night agents remain on Auto
      queue = "Auto";

    } else if (
      preferredQueue !== "" &&
      preferredQueue !== "Auto"
    ) {

      queue = preferredQueue;

    } else {

      // Round Robin across QUEUES
      queue = QUEUES[queuePointer];
      queuePointer++;

      if (queuePointer >= QUEUES.length) {
        queuePointer = 0;
      }
    }

    // =========================================================
    // BREAK SLOT ASSIGNMENT
    // =========================================================
    var selectedSlot = null;
    var availableSlots = BREAK_SLOTS[shift];

    if (shift === "Night") {

      // All Night agents share the same fixed break
      selectedSlot = availableSlots[0];

    } else {

      // Pass 1: Max 2 agents AND no duplicate queue
      for (var i = 0; i < availableSlots.length; i++) {

        var slot = availableSlots[i];
        var assigned = slotAssignments[shift][slot.slot];

        if (assigned.length >= 2) continue;

        var duplicateQueue =
          assigned.some(function(a) { return a.queue === queue; });

        if (duplicateQueue) continue;

        selectedSlot = slot;

        assigned.push({
          agent: agent,
          queue: queue
        });

        break;
      }

      // Pass 2: Relax ONLY the duplicate queue rule
      if (!selectedSlot) {

        for (var j = 0; j < availableSlots.length; j++) {

          var slot2 = availableSlots[j];
          var assigned2 = slotAssignments[shift][slot2.slot];

          if (assigned2.length >= 2) continue;

          selectedSlot = slot2;

          assigned2.push({
            agent: agent,
            queue: queue
          });

          break;
        }
      }

      if (!selectedSlot) {
        throw new Error("No break slot available for " + agent);
      }
    }

    // =========================================================
    // BUILD SCHEDULE ROW (Daily Schedule sheet)
    // =========================================================
    scheduleRows.push([
      scheduleRows.length + 1,
      agent,
      center,
      supervisor,
      shift,
      queue,
      selectedSlot.slot,
      selectedSlot.start,
      selectedSlot.end,
      "Scheduled"
    ]);

    // =========================================================
    // BUILD OPERATIONS ROW (Daily Operations sheet)
    // =========================================================
    operationRows.push([
      operationRows.length + 1,     // A: Date/Row number
      agent,                         // B
      center,                        // C
      supervisor,                    // D
      shift,                         // E
      queue,                         // F
      selectedSlot.slot,             // G
      selectedSlot.start,            // H
      selectedSlot.end,              // I
      "",                            // J: Login
      "",                            // K: Actual Out
      "",                            // L: Actual Back
      "",                            // M: Break Used
      "",                            // N: Variance
      STATUS.ON_QUEUE,               // O: Status
      "",                            // P: Remarks
      "",                            // Q: Override
      ""                             // R: Override Time
    ]);
  });

  // =========================================================
  // WRITE TO SHEETS (BATCH WRITE)
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

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    scheduleRows.length + " agents scheduled successfully."
  );
}