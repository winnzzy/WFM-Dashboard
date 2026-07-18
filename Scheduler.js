/**
 * ==========================================================
 * SCHEDULER ENGINE V5
 * ==========================================================
 */

const BREAK_SLOTS = {

  Morning: [
    { slot: "M1", start: "12:00 PM", end: "1:00 PM" },
    { slot: "M2", start: "12:15 PM", end: "1:15 PM" },
    { slot: "M3", start: "12:30 PM", end: "1:30 PM" },
    { slot: "M4", start: "12:45 PM", end: "1:45 PM" },
    { slot: "M5", start: "1:00 PM", end: "2:00 PM" }
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
 * ==========================================================
 * GENERATE DAILY SCHEDULE
 * ==========================================================
 */
function generateDailySchedule() {

  const ss = SpreadsheetApp.getActive();

  const roster = ss.getSheetByName("Agent Roster");
  const schedule = ss.getSheetByName("Daily Schedule");
  const operations = ss.getSheetByName("Daily Operations");

  if (!roster) throw new Error("Agent Roster sheet not found.");
  if (!schedule) throw new Error("Daily Schedule sheet not found.");
  if (!operations) throw new Error("Daily Operations sheet not found.");

  const rosterData = roster.getRange(2, 1, 300, 10).getValues();

  // Clear old data
  schedule.getRange("A2:J500").clearContent();
  operations.getRange("A6:R500").clearContent();

  let queuePointer = 0;

  const scheduleRows = [];
  const operationRows = [];

  // Track break slot usage
  const slotAssignments = {};

  Object.keys(BREAK_SLOTS).forEach(shift => {

    slotAssignments[shift] = {};

    BREAK_SLOTS[shift].forEach(slot => {

      slotAssignments[shift][slot.slot] = [];

    });

  });

  // Process every active agent
  rosterData.forEach(row => {

    const agent = String(row[0]).trim();

    if (!agent) return;

    const center = row[1];
    const supervisor = row[2];
    const shift = String(row[5]).trim();
    const status = String(row[6]).trim();
    const preferredQueue = String(row[7]).trim();

    if (status !== "Active") return;
    if (shift === "OFF") return;

    if (!BREAK_SLOTS[shift]) {
      throw new Error("Invalid shift: " + shift + " (" + agent + ")");
    }

    let queue;
        //==========================================================
    // QUEUE ASSIGNMENT
    //==========================================================

    if (shift === "Night") {

      // Night agents remain on Auto
      queue = "Auto";

    } else if (
      preferredQueue !== "" &&
      preferredQueue !== "Auto"
    ) {

      queue = preferredQueue;

    } else {

      queue = QUEUES[queuePointer];

      queuePointer++;

      if (queuePointer >= QUEUES.length) {
        queuePointer = 0;
      }

    }

    //==========================================================
    // BREAK SLOT ASSIGNMENT
    //==========================================================

    let selectedSlot = null;

    const availableSlots = BREAK_SLOTS[shift];

    if (shift === "Night") {

      // All Night agents share the same fixed break
      selectedSlot = availableSlots[0];

    } else {

      // Pass 1:
      // Max 2 agents and no duplicate queue
      for (let i = 0; i < availableSlots.length; i++) {

        const slot = availableSlots[i];

        const assigned =
          slotAssignments[shift][slot.slot];

        if (assigned.length >= 2) continue;

        const duplicateQueue =
          assigned.some(a => a.queue === queue);

        if (duplicateQueue) continue;

        selectedSlot = slot;

        assigned.push({
          agent: agent,
          queue: queue
        });

        break;

      }

      // Pass 2:
      // If all slots already contain the same queue,
      // relax ONLY the duplicate queue rule.
      if (!selectedSlot) {

        for (let i = 0; i < availableSlots.length; i++) {

          const slot = availableSlots[i];

          const assigned =
            slotAssignments[shift][slot.slot];

          if (assigned.length >= 2) continue;

          selectedSlot = slot;

          assigned.push({
            agent: agent,
            queue: queue
          });

          break;

        }

      }

      if (!selectedSlot) {

        throw new Error(
          "No break slot available for " + agent
        );

      }

    }

    //==========================================================
    // BUILD SCHEDULE ROW
    //==========================================================

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

    //==========================================================
    // BUILD OPERATIONS ROW
    //==========================================================

    operationRows.push([
      operationRows.length + 1,
      agent,
      center,
      supervisor,
      shift,
      queue,
      selectedSlot.slot,
      selectedSlot.start,
      selectedSlot.end,
      "",              // Login
      "",              // Actual Out
      "",              // Actual Back
      "",              // Break Used
      "",              // Variance
      "On Queue",      // Status
      "",              // Remarks
      "",              // Override
      ""               // Override Time
    ]);

  });
    //==========================================================
  // WRITE TO SHEETS (BULK WRITE)
  //==========================================================

  if (scheduleRows.length > 0) {

    schedule
      .getRange(2, 1, scheduleRows.length, 10)
      .setValues(scheduleRows);

  }

  if (operationRows.length > 0) {

 

  operations
    .getRange(6, 1, operationRows.length, 18)
    .setValues(operationRows);

}

  SpreadsheetApp.flush();

//==========================================================
// REFRESH DASHBOARD (TEMPORARILY DISABLED)
//==========================================================

/*
try {

  refreshDashboard();

} catch (err) {

  Logger.log("Dashboard refresh skipped: " + err);

}
*/

SpreadsheetApp.getUi().alert(
  scheduleRows.length +
  " agents scheduled successfully."
);

}
