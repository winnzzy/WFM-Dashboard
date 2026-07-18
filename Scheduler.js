/**
 * ==========================================================
 * SCHEDULER
 * ==========================================================
 * Generates a daily break schedule by auto-assigning
 * staggered break slots per agent.
 *
 * Business logic preserved exactly as original.
 */

/**
 * Generates a daily break schedule for all roster agents.
 * Clears existing schedule data before writing.
 */
function generateDailySchedule() {

  var ss = SpreadsheetApp.getActive();

  var roster = getSheetOrThrow(ss, SHEETS.AGENT_ROSTER);
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  // Clear existing schedule columns B–I in Daily Operations
  ops.getRange(
    OPS_DATA_START_ROW,
    2,
    MAX_OPS_ROWS,
    8
  ).clearContent();

  // Read roster data
  var rosterData = roster.getRange(
    ROSTER_DATA_START_ROW,
    1,
    MAX_ROSTER_ROWS,
    9
  ).getValues();

  // Parse agents from roster
  var agents = [];

  for (var i = 0; i < rosterData.length; i++) {

    var r = rosterData[i];

    if (!r[0]) continue; // skip blank names

    agents.push({
      name: r[0],
      center: r[1],
      supervisor: r[2],
      shift: r[5],
      queue: r[7],
      group: r[8]
    });
  }

  // Generate break slots
  var schedule = [];
  var agentIndex = 0;

  for (var g = 0; g < agents.length; g++) {

    var agent = agents[g];

    // Assign a break slot (staggered by index)
    var slotHour = 12 + (agentIndex % 3);
    var slotMinute = (agentIndex * 10) % 60;

    var breakSlot = new Date();
    breakSlot.setHours(slotHour, slotMinute, 0, 0);

    var expectedBack = new Date(breakSlot.getTime() + 60 * 60 * 1000);

    schedule.push([
      agent.name,
      agent.center,
      agent.supervisor,
      agent.shift,
      agent.queue,
      breakSlot,
      breakSlot,
      expectedBack
    ]);

    agentIndex++;
  }

  // Write schedule to Daily Operations
  if (schedule.length > 0) {

    ops.getRange(
      OPS_DATA_START_ROW,
      2,
      schedule.length,
      8
    ).setValues(schedule);

    // Format time columns G–I
    ops.getRange(
      OPS_DATA_START_ROW,
      7,
      schedule.length,
      3
    ).setNumberFormat("h:mm AM/PM");
  }
}