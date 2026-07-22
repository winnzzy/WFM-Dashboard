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
 *   - Queue assignment: Uses Agent Roster queue (or all queues only when roster value is Auto)
 *   - Break slots: Max 2 agents per slot
 *   - No overlapping breaks within the same queue
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
 * Legacy entrypoint kept for compatibility.
 * Daily schedule generation now always uses Weekly Schedule as the source.
 */
function generateDailySchedule() {
  generateFromWeeklySchedule(0);
}