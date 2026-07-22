/**
 * ==========================================================
 * WFM DASHBOARD
 * Utilities.gs — Shared constants and helper functions
 * ==========================================================
 */

/**
 * Daily Operations column indexes (0-based for array access).
 * Maps to columns A–R in the "Daily Operations" sheet.
 */
const COL = {
  DATE: 0,              // A
  AGENT: 1,             // B
  TEAM: 2,              // C
  SUPERVISOR: 3,        // D
  SHIFT: 4,             // E
  QUEUE: 5,             // F
  BREAK_SLOT: 6,        // G
  SCHEDULED_OUT: 7,     // H
  SCHEDULED_BACK: 8,    // I
  LOGIN: 9,             // J
  ACTUAL_OUT: 10,       // K
  ACTUAL_BACK: 11,      // L
  BREAK_USED: 12,       // M
  VARIANCE: 13,         // N
  STATUS: 14,           // O
  REMARKS: 15,          // P
  OVERRIDE: 16,         // Q
  OVERRIDE_TIME: 17     // R
};

/**
 * Total number of columns in the Daily Operations sheet.
 */
const OPS_COL_COUNT = 18;

/**
 * Roster sheet column indexes (0-based for array access).
 * Maps to columns A–I in the "Agent Roster" sheet.
 */
const ROSTER_COL = {
  NAME: 0,          // A
  CENTER: 1,        // B
  SUPERVISOR: 2,    // C
  /* 3 = D (unused) */
  /* 4 = E (unused) */
  SHIFT: 5,         // F
  AGENT_STATUS: 6,  // G (Active/Inactive)
  QUEUE: 7,         // H
  GROUP: 8           // I
};

/**
 * Queue names — single source of truth.
 */
const QUEUES = [
  "Call",
  "Email",
  "Clara",
  "Ebanqo"
];

/**
 * Normalizes and validates the roster queue for scheduling.
 * @param {string} rosterQueue
 * @param {string} agent
 * @returns {string}
 */
function getAssignedQueueFromRoster(rosterQueue, agent) {
  var normalizedQueue = String(rosterQueue || "").trim();
  var allowedQueues = QUEUES.concat(["Auto"]);

  if (!normalizedQueue) {
    throw new Error(
      "Missing roster queue for " + agent +
      ". Set Queue Preference in Agent Roster before scheduling."
    );
  }

  if (allowedQueues.indexOf(normalizedQueue) === -1) {
    throw new Error(
      "Invalid roster queue '" + normalizedQueue + "' for " + agent +
      ". Allowed values: " + allowedQueues.join(", ") + "."
    );
  }

  return normalizedQueue;
}

/**
 * Returns queue candidates based on roster configuration.
 * Fixed queue values must stay on that queue.
 * "Auto" means try all queues in default order.
 *
 * @param {string} rosterQueue
 * @param {string} agent
 * @returns {Array<string>}
 */
function getQueueCandidates(rosterQueue, agent) {
  var assignedQueue = getAssignedQueueFromRoster(rosterQueue, agent);

  if (assignedQueue === "Auto") {
    return QUEUES.slice();
  }

  return [assignedQueue];
}

/**
 * Counts how many scheduled breaks already exist for a queue in a shift.
 * @param {Array<Object>} assignments
 * @param {string} shift
 * @param {string} queue
 * @returns {number}
 */
function getShiftQueueLoad(assignments, shift, queue) {
  var queueKey = normalizeQueueName(queue);

  return assignments.filter(function(assignment) {
    return assignment.shift === shift &&
      normalizeQueueName(assignment.queue) === queueKey;
  }).length;
}

/**
 * Selects the first valid slot for an agent.
 * Queue comes from roster (or all queues if roster value is "Auto").
 * @param {Object} options
 * @returns {{slot: Object, scheduledOut: Date, scheduledBack: Date, queue: string}|null}
 */
function findBreakAssignment(options) {
  var queueCandidates = getQueueCandidates(
    options.preferredQueue,
    options.agent
  );

  var orderedQueueCandidates = queueCandidates;

  // Auto can use any queue; prefer least-loaded queue for this shift.
  if (queueCandidates.length > 1) {
    orderedQueueCandidates = queueCandidates.slice().sort(function(a, b) {
      return getShiftQueueLoad(options.scheduledBreaks, options.shift, a) -
        getShiftQueueLoad(options.scheduledBreaks, options.shift, b);
    });
  }

  for (var i = 0; i < options.availableSlots.length; i++) {
    var slot = options.availableSlots[i];
    var candidateOut = buildDateTime(options.scheduleDate, slot.start);
    var candidateBack = buildDateTime(options.scheduleDate, slot.end);

    if (!candidateOut || !candidateBack) {
      throw new Error(
        "Invalid break slot time for " + options.agent +
        " (" + slot.start + " - " + slot.end + ")"
      );
    }

    for (var q = 0; q < orderedQueueCandidates.length; q++) {
      var candidateQueue = orderedQueueCandidates[q];

      if (shouldEnforceQueueCoverage(options.shift)) {
        if (hasQueueCoverageConflict(
          options.scheduledBreaks,
          candidateQueue,
          candidateOut,
          candidateBack
        )) {
          continue;
        }
      }

      return {
        slot: slot,
        scheduledOut: candidateOut,
        scheduledBack: candidateBack,
        queue: candidateQueue
      };
    }
  }

  // Daytime overflow fallback: create the next back-to-back slot
  // for the same queue so no agent is left without a break.
  if (shouldEnforceQueueCoverage(options.shift) && orderedQueueCandidates.length > 0) {
    var fallbackQueue = orderedQueueCandidates[0];

    var sameQueueBreaks = options.scheduledBreaks.filter(function(assignment) {
      return assignment.shift === options.shift &&
        normalizeQueueName(assignment.queue) === normalizeQueueName(fallbackQueue);
    });

    if (sameQueueBreaks.length > 0 && options.availableSlots.length > 0) {
      var firstSlot = options.availableSlots[0];
      var firstOut = buildDateTime(options.scheduleDate, firstSlot.start);
      var firstBack = buildDateTime(options.scheduleDate, firstSlot.end);

      if (firstOut && firstBack) {
        var durationMs = firstBack.getTime() - firstOut.getTime();
        var latestBack = sameQueueBreaks.reduce(function(maxBack, assignment) {
          return assignment.end.getTime() > maxBack.getTime() ? assignment.end : maxBack;
        }, sameQueueBreaks[0].end);

        var overflowOut = new Date(latestBack.getTime());
        var overflowBack = new Date(latestBack.getTime() + durationMs);

        var prefix = String(firstSlot.slot || options.shift.charAt(0)).charAt(0);
        var maxSlotIndex = options.availableSlots.length;

        options.scheduledBreaks.forEach(function(assignment) {
          if (assignment.shift !== options.shift) return;
          var slotText = String(assignment.slot || "").trim();
          if (slotText.charAt(0) !== prefix) return;

          var idx = parseInt(slotText.substring(1), 10);
          if (!isNaN(idx) && idx > maxSlotIndex) {
            maxSlotIndex = idx;
          }
        });

        return {
          slot: {
            slot: prefix + String(maxSlotIndex + 1),
            start: "",
            end: ""
          },
          scheduledOut: overflowOut,
          scheduledBack: overflowBack,
          queue: fallbackQueue
        };
      }
    }
  }

  return null;
}

/**
 * Agent status values — single source of truth.
 */
const STATUS = {
  ON_QUEUE: "On Queue",
  ON_BREAK: "On Break",
  BREAK_OVERDUE: "Break Overdue",
  ABSENT: "Absent",
  LOGGED_OUT: "Logged Out"
};

/**
 * Dashboard section ranges (on the "Dashboard" sheet).
 */
const DASHBOARD = {
  BREAK_MONITOR: "A24:H200",
  QUEUE_SHARE: "J14:N19",
  LATE_RETURNS: "A45:H70"
};

/**
 * Shared color constants.
 */
const COLORS = {
  HEADER_BG: "#0F4C81",
  HEADER_TEXT: "#FFFFFF",
  ALT_ROW: "#F8F9FA",
  BODY_BG: "#FFFFFF",
  SECTION_HEADER_BG: "#D9EAD3",
  LATE_HEADER_BG: "#C62828",
  LATE_HEADER_SUB: "#F4CCCC",
  OVERRIDE_HEADER_BG: "#F57C00",
  OVERRIDE_HEADER_SUB: "#FCE5CD"
};

/**
 * Sheet name constants — single source of truth for sheet references.
 */
const SHEETS = {
  DASHBOARD: "Dashboard",
  DAILY_OPS: "Daily Operations",
  AGENT_ROSTER: "Agent Roster",
  DAILY_SCHEDULE: "Daily Schedule",
  WEEKLY_SCHEDULE: "Weekly Schedule",
  SETTINGS: "Settings"
};

/**
 * Maximum number of agent rows in Daily Operations.
 */
const MAX_OPS_ROWS = 500;

/**
 * Data starts at this row in Daily Operations (1-based).
 */
const OPS_DATA_START_ROW = 6;

/**
 * Data starts at this row in Agent Roster (1-based).
 */
const ROSTER_DATA_START_ROW = 2;

/**
 * Maximum roster rows to scan.
 */
const MAX_ROSTER_ROWS = 300;

/**
 * Maximum rows in the Weekly Schedule sheet.
 */
const MAX_WEEKLY_ROWS = 50;

/**
 * Data starts at this row in Weekly Schedule (1-based).
 */
const WEEKLY_DATA_START_ROW = 2;

/**
 * Allowed shift values for Weekly Schedule dropdown validation.
 */
const WEEKLY_SHIFTS = ["Morning", "Afternoon", "Night", "OFF", "Leave"];

/**
 * Converts a Date to minutes since midnight.
 * @param {Date} value
 * @returns {number|null}
 */
function toMinutes(value) {
  if (!(value instanceof Date)) return null;
  return value.getHours() * 60 + value.getMinutes();
}

/**
 * Checks if a value is a valid Date.
 * @param {*} value
 * @returns {boolean}
 */
function isValidDate(value) {
  return value instanceof Date && !isNaN(value);
}

/**
 * Builds a full DateTime from a base date and 12-hour time text.
 * Example: baseDate=2026-07-20, timeText="4:15 PM" => 2026-07-20 16:15
 *
 * @param {Date} baseDate
 * @param {string} timeText
 * @returns {Date|null}
 */
function buildDateTime(baseDate, timeText) {
  if (!isValidDate(baseDate) || !timeText) return null;

  var match = String(timeText).trim().match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (!match) return null;

  var hour = parseInt(match[1], 10);
  var minute = parseInt(match[2], 10);
  var period = match[3].toUpperCase();

  if (hour < 1 || hour > 12 || minute < 0 || minute > 59) return null;

  if (period === "AM") {
    if (hour === 12) hour = 0;
  } else {
    if (hour !== 12) hour += 12;
  }

  var dt = new Date(baseDate);
  dt.setHours(hour, minute, 0, 0);
  return dt;
}

/**
 * Normalizes queue names for safe equality checks.
 * @param {*} queue
 * @returns {string}
 */
function normalizeQueueName(queue) {
  return String(queue || "").trim().toLowerCase();
}

/**
 * Queue overlap enforcement applies to Morning/Afternoon only.
 * Night remains unchanged.
 * @param {string} shift
 * @returns {boolean}
 */
function shouldEnforceQueueCoverage(shift) {
  return shift !== "Night";
}

/**
 * Returns true when two break intervals overlap.
 * @param {Date} startA
 * @param {Date} endA
 * @param {Date} startB
 * @param {Date} endB
 * @returns {boolean}
 */
function breaksOverlap(startA, endA, startB, endB) {
  if (!isValidDate(startA) || !isValidDate(endA)) return false;
  if (!isValidDate(startB) || !isValidDate(endB)) return false;

  return startA.getTime() < endB.getTime() &&
    startB.getTime() < endA.getTime();
}

/**
 * Returns true when a queue already has an overlapping break.
 * @param {Array<Object>} assignments
 * @param {string} queue
 * @param {Date} start
 * @param {Date} end
 * @returns {boolean}
 */
function hasQueueCoverageConflict(assignments, queue, start, end) {
  var queueKey = normalizeQueueName(queue);

  return assignments.some(function(assignment) {
    return normalizeQueueName(assignment.queue) === queueKey &&
      breaksOverlap(assignment.start, assignment.end, start, end);
  });
}

/**
 * Validates break capacity and queue coverage on a generated schedule.
 * @param {Array<Object>} assignments
 */
function validateGeneratedBreakAssignments(assignments) {
  for (var i = 0; i < assignments.length; i++) {
    var current = assignments[i];
    var currentQueueKey = normalizeQueueName(current.queue);

    for (var j = i + 1; j < assignments.length; j++) {
      var other = assignments[j];

      if (!shouldEnforceQueueCoverage(current.shift) ||
          !shouldEnforceQueueCoverage(other.shift)) {
        continue;
      }

      if (currentQueueKey !== normalizeQueueName(other.queue)) continue;
      if (!breaksOverlap(current.start, current.end, other.start, other.end)) {
        continue;
      }

      throw new Error(
        "Queue coverage violated for queue '" + current.queue +
        "' between " + current.agent + " and " + other.agent + "."
      );
    }
  }
}

/**
 * Returns a formatted status string with emoji for the dashboard.
 * @param {string} status
 * @returns {string}
 */
function formatStatus(status) {
  switch (status) {
    case STATUS.ON_QUEUE:
      return "🟢 On Queue";
    case STATUS.ON_BREAK:
      return "🟡 On Break";
    case STATUS.BREAK_OVERDUE:
      return "🔴 Break Overdue";
    case STATUS.ABSENT:
      return "⚫ Absent";
    case STATUS.LOGGED_OUT:
      return "⚪ Logged Out";
    default:
      return status;
  }
}

/**
 * Returns a sheet by name, or throws if not found.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {string} name — use SHEETS constant
 * @returns {SpreadsheetApp.Sheet}
 */
function getSheetOrThrow(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet "' + name + '" not found.');
  }
  return sheet;
}