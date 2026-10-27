/**
 * ==========================================================
 * BREAK / SCHEDULE NOTIFICATIONS
 * ==========================================================
 * Sends each agent their own schedule, from the Agent Roster
 * Email (column E) and Phone (column D) — no code edits.
 *
 * EMAIL — works today, no setup beyond the Roster email:
 *   • Team Lead menu → "Email Today's Break Schedules":
 *     one email per scheduled agent with shift, queue,
 *     break slot and break times for that day.
 *   • "Send Break Reminders (next 15 min)": emails agents
 *     whose break starts within the next 15 minutes.
 *   • A 15-minute trigger (menu → System → Install Break
 *     Reminder Trigger) runs the reminder check hands-free.
 *   Every send is de-duplicated per agent per day per break,
 *   and logged to the "Notification Log" sheet, so a
 *   re-run never double-sends.
 *
 * WHATSAPP — two tiers, be clear about which one you have:
 *   1) MANUAL, works today, no API: after generation, the
 *      Daily Schedule gets a column K "WhatsApp" link per
 *      agent (wa.me + prefilled break message). The team
 *      lead taps it and hits send in WhatsApp. This is the
 *      default — it needs only a phone number in Roster D.
 *   2) AUTOMATIC, needs a WhatsApp Business Cloud API
 *      setup: set these Script Properties (Extensions →
 *      Apps Script → Project Settings → Script Properties):
 *        WHATSAPP_TOKEN, WHATSAPP_PHONE_NUMBER_ID, and
 *        WHATSAPP_TEMPLATE_NAME (an approved template,
 *        e.g. break_reminder, with {{1}} name, {{2}} time).
 *      Until those are set, automatic WhatsApp is skipped
 *      with a plain reason in the log — it never fails
 *      the email send. Do NOT put tokens in the sheet or
 *      in code; Script Properties only.
 *
 * TIMING the user asked for: "if they are scheduled for
 * 12, before 12 they can receive a message" — that is the
 * 15-minute reminder above (a 12:00 break gets its reminder
 * at ~11:45). The full-day email is sent when the team
 * lead runs it after generating the schedule, so agents
 * also see the whole day in advance.
 */

/**
 * Minutes before a break starts when a reminder goes out.
 */
var BREAK_REMINDER_MINUTES_BEFORE = 15;

/**
 * Formats a Date as "12:00 PM" in the spreadsheet timezone.
 */
function formatBreakTime(value) {
  if (!isValidDate(value)) return String(value || "");
  return Utilities.formatDate(value, "Africa/Lagos", "h:mm a");
}

/**
 * Formats a Date as "2026-10-27" for dedupe keys and subjects.
 */
function formatDateKey(value) {
  return Utilities.formatDate(value, "Africa/Lagos", "yyyy-MM-dd");
}

/**
 * Normalises a Roster phone to E.164 for WhatsApp.
 * Nigerian mobile "0803..." → "+234803..."; "234803..." → "+234803...".
 * Returns "" when nothing usable was entered.
 */
function normalizePhoneForWhatsApp(phone) {
  var digits = String(phone || "").replace(/[^\d+]/g, "");
  if (!digits) return "";
  if (digits.charAt(0) === "+") return digits;
  if (digits.indexOf("234") === 0 && digits.length >= 12) return "+" + digits;
  if (digits.charAt(0) === "0" && digits.length === 11) return "+234" + digits.substring(1);
  return digits;
}

/**
 * wa.me click-to-chat link with a prefilled message.
 * Works without any API account.
 */
function getWhatsAppChatLink(phone, message) {
  var e164 = normalizePhoneForWhatsApp(phone);
  if (!e164) return "";
  return "https://wa.me/" + e164.replace("+", "") +
    "?text=" + encodeURIComponent(message);
}

/**
 * The break message itself, shared by email and WhatsApp
 * so agents get identical wording on both channels.
 */
function buildBreakMessage(agentName, shift, queue, slot, outTime, backTime, dateLabel) {
  return "Hi " + agentName + ",\n\n" +
    "Your break schedule for " + dateLabel + ":\n" +
    "• Shift: " + shift + "\n" +
    "• Queue: " + queue + "\n" +
    "• Break slot: " + slot + "\n" +
    "• Break: " + formatBreakTime(outTime) + " – " + formatBreakTime(backTime) + "\n\n" +
    "Please be back on queue by " + formatBreakTime(backTime) + ".";
}

/**
 * Ensures the Notification Log sheet exists with headers.
 */
function ensureNotificationLogSheet() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.NOTIFICATION_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.NOTIFICATION_LOG);
    sh.getRange(1, 1, 1, 6).setValues([[
      "Sent At", "Agent", "Channel", "Type", "Detail", "Result"
    ]]).setBackground(COLORS.HEADER_BG)
      .setFontColor(COLORS.HEADER_TEXT)
      .setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function logNotification(agent, channel, type, detail, result) {
  try {
    var sh = ensureNotificationLogSheet();
    sh.appendRow([new Date(), agent, channel, type, detail, result]);
  } catch (e) {
    // Logging must never break a send.
  }
}

/**
 * Dedupe via Script Properties: one key per agent/day/break/type.
 */
function hasAlreadySent(key) {
  return PropertiesService.getScriptProperties().getProperty(key) === "sent";
}

function markSent(key) {
  PropertiesService.getScriptProperties().setProperty(key, "sent");
}

/**
 * Reads the Daily Schedule sheet into rows:
 * { agent, center, supervisor, shift, queue, slot, out, back }.
 * Daily Schedule columns: A # | B Agent | C Center | D Supervisor |
 * E Shift | F Queue | G Slot | H Out | I Back | J Status.
 */
function readDailyScheduleRows() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.DAILY_SCHEDULE);
  if (!sh) return [];
  var data = sh.getRange(2, 1, 499, 10).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][1] || "").trim();
    if (!name) continue;
    rows.push({
      agent: name,
      center: data[i][2],
      supervisor: data[i][3],
      shift: data[i][4],
      queue: data[i][5],
      slot: data[i][6],
      out: data[i][7],
      back: data[i][8]
    });
  }
  return rows;
}

/**
 * Adds/refreshes column K on Daily Schedule: a per-agent
 * "Send on WhatsApp" link with the break message prefilled.
 * Call after generation, or from the Team Lead menu. Safe to
 * re-run — it rewrites column K only.
 */
function addWhatsAppLinksToDailySchedule() {
  var ss = SpreadsheetApp.getActive();
  var sh = getSheetOrThrow(ss, SHEETS.DAILY_SCHEDULE);
  var contacts = getRosterMap();
  var rows = readDailyScheduleRows();

  sh.getRange(1, 11).setValue("WhatsApp")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sh.setColumnWidth(11, 200);

  var dateLabel = "";
  try {
    var ops = ss.getSheetByName(SHEETS.DAILY_OPS);
    if (ops) {
      var d = ops.getRange(OPS_DATA_START_ROW, 1, 1, 1).getValue();
      if (isValidDate(d)) dateLabel = formatDateKey(d);
    }
  } catch (e) {}
  if (!dateLabel) dateLabel = formatDateKey(new Date());

  var formulas = [];
  rows.forEach(function(r) {
    var contact = contacts[r.agent] || {};
    var link = getWhatsAppChatLink(
      contact.phone,
      buildBreakMessage(r.agent, r.shift, r.queue, r.slot, r.out, r.back, dateLabel)
    );
    formulas.push([link
      ? '=HYPERLINK("' + link + '", "Send on WhatsApp")'
      : "No phone in Roster (col D)"]);
  });

  if (formulas.length > 0) {
    sh.getRange(2, 11, formulas.length, 1).setValues(formulas);
  }
  return rows.length;
}

/**
 * Sends every scheduled agent their full-day break schedule
 * by email. Skips agents with no/invalid Roster email and
 * agents already emailed for this schedule date. Returns a
 * summary { sent, skipped, errors[] }.
 */
function sendScheduleEmailsForToday() {
  var rows = readDailyScheduleRows();
  if (rows.length === 0) {
    throw new Error("No Daily Schedule found. Generate today's schedule first.");
  }

  var contacts = getRosterMap();
  var dateLabel = formatDateKey(new Date());
  // Prefer the schedule's own date when Daily Operations has one.
  try {
    var ops = SpreadsheetApp.getActive().getSheetByName(SHEETS.DAILY_OPS);
    if (ops) {
      var d = ops.getRange(OPS_DATA_START_ROW, 1, 1, 1).getValue();
      if (isValidDate(d)) dateLabel = formatDateKey(d);
    }
  } catch (e) {}

  var sent = 0;
  var skipped = [];
  var errors = [];
  var whatsappSent = 0;

  rows.forEach(function(r) {
    var contact = contacts[r.agent] || {};
    var key = "schedule-email:" + dateLabel + ":" + r.agent;
    if (hasAlreadySent(key)) {
      skipped.push(r.agent + " (already sent today)");
      return;
    }
    if (!contact.email || !isValidEmailAddress(contact.email)) {
      skipped.push(r.agent + " (no valid email in Roster col E)");
      logNotification(r.agent, "Email", "Schedule", r.shift + " / " + r.slot, "Skipped — no valid email");
      return;
    }
    try {
      MailApp.sendEmail({
        to: contact.email,
        subject: "Your break schedule for " + dateLabel + " — " + r.shift + " shift",
        body: buildBreakMessage(r.agent, r.shift, r.queue, r.slot, r.out, r.back, dateLabel) +
          "\n\n— Team Lead Dashboard"
      });
      markSent(key);
      sent++;
      logNotification(r.agent, "Email", "Schedule", r.shift + " / " + r.slot + " " + formatBreakTime(r.out) + "–" + formatBreakTime(r.back), "Sent");
      // Automatic WhatsApp, only when the Business API is configured.
      // A missing setup never fails the email that just went out.
      var waKey = "schedule-whatsapp:" + dateLabel + ":" + r.agent;
      if (!hasAlreadySent(waKey)) {
        var wa = sendWhatsAppViaApi(contact.phone, [
          r.agent,
          formatBreakTime(r.out) + " – " + formatBreakTime(r.back),
          r.shift
        ]);
        if (wa.sent) {
          markSent(waKey);
          whatsappSent++;
          logNotification(r.agent, "WhatsApp", "Schedule", r.shift + " / " + r.slot, "Sent via Business API");
        } else if (contact.phone) {
          logNotification(r.agent, "WhatsApp", "Schedule", r.shift + " / " + r.slot, "Not auto-sent — " + wa.reason);
        }
      }
    } catch (e) {
      errors.push(r.agent + ": " + e.message);
      logNotification(r.agent, "Email", "Schedule", r.shift + " / " + r.slot, "Failed — " + e.message);
    }
  });

  return { sent: sent, skipped: skipped, errors: errors, dateLabel: dateLabel, whatsappSent: whatsappSent };
}

function sendTodayScheduleEmailsMenu() {
  var result;
  try {
    result = sendScheduleEmailsForToday();
  } catch (e) {
    SpreadsheetApp.getUi().alert("Schedule Emails", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var msg = "Schedule emails for " + result.dateLabel + ":\n\n" +
    "Sent: " + result.sent + "\n" +
    "Automatic WhatsApp sends: " + (result.whatsappSent || 0) + "\n" +
    "Skipped: " + result.skipped.length +
    (result.skipped.length ? "\n• " + result.skipped.join("\n• ") : "") + "\n" +
    "Errors: " + result.errors.length +
    (result.errors.length ? "\n• " + result.errors.join("\n• ") : "") +
    "\n\nWhatsApp: use column K on Daily Schedule (tap 'Send on WhatsApp'), or set up the WhatsApp Business API in Script Properties for automatic sends — see Notifications.js.";
  SpreadsheetApp.getUi().alert("Schedule Emails", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Automatic WhatsApp via the WhatsApp Business Cloud API.
 * Returns { sent, reason }. Never throws for a missing setup —
 * it reports it, so email sending is never blocked by this.
 */
function sendWhatsAppViaApi(phone, templateParams) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty("WHATSAPP_TOKEN");
  var phoneNumberId = props.getProperty("WHATSAPP_PHONE_NUMBER_ID");
  var templateName = props.getProperty("WHATSAPP_TEMPLATE_NAME");
  var e164 = normalizePhoneForWhatsApp(phone);

  if (!token || !phoneNumberId || !templateName) {
    return { sent: false, reason: "WhatsApp Business API not configured (WHATSAPP_TOKEN / WHATSAPP_PHONE_NUMBER_ID / WHATSAPP_TEMPLATE_NAME in Script Properties)." };
  }
  if (!e164) {
    return { sent: false, reason: "No usable phone number in Roster col D." };
  }

  var payload = {
    messaging_product: "whatsapp",
    to: e164.replace("+", ""),
    type: "template",
    template: {
      name: templateName,
      language: { code: "en_US" },
      components: [{ type: "body", parameters: templateParams.map(function(p) {
        return { type: "text", text: String(p) };
      }) }]
    }
  };

  try {
    var resp = UrlFetchApp.fetch(
      "https://graph.facebook.com/v21.0/" + phoneNumberId + "/messages",
      {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + token },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );
    if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) {
      return { sent: true, reason: "" };
    }
    return { sent: false, reason: "WhatsApp API HTTP " + resp.getResponseCode() + ": " + resp.getContentText().substring(0, 200) };
  } catch (e) {
    return { sent: false, reason: "WhatsApp API error: " + e.message };
  }
}

/**
 * Emails (and, when configured, WhatsApps) agents whose break
 * starts within the next BREAK_REMINDER_MINUTES_BEFORE minutes.
 * Safe to run every 15 minutes via trigger — dedupe prevents
 * repeat sends for the same break.
 */
function sendBreakReminders() {
  var rows = readDailyScheduleRows();
  var contacts = getRosterMap();
  var now = new Date();
  var windowMs = BREAK_REMINDER_MINUTES_BEFORE * 60000;

  var sent = 0;
  var skipped = [];
  var waAttempted = 0;

  rows.forEach(function(r) {
    if (!isValidDate(r.out)) return;
    var msUntil = r.out.getTime() - now.getTime();
    if (msUntil < 0 || msUntil > windowMs) return;

    var dateLabel = formatDateKey(r.out);
    var key = "break-reminder:" + dateLabel + ":" + r.agent + ":" + r.slot;
    if (hasAlreadySent(key)) return;

    var contact = contacts[r.agent] || {};
    var minutes = Math.round(msUntil / 60000);
    var sentThis = false;

    if (contact.email && isValidEmailAddress(contact.email)) {
      try {
        MailApp.sendEmail({
          to: contact.email,
          subject: "Break in " + minutes + " minutes — " + formatBreakTime(r.out),
          body: "Hi " + r.agent + ",\n\nYour break starts in about " + minutes +
            " minutes (" + formatBreakTime(r.out) + " – " + formatBreakTime(r.back) +
            ", " + r.shift + " shift, " + r.queue + ").\n\nPlease wrap up and be ready to go on break on time." +
            "\n\n— Team Lead Dashboard"
        });
        sentThis = true;
      } catch (e) {
        logNotification(r.agent, "Email", "Break reminder", r.slot, "Failed — " + e.message);
      }
    } else {
      skipped.push(r.agent + " (no valid email in Roster col E)");
    }

    // Automatic WhatsApp only when the Business API is configured;
    // otherwise the team lead uses the column K link.
    var wa = sendWhatsAppViaApi(contact.phone, [
      r.agent,
      formatBreakTime(r.out) + " – " + formatBreakTime(r.back),
      r.shift
    ]);
    if (wa.sent) {
      waAttempted++;
      logNotification(r.agent, "WhatsApp", "Break reminder", r.slot, "Sent via Business API");
    } else if (contact.phone) {
      logNotification(r.agent, "WhatsApp", "Break reminder", r.slot, "Not auto-sent — " + wa.reason);
    }

    if (sentThis) {
      markSent(key);
      sent++;
      logNotification(r.agent, "Email", "Break reminder", r.slot + " in " + minutes + " min", "Sent");
    }
  });

  return { sent: sent, skipped: skipped, whatsappAutoSent: waAttempted };
}

function sendBreakRemindersMenu() {
  var result = sendBreakReminders();
  SpreadsheetApp.getUi().alert(
    "Break Reminders",
    "Break reminders sent by email: " + result.sent +
    "\nAutomatic WhatsApp sends: " + result.whatsappAutoSent +
    (result.skipped.length ? "\n\nSkipped (no valid email):\n• " + result.skipped.join("\n• ") : "") +
    "\n\nFor agents without email, use the Daily Schedule column K WhatsApp link.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Installs the hands-free 15-minute notification triggers:
 * break reminders AND overdue-break alerts. Replaces any
 * previous triggers for those two functions only; leaves
 * the existing status-update triggers alone. Note: System
 * -> Reinstall Triggers deletes ALL triggers, so re-run
 * this installer after a system trigger reinstall.
 */
function installNotificationTriggers() {
  var ui = SpreadsheetApp.getUi();
  var existing = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < existing.length; i++) {
    var fn = existing[i].getHandlerFunction();
    if (fn === "sendBreakReminders" || fn === "checkOverdueBreakAlerts") {
      ScriptApp.deleteTrigger(existing[i]);
      removed++;
    }
  }
  ScriptApp.newTrigger("sendBreakReminders")
    .timeBased()
    .everyMinutes(15)
    .create();
  ScriptApp.newTrigger("checkOverdueBreakAlerts")
    .timeBased()
    .everyMinutes(15)
    .create();

  ui.alert(
    "Notification Triggers Installed",
    "sendBreakReminders and checkOverdueBreakAlerts now run every 15 minutes.\n" +
    "Previous notification triggers removed: " + removed + ".\n\n" +
    "Note: the full-day schedule email is still sent manually from " +
    "Team Lead → Email Today's Break Schedules, right after you generate " +
    "the schedule — so you control exactly when agents get the day's plan.",
    ui.ButtonSet.OK
  );
}

/**
 * ==========================================================
 * OVERDUE BREAK ALERTS
 * ==========================================================
 * The 1-minute status engine already flags Break Overdue,
 * but it tells nobody. This check tells the team lead.
 *
 * A break counts as overdue here only when it is real:
 * Actual Out is filled, Actual Back is blank, and now is
 * more than OVERDUE_ALERT_MINUTES past Scheduled Back.
 * That Actual Back check matters: the status engine alone
 * would flag everyone past their back time, returned or not.
 *
 * Alerts go to the person running the sheet (the team lead,
 * via Session.getEffectiveUser), once per agent per day per
 * break, and are logged to Notification Log. Run it from
 * the menu, or let the 15-minute notification trigger run it.
 */

/**
 * Minutes past Scheduled Back before an overdue alert fires.
 */
var OVERDUE_ALERT_MINUTES = 10;

function checkOverdueBreakAlerts() {
  var ss = SpreadsheetApp.getActive();
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);
  var data = ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).getValues();

  var leadEmail = "";
  try { leadEmail = Session.getEffectiveUser().getEmail(); } catch (e) {}
  var now = new Date();
  var alerted = [];
  var skippedNoLead = false;

  data.forEach(function(r) {
    if (!r[COL.AGENT]) return;
    if (!isValidDate(r[COL.ACTUAL_OUT])) return;
    if (isValidDate(r[COL.ACTUAL_BACK])) return;
    if (!isValidDate(r[COL.SCHEDULED_BACK])) return;

    var minutesOverdue = Math.round((now.getTime() - r[COL.SCHEDULED_BACK].getTime()) / 60000);
    if (minutesOverdue < OVERDUE_ALERT_MINUTES) return;

    var dateLabel = formatDateKey(r[COL.SCHEDULED_BACK]);
    var key = "overdue-alert:" + dateLabel + ":" + r[COL.AGENT] + ":" + r[COL.BREAK_SLOT];
    if (hasAlreadySent(key)) return;

    if (!leadEmail) {
      skippedNoLead = true;
      return;
    }

    try {
      MailApp.sendEmail({
        to: leadEmail,
        subject: "Overdue break: " + r[COL.AGENT] + " (" + minutesOverdue + " min)",
        body: "Hi,\n\n" + r[COL.AGENT] + " (" + r[COL.SHIFT] + " shift, " + r[COL.QUEUE] + ", slot " + r[COL.BREAK_SLOT] + ") " +
          "was due back at " + formatBreakTime(r[COL.SCHEDULED_BACK]) +
          " and is now " + minutesOverdue + " minutes overdue (went out " + formatBreakTime(r[COL.ACTUAL_OUT]) + ").\n\n" +
          "— Team Lead Dashboard"
      });
      markSent(key);
      alerted.push(r[COL.AGENT] + " (" + minutesOverdue + " min)");
      logNotification(r[COL.AGENT], "Email", "Overdue alert", r[COL.BREAK_SLOT] + " " + minutesOverdue + " min overdue", "Sent to team lead");
    } catch (e) {
      logNotification(r[COL.AGENT], "Email", "Overdue alert", r[COL.BREAK_SLOT], "Failed — " + e.message);
    }
  });

  return { alerted: alerted, skippedNoLead: skippedNoLead };
}

function checkOverdueBreakAlertsMenu() {
  var result = checkOverdueBreakAlerts();
  var msg = result.alerted.length > 0
    ? "Overdue alerts sent to the team lead:\n• " + result.alerted.join("\n• ")
    : "No overdue breaks needing an alert right now (overdue = back time + " + OVERDUE_ALERT_MINUTES + " min, still out).";
  if (result.skippedNoLead) {
    msg += "\n\nNote: no team-lead email address was available, so alerts were not sent.";
  }
  SpreadsheetApp.getUi().alert("Overdue Breaks", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * WhatsApp setup check — reports what is present and what
 * is missing, without ever showing a token value. Run from
 * Team Lead -> Check WhatsApp Auto-Send Setup.
 *
 * To enable automatic WhatsApp: in the Apps Script editor,
 * Project Settings -> Script Properties, set:
 *   WHATSAPP_TOKEN, WHATSAPP_PHONE_NUMBER_ID,
 *   WHATSAPP_TEMPLATE_NAME
 * The template must be an approved WhatsApp template with
 * body parameters in this order: {{1}} agent name,
 * {{2}} break time range, {{3}} shift. Until all three
 * are set, sends fall back to the manual column K links.
 */
function checkWhatsAppSetupMenu() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty("WHATSAPP_TOKEN");
  var phoneId = props.getProperty("WHATSAPP_PHONE_NUMBER_ID");
  var template = props.getProperty("WHATSAPP_TEMPLATE_NAME");

  var lines = [
    "WhatsApp automatic sends: " + ((token && phoneId && template) ? "READY" : "NOT CONFIGURED"),
    "",
    "WHATSAPP_TOKEN: " + (token ? "set (value not shown)" : "MISSING"),
    "WHATSAPP_PHONE_NUMBER_ID: " + (phoneId ? "set (value not shown)" : "MISSING"),
    "WHATSAPP_TEMPLATE_NAME: " + (template ? template : "MISSING"),
    "",
    "Template body parameters must be, in order:",
    "{{1}} agent name, {{2}} break time range, {{3}} shift.",
    "",
    "Until all three are set, use the Daily Schedule",
    "column K 'Send on WhatsApp' links — no API needed."
  ];
  SpreadsheetApp.getUi().alert("WhatsApp Setup", lines.join("\n"), SpreadsheetApp.getUi().ButtonSet.OK);
}
