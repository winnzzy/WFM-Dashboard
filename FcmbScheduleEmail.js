/**
 * ==========================================================
 * FCMB SCHEDULE EMAIL — PREVIEW FIRST, NEVER AUTO-SEND
 * ==========================================================
 *
 * Replaces the manual workflow of generating the Daily
 * Schedule, exporting it, copying it to a new sheet and
 * sharing that sheet: this emails the one-day Daily
 * Schedule to the FCMB parties instead.
 *
 * The preview gate is the whole point. Winner asked for it
 * explicitly (2026-10-27): he must see the email first.
 *
 *   1) Team Lead -> "Preview Schedule Email to FCMB"
 *      builds a sheet "FCMB Schedule Email - PREVIEW"
 *      with the recipients, subject and the full schedule
 *      table. NOTHING is sent by this step.
 *   2) Winner reads that sheet. If it is right, he runs
 *      Team Lead -> "Send Previewed Schedule Email".
 *      That step asks for a Yes/No confirmation, sends
 *      exactly the previewed table, then marks the preview
 *      sheet SENT so the same preview can never be sent
 *      twice. Regenerating a preview creates a fresh,
 *      unsent preview.
 *
 * Recipients were supplied by Winner on 2026-10-27.
 * They are client contacts, not agents, so they never
 * come from the Agent Roster. Do not add recipients
 * without Winner naming them.
 *
 * Format: the schedule goes in the email body as a
 * structured HTML table, with a plain-text list as the
 * fallback. No attachment, so no Drive scope is needed
 * and the recipients need no access to the dashboard
 * spreadsheet. If Winner later wants a PDF attachment
 * instead, that is a separate change (Drive scope +
 * re-authorisation).
 */

var FCMB_SCHEDULE_RECIPIENTS = [
  { name: "Keziah Tsuwang", email: "Keziah.Tsuwang@fcmb.com" },
  { name: "Oladapo Koleowo", email: "Oladapo.Koleowo@fcmb.com" },
  { name: "Sarah Obule", email: "Sarah.Obule@fcmb.com" }
];

var FCMB_PREVIEW_SHEET = "FCMB Schedule Email - PREVIEW";

/**
 * The schedule date for the email: the Daily Operations
 * date when one is set, otherwise today. Same rule as
 * the agent schedule emails.
 */
function fcmbScheduleDateLabel() {
  try {
    var ss = SpreadsheetApp.getActive();
    var ops = ss.getSheetByName(SHEETS.DAILY_OPS);
    if (ops) {
      var d = ops.getRange(OPS_DATA_START_ROW, 1, 1, 1).getValue();
      if (isValidDate(d)) return formatDateKey(d);
    }
  } catch (e) {}
  return formatDateKey(new Date());
}

function fcmbRecipientList() {
  return FCMB_SCHEDULE_RECIPIENTS.map(function(r) {
    return r.name + " <" + r.email + ">";
  }).join(", ");
}

function fcmbRecipientEmails() {
  return FCMB_SCHEDULE_RECIPIENTS.map(function(r) {
    return r.email;
  }).join(",");
}

function fcmbScheduleSubject(dateLabel) {
  return "CDFC Daily Schedule for " + formatDateLong(dateLabel);
}

/**
 * Plain-text fallback: one line per agent, so the email
 * is still readable in any client.
 */
function buildFcmbSchedulePlainBody(dateLabel, rows) {
  var lines = [
    "Hello,",
    "",
    "Please find below the daily schedule for " + formatDateLong(dateLabel) + ":",
    ""
  ];
  rows.forEach(function(r) {
    lines.push(r.agent + " — " + r.shift + " shift, " + r.queue +
      ", " + r.slot + ", break " + formatBreakTime(r.out) +
      " – " + formatBreakTime(r.back));
  });
  lines.push("");
  lines.push("Thank you.");
  lines.push("");
  lines.push("— Team Lead Dashboard (CDFC Daily Operations)");
  return lines.join("\n");
}

/**
 * HTML version, same look as the agent schedule emails:
 * teal header, break times prominent, one clean table.
 * Inline styles only (Gmail-safe).
 */
function buildFcmbScheduleHtml(dateLabel, rows) {
  var bodyRows = rows.map(function(r) {
    return "<tr>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #f3f4f6;\">" + escapeHtml(r.agent) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #f3f4f6;\">" + escapeHtml(r.centre) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #f3f4f6;\">" + escapeHtml(r.shift) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #f3f4f6;\">" + escapeHtml(r.queue) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #f3f4f6;\">" + escapeHtml(r.slot) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #f3f4f6;font-weight:bold;\">" + escapeHtml(formatBreakTime(r.out)) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #f3f4f6;font-weight:bold;\">" + escapeHtml(formatBreakTime(r.back)) + "</td>" +
      "</tr>";
  }).join("");
  return "<div style=\"font-family:Arial,Helvetica,sans-serif;color:#1f2937;max-width:640px;margin:0 auto;\">" +
    "<div style=\"background:#0b5563;color:#ffffff;padding:20px 24px;border-radius:12px 12px 0 0;\">" +
    "<div style=\"font-size:22px;font-weight:bold;\">CDFC Daily Schedule</div>" +
    "<div style=\"margin-top:4px;font-size:14px;opacity:0.9;\">" + escapeHtml(formatDateLong(dateLabel)) + "</div>" +
    "</div>" +
    "<div style=\"border:1px solid #e5e7eb;border-top:none;padding:24px;border-radius:0 0 12px 12px;\">" +
    "<p style=\"margin:0 0 20px;font-size:15px;line-height:1.5;\">Hello,</p>" +
    "<p style=\"margin:0 0 20px;font-size:15px;line-height:1.5;\">Please find below the daily schedule for " +
    escapeHtml(formatDateLong(dateLabel)) + ".</p>" +
    "<table style=\"width:100%;border-collapse:collapse;font-size:14px;\">" +
    "<tr>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;\">Agent</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;\">Centre</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;\">Shift</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;\">Queue</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;\">Break slot</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;\">Out</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;\">Back</th>" +
    "</tr>" + bodyRows + "</table>" +
    "<p style=\"margin:20px 0 0;font-size:15px;line-height:1.5;\">Thank you.</p>" +
    "<p style=\"color:#6b7280;font-size:13px;line-height:1.5;\">Sent from the Team Lead Dashboard (CDFC Daily Operations).</p>" +
    "</div></div>";
}

/**
 * Step 1 — PREVIEW ONLY. Builds (or rebuilds) the preview
 * sheet and sends nothing. A rebuilt preview is always
 * unsent, so a corrected schedule can be previewed and
 * sent again.
 */
function previewFcmbScheduleEmailMenu() {
  var ui = SpreadsheetApp.getUi();
  var rows = readDailyScheduleRows();
  if (rows.length === 0) {
    ui.alert("Preview Schedule Email to FCMB",
      "No Daily Schedule found. Generate today's schedule first.",
      ui.ButtonSet.OK);
    return;
  }
  var dateLabel = fcmbScheduleDateLabel();
  var ss = SpreadsheetApp.getActive();
  var old = ss.getSheetByName(FCMB_PREVIEW_SHEET);
  if (old) ss.deleteSheet(old);
  var sh = ss.insertSheet(FCMB_PREVIEW_SHEET);
  sh.getRange(1, 1).setValue("STATUS: PREVIEW — NOT SENT").setFontWeight("bold");
  sh.getRange(2, 1).setValue("To").setFontWeight("bold");
  sh.getRange(2, 2).setValue(fcmbRecipientList());
  sh.getRange(3, 1).setValue("Subject").setFontWeight("bold");
  sh.getRange(3, 2).setValue(fcmbScheduleSubject(dateLabel));
  sh.getRange(4, 1).setValue("Schedule date").setFontWeight("bold");
  sh.getRange(4, 2).setValue(dateLabel);
  sh.getRange(6, 1, 1, 7).setValues([["Agent", "Centre", "Shift", "Queue", "Break slot", "Out", "Back"]])
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  var values = rows.map(function(r) {
    return [r.agent, r.centre, r.shift, r.queue, r.slot,
      formatBreakTime(r.out), formatBreakTime(r.back)];
  });
  sh.getRange(7, 1, values.length, 7).setValues(values);
  sh.setFrozenRows(6);
  sh.autoResizeColumns(1, 7);
  ui.alert("Preview Schedule Email to FCMB",
    "PREVIEW ONLY — nothing has been sent.\n\n" +
    "To: " + fcmbRecipientList() + "\n" +
    "Subject: " + fcmbScheduleSubject(dateLabel) + "\n" +
    "Agents in schedule: " + rows.length + "\n\n" +
    "Read the sheet \"" + FCMB_PREVIEW_SHEET + "\" to check the exact table. " +
    "When it is right, run Team Lead → Send Previewed Schedule Email.",
    ui.ButtonSet.OK);
}

/**
 * Step 2 — SEND, only from a preview sheet, only once
 * per preview, and only after a Yes/No confirmation.
 * Sends exactly the previewed table.
 */
function sendPreviewedFcmbScheduleEmailMenu() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(FCMB_PREVIEW_SHEET);
  if (!sh) {
    ui.alert("Send Previewed Schedule Email",
      "No preview found. Run Team Lead → Preview Schedule Email to FCMB first, and read it before sending.",
      ui.ButtonSet.OK);
    return;
  }
  var status = String(sh.getRange(1, 1).getValue() || "");
  if (status.indexOf("SENT") >= 0) {
    ui.alert("Send Previewed Schedule Email",
      "This preview has already been sent. To send a corrected schedule, run Preview Schedule Email to FCMB again first.",
      ui.ButtonSet.OK);
    return;
  }
  var dateLabel = String(sh.getRange(4, 2).getValue() || fcmbScheduleDateLabel());
  var data = sh.getRange(7, 1, Math.max(sh.getLastRow() - 6, 0), 7).getValues();
  var rows = [];
  data.forEach(function(d) {
    if (String(d[0] || "").trim() !== "") {
      rows.push({ agent: d[0], centre: d[1], shift: d[2], queue: d[3], slot: d[4], out: d[5], back: d[6] });
    }
  });
  if (rows.length === 0) {
    ui.alert("Send Previewed Schedule Email",
      "The preview sheet has no schedule rows. Run Preview Schedule Email to FCMB again.",
      ui.ButtonSet.OK);
    return;
  }
  var confirm = ui.alert("Send Previewed Schedule Email",
    "Send this schedule to:\n" + fcmbRecipientList() + "\n\n" +
    "Subject: " + fcmbScheduleSubject(dateLabel) + "\n" +
    "Agents: " + rows.length + "\n\n" +
    "This sends a real email to FCMB. Are you sure?",
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  try {
    MailApp.sendEmail({
      to: fcmbRecipientEmails(),
      subject: fcmbScheduleSubject(dateLabel),
      body: buildFcmbSchedulePlainBody(dateLabel, rows),
      htmlBody: buildFcmbScheduleHtml(dateLabel, rows)
    });
  } catch (e) {
    ui.alert("Send Previewed Schedule Email",
      "Send failed — " + e.message + "\n\nNothing was marked sent. Fix the permission or address and try again.",
      ui.ButtonSet.OK);
    return;
  }
  sh.getRange(1, 1).setValue("STATUS: SENT on " + new Date().toLocaleString("en-GB", { timeZone: "Africa/Lagos" }));
  try {
    logNotification("FCMB parties", "Email", "Schedule share",
      fcmbScheduleSubject(dateLabel) + " (" + rows.length + " agents)", "Sent");
  } catch (e) {}
  ui.alert("Send Previewed Schedule Email",
    "Sent to " + fcmbRecipientList() + "\n\n" +
    "Subject: " + fcmbScheduleSubject(dateLabel) + "\n" +
    "Agents: " + rows.length,
    ui.ButtonSet.OK);
}
