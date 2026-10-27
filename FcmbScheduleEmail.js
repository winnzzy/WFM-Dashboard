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
 * Centre for this schedule, supplied by Winner 2026-10-27:
 * Abuja for every agent. It is shown once in the header,
 * not as a per-agent column — a Centre column repeating
 * the same value only made the table wide for no reason.
 * (The Daily Schedule Centre column itself is still blank
 * in the sheet; filling Agent Roster Centre with Abuja
 * would carry it through generation, but the FCMB email
 * does not wait for that.)
 */
var FCMB_CENTRE = "Abuja";

/**
 * Signature for script-sent emails — taken from Gmail
 * itself, never typed in by hand (Winner, 2026-10-27).
 *
 * Google only adds a Gmail signature in the Gmail
 * web/app composer; MailApp sends from Apps Script do
 * NOT get it automatically, which is why it looked
 * "removed". The fix is to read the signature Winner
 * already has set in Gmail (Gmail API, sendAs settings)
 * and append that exact HTML at send time. Needs the
 * gmail.settings.basic scope in appsscript.json.
 *
 * If the fetch fails, or no signature is set in Gmail,
 * the email still sends — with just "Warm regards," and
 * no invented name, title or phone number. A signature
 * failure must never fail the schedule email.
 */
var gmailSignatureCache_ = null;

var gmailSignatureStatus_ = "";

function getGmailSignature_() {
  if (gmailSignatureCache_) return gmailSignatureCache_;
  var empty = { html: "", plain: "" };
  try {
    var email = Session.getEffectiveUser().getEmail();
    if (!email) {
      gmailSignatureStatus_ = "No effective user email.";
      return empty;
    }
    // List sendAs addresses and pick the primary one, instead
    // of guessing a single sendAs URL — that single-URL call
    // is the likeliest reason the first test came back with
    // no signature even though one may be set.
    var resp = UrlFetchApp.fetch(
      "https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs",
      {
        method: "get",
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      }
    );
    if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
      gmailSignatureStatus_ = "Gmail settings fetch failed: HTTP " + resp.getResponseCode() +
        " (check the gmail.settings.basic scope was allowed on re-authorisation).";
      return empty;
    }
    var data = JSON.parse(resp.getContentText());
    var list = data.sendAs || [];
    var chosen = null;
    for (var i = 0; i < list.length; i++) {
      if (list[i].isPrimary) { chosen = list[i]; break; }
    }
    if (!chosen) {
      for (var j = 0; j < list.length; j++) {
        if (String(list[j].sendAsEmail || "") === email) { chosen = list[j]; break; }
      }
    }
    if (!chosen && list.length > 0) chosen = list[0];
    if (!chosen) {
      gmailSignatureStatus_ = "No sendAs address found for " + email + ".";
      return empty;
    }
    var html = String(chosen.signature || "");
    if (!html.trim()) {
      gmailSignatureStatus_ = "No signature is set in Gmail for " + chosen.sendAsEmail +
        " — set one in Gmail Settings → Signature, then retest. Nothing was invented.";
      return empty;
    }
    gmailSignatureStatus_ = "Signature found for " + chosen.sendAsEmail + " (" + html.length + " characters).";
    var plain = html
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/p>/gi, "\n")
      .replace(/<\/div>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
    gmailSignatureCache_ = { html: html, plain: plain };
    return gmailSignatureCache_;
  } catch (e) {
    return empty;
  }
}

function fcmbSignatureHtml() {
  var sig = getGmailSignature_();
  return "<p style=\"margin:20px 0 0;font-size:15px;line-height:1.5;\">Warm regards,</p>" +
    (sig.html ? "<div style=\"font-size:15px;line-height:1.5;\">" + sig.html + "</div>" : "") +
    "<p style=\"color:#6b7280;font-size:12px;line-height:1.5;\">Sent via the Team Lead Dashboard (CDFC Daily Operations).</p>";
}

function fcmbSignaturePlain() {
  var sig = getGmailSignature_();
  return "Warm regards,\n" +
    (sig.plain ? sig.plain + "\n\n" : "") +
    "Sent via the Team Lead Dashboard (CDFC Daily Operations).";
}

/**
 * Diagnostic only, editor-run like the test send: reports
 * whether a Gmail signature was found, for which address,
 * or exactly why not — so a missing signature is diagnosed
 * instead of guessed at. Sends nothing.
 */
function checkGmailSignature() {
  gmailSignatureCache_ = null;
  gmailSignatureStatus_ = "";
  var sig = getGmailSignature_();
  SpreadsheetApp.getUi().alert("Gmail Signature Check",
    "Account: " + Session.getEffectiveUser().getEmail() + "\n\n" +
    gmailSignatureStatus_ + "\n\n" +
    (sig.html ? "Signature preview (plain text):\n" + sig.plain : "No signature will be appended to the next send."),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

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
    "Centre: " + FCMB_CENTRE,
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
  lines.push(fcmbSignaturePlain());
  return lines.join("\n");
}

/**
 * HTML version, same look as the agent schedule emails:
 * teal header, break times prominent, one clean table.
 * Inline styles only (Gmail-safe).
 */
function buildFcmbScheduleHtml(dateLabel, rows) {
  // Fixed table layout + nowrap times: Gmail was wrapping
  // "11:00 AM" across two lines and stretching the Agent
  // column, which is what made the test email look ugly.
  // Column widths are fixed, times never wrap, every
  // second row is shaded so rows stay easy to follow.
  var bodyRows = rows.map(function(r, i) {
    var bg = (i % 2 === 1) ? "background:#f9fafb;" : "";
    return "<tr style=\"" + bg + "\">" +
      "<td style=\"padding:8px 10px;border-top:1px solid #e5e7eb;line-height:1.4;\">" + escapeHtml(r.agent) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #e5e7eb;white-space:nowrap;\">" + escapeHtml(r.shift) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #e5e7eb;white-space:nowrap;\">" + escapeHtml(r.queue) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #e5e7eb;white-space:nowrap;\">" + escapeHtml(r.slot) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #e5e7eb;font-weight:bold;white-space:nowrap;\">" + escapeHtml(formatBreakTime(r.out)) + "</td>" +
      "<td style=\"padding:8px 10px;border-top:1px solid #e5e7eb;font-weight:bold;white-space:nowrap;\">" + escapeHtml(formatBreakTime(r.back)) + "</td>" +
      "</tr>";
  }).join("");
  return "<div style=\"font-family:Arial,Helvetica,sans-serif;color:#1f2937;max-width:640px;margin:0 auto;\">" +
    "<div style=\"background:#0b5563;color:#ffffff;padding:20px 24px;border-radius:12px 12px 0 0;\">" +
    "<div style=\"font-size:22px;font-weight:bold;\">CDFC Daily Schedule</div>" +
    "<div style=\"margin-top:4px;font-size:14px;opacity:0.9;\">" + escapeHtml(formatDateLong(dateLabel)) + " &middot; Centre: " + escapeHtml(FCMB_CENTRE) + "</div>" +
    "</div>" +
    "<div style=\"border:1px solid #e5e7eb;border-top:none;padding:24px;border-radius:0 0 12px 12px;\">" +
    "<p style=\"margin:0 0 20px;font-size:15px;line-height:1.5;\">Hello,</p>" +
    "<p style=\"margin:0 0 20px;font-size:15px;line-height:1.5;\">Please find below the daily schedule for " +
    escapeHtml(formatDateLong(dateLabel)) + ".</p>" +
    "<table style=\"width:100%;table-layout:fixed;border-collapse:collapse;font-size:14px;\">" +
    "<colgroup><col style=\"width:28%\"><col style=\"width:15%\"><col style=\"width:15%\"><col style=\"width:12%\"><col style=\"width:15%\"><col style=\"width:15%\"></colgroup>" +
    "<tr style=\"background:#eaf6f5;\">" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;border-bottom:2px solid #0b5563;\">Agent</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;border-bottom:2px solid #0b5563;white-space:nowrap;\">Shift</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;border-bottom:2px solid #0b5563;white-space:nowrap;\">Queue</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;border-bottom:2px solid #0b5563;white-space:nowrap;\">Break slot</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;border-bottom:2px solid #0b5563;white-space:nowrap;\">Out</th>" +
    "<th style=\"padding:8px 10px;text-align:left;color:#0b5563;border-bottom:2px solid #0b5563;white-space:nowrap;\">Back</th>" +
    "</tr>" + bodyRows + "</table>" +
    "<p style=\"margin:20px 0 0;font-size:15px;line-height:1.5;\">Thank you.</p>" +
    fcmbSignatureHtml() +
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
  sh.getRange(5, 1).setValue("Centre").setFontWeight("bold");
  sh.getRange(5, 2).setValue(FCMB_CENTRE);
  sh.getRange(6, 1, 1, 6).setValues([["Agent", "Shift", "Queue", "Break slot", "Out", "Back"]])
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");
  var values = rows.map(function(r) {
    return [r.agent, r.shift, r.queue, r.slot,
      formatBreakTime(r.out), formatBreakTime(r.back)];
  });
  sh.getRange(7, 1, values.length, 6).setValues(values);
  sh.setFrozenRows(6);
  // Fixed widths, not auto-resize: the table stays tight
  // like the break emails instead of stretching a column.
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 90);
  sh.setColumnWidth(6, 90);
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
  var data = sh.getRange(7, 1, Math.max(sh.getLastRow() - 6, 0), 6).getValues();
  var rows = [];
  data.forEach(function(d) {
    if (String(d[0] || "").trim() !== "") {
      rows.push({ agent: d[0], shift: d[1], queue: d[2], slot: d[3], out: d[4], back: d[5] });
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

/**
 * TEST ONLY — not in the Team Lead menu, on purpose.
 * Winner asked (2026-10-27) to test the FCMB schedule
 * email to his own nnamuahn@gmail.com after the push,
 * before anything goes to FCMB.
 *
 * Run from Extensions -> Apps Script, function dropdown:
 * sendTestFcmbScheduleEmail. It sends the current Daily
 * Schedule, in the exact FCMB format (HTML table, Centre
 * Abuja header, Gmail signature fetched at send time),
 * to nnamuahn@gmail.com ONLY. It does not touch the
 * preview sheet, does not mark any preview SENT, and
 * does not email Keziah, Oladapo or Sarah. Delete this
 * function once the test has passed, if Winner wants
 * it gone.
 */
function sendTestFcmbScheduleEmail() {
  var ui = SpreadsheetApp.getUi();
  var rows = readDailyScheduleRows();
  if (rows.length === 0) {
    ui.alert("Test Schedule Email",
      "No Daily Schedule found. Generate today's schedule first, then run this test again.",
      ui.ButtonSet.OK);
    return;
  }
  var dateLabel = fcmbScheduleDateLabel();
  try {
    MailApp.sendEmail({
      to: "nnamuahn@gmail.com",
      subject: "[TEST] " + fcmbScheduleSubject(dateLabel),
      body: buildFcmbSchedulePlainBody(dateLabel, rows),
      htmlBody: buildFcmbScheduleHtml(dateLabel, rows)
    });
  } catch (e) {
    ui.alert("Test Schedule Email",
      "Test send failed — " + e.message,
      ui.ButtonSet.OK);
    return;
  }
  ui.alert("Test Schedule Email",
    "Test sent to nnamuahn@gmail.com only.\n\n" +
    "Subject: [TEST] " + fcmbScheduleSubject(dateLabel) + "\n" +
    "Agents: " + rows.length + "\n\n" +
    "Nothing was sent to FCMB. Check the nnamuahn@gmail.com inbox for the table, Centre Abuja header and your Gmail signature.",
    ui.ButtonSet.OK);
}
