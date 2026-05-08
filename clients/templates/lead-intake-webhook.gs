/**
 * Lead Intake Webhook — Google Apps Script
 * ─────────────────────────────────────────────────────────────────────────────
 * Handles two types of incoming leads:
 *   1. Form submissions — JSON POST from the client's website contact form
 *   2. Voicemail transcriptions — form-encoded POST from Twilio after a missed call
 *
 * For every lead (either type), this script:
 *   - Logs the lead to the client's Google Sheet
 *   - Sends an instant email alert to the client
 *   - Sends an instant SMS alert via Twilio
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * SETUP — SCRIPT PROPERTIES (one-time per client)
 *   Apps Script editor → Project Settings (gear) → Script Properties
 *
 *   Property Name          | Value
 *   ───────────────────────|──────────────────────────────────────────────────
 *   NOTIFY_EMAIL           | Client's email address for lead alerts
 *   NOTIFY_PHONE           | Client's cell — digits + country code, no symbols
 *                          |   e.g. 13125550100  (not +1-312-555-0100)
 *   TWILIO_ACCOUNT_SID     | Your Twilio Account SID (starts with AC)
 *   TWILIO_AUTH_TOKEN      | Your Twilio Auth Token
 *   TWILIO_FROM_NUMBER     | Your Twilio SMS number — e.g. +15551234567
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * SETUP — DEPLOYING AS A WEB APP
 *   Deploy → New deployment → Type: Web App
 *   Execute as: Me  |  Who has access: Anyone
 *   Click Deploy → copy the URL
 *   → Paste into index.html as WEBHOOK_URL
 *   → Paste into twiml-voicemail.xml as [APPS_SCRIPT_WEBHOOK_URL]
 *
 *   NOTE: Editing this script requires a NEW deployment (not updating the
 *   existing one) for changes to take effect. Script Properties changes
 *   do NOT require redeployment.
 * ─────────────────────────────────────────────────────────────────────────────
 */


// ── CONFIGURATION ──────────────────────────────────────────────────────────────

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    notifyEmail:  props.getProperty('NOTIFY_EMAIL'),
    notifyPhone:  props.getProperty('NOTIFY_PHONE'),       // digits only, e.g. 13125550100
    twilioSid:    props.getProperty('TWILIO_ACCOUNT_SID'),
    twilioToken:  props.getProperty('TWILIO_AUTH_TOKEN'),
    twilioFrom:   props.getProperty('TWILIO_FROM_NUMBER'), // e.g. +15551234567
  };
}


// ── SHEET CONFIGURATION ────────────────────────────────────────────────────────

const HEADERS = [
  'Timestamp',
  'Name',
  'Phone',
  'Email',
  'Service',
  'Area',
  'Notes',
  'Source',   // "Form" or "Voicemail"
  'Status',
];

const SHEET_NAME = 'Leads';


// ── MAIN WEBHOOK HANDLER ───────────────────────────────────────────────────────
// Twilio transcription callbacks are form-encoded and always include
// TranscriptionText as a named parameter. Form submissions arrive as JSON.
// We use this to route each request to the correct handler.

function doPost(e) {
  try {
    if (e.parameter && e.parameter.TranscriptionText !== undefined) {
      handleVoicemailTranscription(e.parameter);
    } else {
      const raw  = e.postData ? e.postData.contents : '{}';
      handleFormSubmission(JSON.parse(raw));
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── FORM SUBMISSION HANDLER ────────────────────────────────────────────────────

function handleFormSubmission(data) {
  const lead = {
    timestamp: now(),
    name:      (data.name    || '').trim(),
    phone:     (data.phone   || '').trim(),
    email:     (data.email   || '').trim(),
    service:   (data.service || '').trim(),
    area:      (data.area    || '').trim(),
    notes:     (data.notes   || '').trim(),
    source:    'Form',
    status:    'New',
  };

  logToSheet(lead);
  sendEmailAlert(lead);
  sendSmsAlert(lead);
}


// ── VOICEMAIL TRANSCRIPTION HANDLER ───────────────────────────────────────────
// Called when Twilio POSTs a transcription to this webhook URL.
// Twilio parameters: TranscriptionText, TranscriptionStatus, From, To, CallSid

function handleVoicemailTranscription(params) {
  const transcription = (params.TranscriptionStatus === 'failed' || !params.TranscriptionText)
    ? '(transcription unavailable — listen to voicemail in Twilio console)'
    : params.TranscriptionText;

  const lead = {
    timestamp: now(),
    name:      'Voicemail caller',
    phone:     params.From || 'Unknown',
    email:     '',
    service:   'See voicemail notes',
    area:      '',
    notes:     transcription,
    source:    'Voicemail',
    status:    'New',
  };

  logToSheet(lead);
  sendVoicemailEmailAlert(lead);
  sendSmsAlert(lead);
}


// ── LOG TO GOOGLE SHEET ────────────────────────────────────────────────────────

function logToSheet(lead) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#2c1f0e')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    applyStatusDropdown(sheet);
  }

  sheet.appendRow([
    lead.timestamp,
    lead.name,
    lead.phone,
    lead.email,
    lead.service,
    lead.area,
    lead.notes,
    lead.source,
    lead.status,
  ]);
}


// ── EMAIL ALERT — FORM LEAD ────────────────────────────────────────────────────

function sendEmailAlert(lead) {
  const config  = getConfig();
  const subject = 'New lead: ' + lead.name + ' — ' + lead.service;
  const body    = [
    'New form lead received',
    '─────────────────────',
    'Name:    ' + lead.name,
    'Phone:   ' + lead.phone,
    'Email:   ' + lead.email,
    'Service: ' + lead.service,
    'Area:    ' + lead.area,
    'Notes:   ' + lead.notes,
    '─────────────────────',
    'Received: ' + lead.timestamp,
    '',
    'Open your lead log to update the status after you call back.',
  ].join('\n');

  MailApp.sendEmail({ to: config.notifyEmail, subject: subject, body: body });
}


// ── EMAIL ALERT — VOICEMAIL LEAD ──────────────────────────────────────────────

function sendVoicemailEmailAlert(lead) {
  const config  = getConfig();
  const subject = 'New voicemail from ' + lead.phone;
  const body    = [
    'A customer left a voicemail — transcription below.',
    '─────────────────────',
    'Caller:  ' + lead.phone,
    'Time:    ' + lead.timestamp,
    '',
    'Transcription:',
    lead.notes,
    '─────────────────────',
    'Call this number back as soon as possible.',
    'Open your lead log to update the status once you have reached them.',
  ].join('\n');

  MailApp.sendEmail({ to: config.notifyEmail, subject: subject, body: body });
}


// ── SMS ALERT — BOTH LEAD TYPES ───────────────────────────────────────────────

function sendSmsAlert(lead) {
  const config = getConfig();

  const message = lead.source === 'Voicemail'
    ? 'Voicemail lead: ' + lead.phone + ' — ' + lead.notes.substring(0, 100)
    : 'New lead: ' + lead.name + ' · ' + lead.phone + ' · ' + lead.service + (lead.area ? ' (' + lead.area + ')' : '');

  const url = 'https://api.twilio.com/2010-04-01/Accounts/' + config.twilioSid + '/Messages.json';

  UrlFetchApp.fetch(url, {
    method:  'post',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(config.twilioSid + ':' + config.twilioToken),
    },
    payload: {
      To:   '+' + config.notifyPhone,
      From: config.twilioFrom,
      Body: message,
    },
    muteHttpExceptions: true,
  });
}


// ── MORNING DIGEST ─────────────────────────────────────────────────────────────
// Set up a daily time-based trigger in Apps Script:
// Triggers (clock icon) → Add Trigger → sendMorningDigest → Day timer → 7am–8am

function sendMorningDigest() {
  const config = getConfig();
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const data      = sheet.getDataRange().getValues();
  const rows      = data.slice(1);
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);

  const recentLeads = rows.filter(row => {
    const ts = new Date(row[0]);
    return ts >= yesterday && row[8] === 'New'; // col 8 = Status
  });

  if (recentLeads.length === 0) return;

  const lines = recentLeads.map(row => {
    const source = row[7]; // Source column
    const prefix = source === 'Voicemail' ? '[VM] ' : '[Form] ';
    return '• ' + prefix + row[1] + ' · ' + row[2] + ' · ' + row[4];
  });

  const subject = 'Morning lead digest — ' + recentLeads.length + ' new lead' + (recentLeads.length !== 1 ? 's' : '');
  const body    = [
    'Good morning — here are your new leads from the past 24 hours:',
    '',
    ...lines,
    '',
    'Open your lead log to update statuses as you work through them.',
  ].join('\n');

  MailApp.sendEmail({ to: config.notifyEmail, subject: subject, body: body });
}


// ── STATUS DROPDOWN ────────────────────────────────────────────────────────────
// Applies a dropdown to the entire Status column (rows 2 onward) so clients
// can close out leads with a single click instead of typing.

function applyStatusDropdown(sheet) {
  const statusCol = HEADERS.indexOf('Status') + 1; // column 9
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New', 'Called — No Answer', 'Booked', 'Not Interested', 'Follow Up'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, statusCol, 1000, 1).setDataValidation(rule);
}

// Run this manually from the Apps Script editor (select it in the dropdown,
// click Run) to add the status dropdown to an existing Leads sheet.
function setupStatusDropdown() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('No "' + SHEET_NAME + '" sheet found — submit a test lead first to create it.');
    return;
  }
  applyStatusDropdown(sheet);
  Logger.log('Status dropdown applied to column ' + HEADERS.indexOf('Status') + 1 + '.');
}


// ── HELPERS ────────────────────────────────────────────────────────────────────

function now() {
  return new Date().toLocaleString('en-US', { timeZone: 'America/Chicago' });
}


// ── TEST FUNCTIONS ─────────────────────────────────────────────────────────────
// Select either function from the dropdown and click Run to test without
// submitting a real form or placing a real call.

function testFormSubmission() {
  handleFormSubmission({
    name:    'Test Lead',
    phone:   '555-000-0001',
    email:   'test@example.com',
    service: 'Plumbing repair',
    area:    'Chicago, IL',
    notes:   'Test form submission — please ignore.',
  });
  Logger.log('Form test complete — check Sheet, email, and SMS.');
}

function testVoicemailTranscription() {
  handleVoicemailTranscription({
    TranscriptionText:   'Hi yeah this is Dave calling about a leaking pipe under my kitchen sink, its been going on for a couple days, Im in Naperville, my number is 630-555-0182, please call me back thanks.',
    TranscriptionStatus: 'completed',
    From:                '+16305550182',
    To:                  '+15551234567',
    CallSid:             'CAtest1234567890',
  });
  Logger.log('Voicemail test complete — check Sheet, email, and SMS.');
}
