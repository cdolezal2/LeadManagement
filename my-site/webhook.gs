// ================================================================
// BOW LANE LEADS — Prospect Intake Script
// Captures form submissions from bowlaneleads.com
// ================================================================
// SETUP:
// 1. Delete all existing code in the Apps Script editor
// 2. Paste this entire script
// 3. Fill in your Anthropic API key in CONFIG
// 4. Deploy → New deployment → Web App
//    Set: Execute as = Me | Who has access = Anyone
// 5. Copy the Web App URL → paste into index.html as WEBHOOK_URL
// ================================================================

// ── CONFIG ───────────────────────────────────────────────────────
const CONFIG = {
  ANTHROPIC_API_KEY:  'YOUR_ANTHROPIC_API_KEY_HERE',
  NOTIFICATION_EMAIL: 'creighton@bowlaneleads.com',
  BUSINESS_NAME:      'Bow Lane Leads',
  LEAD_TYPE:          'prospect',
};
// ─────────────────────────────────────────────────────────────────

const HEADERS = [
  'Timestamp', 'Name', 'Business', 'Email', 'Phone',
  'Industry', 'Package Interest', 'Notes', 'AI Priority', 'AI Summary', 'Status'
];

// ── MAIN FUNCTION ─────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    setupSheet(sheet);
    sheet.appendRow(buildRow(data));
    const summary = getAiSummary(data);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 9).setValue(summary.priority);
    sheet.getRange(lastRow, 10).setValue(summary.summary);
    sheet.getRange(lastRow, 11).setValue('New');
    colorRow(sheet, lastRow, summary.priority);
    sendEmail(data, summary);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('Error:', err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── SHEET SETUP ───────────────────────────────────────────────────
function setupSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    const hRange = sheet.getRange(1, 1, 1, HEADERS.length);
    hRange.setBackground('#1a3a2a');
    hRange.setFontColor('#f5edd8');
    hRange.setFontWeight('bold');
    hRange.setFontSize(11);
    sheet.setFrozenRows(1);
    [160, 150, 200, 180, 140, 160, 220, 280, 90, 360, 100]
      .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
}

// ── BUILD ROW ─────────────────────────────────────────────────────
function buildRow(data) {
  const ts = Utilities.formatDate(
    new Date(data.submitted_at || new Date()),
    Session.getScriptTimeZone(), 'MMM d, yyyy  h:mm a'
  );
  return [
    ts, data.name||'', data.business||'', data.email||'', data.phone||'',
    data.industry||'', data.package||'', data.notes||''
  ];
}

// ── COLOR CODE BY PRIORITY ────────────────────────────────────────
function colorRow(sheet, rowNum, priority) {
  const colors = { High: '#fce8e8', Medium: '#fef9e7', Low: '#e8f5e9' };
  sheet.getRange(rowNum, 1, 1, HEADERS.length)
    .setBackground(colors[priority] || '#f5edd8');
}

// ── CALL CLAUDE AI ────────────────────────────────────────────────
function getAiSummary(data) {
  const prompt = `You are a prospect intake assistant for Bow Lane Leads, a company that builds websites and lead management systems for small home service businesses.

Analyze this inbound prospect and respond with JSON ONLY — no markdown, no explanation.

Name: ${data.name||'Not provided'}
Business: ${data.business||'N/A'}
Email: ${data.email||'Not provided'}
Phone: ${data.phone||'Not provided'}
Industry: ${data.industry||'Not specified'}
Package Interest: ${data.package||'Not specified'}
Notes: ${data.notes||'None'}

Return exactly this JSON:
{
  "priority": "High, Medium, or Low",
  "next_step": "one specific action to take with this prospect",
  "summary": "2-3 sentence plain English summary of the prospect and fit",
  "flags": "any red flags or concerns, or None"
}

Priority rules:
High = strong buying signals, mentions urgency, no website, or package already selected.
Medium = interested but vague, has a website but wants lead management, or needs more info.
Low = just exploring, no clear need, or outside target market.`;

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': CONFIG.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 400,
      messages: [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  try {
    const result = JSON.parse(response.getContentText());
    const text = result.content[0].text.replace(/```json|```/g, '').trim();
    return JSON.parse(text);
  } catch(e) {
    return {
      priority: 'Medium',
      next_step: 'Follow up within 24 hours.',
      summary: 'Prospect received — AI summary unavailable. Review manually.',
      flags: 'AI parsing failed — check your Anthropic API key in CONFIG.',
    };
  }
}

// ── SEND EMAIL ────────────────────────────────────────────────────
function sendEmail(data, summary) {
  const emoji = { High: '🔴', Medium: '🟡', Low: '🟢' }[summary.priority] || '⚪';
  const subject = `${emoji} New Prospect — ${data.business||'No Business Listed'} | ${data.name||'Unknown'}`;
  const body = [
    `New prospect from bowlaneleads.com`,
    '─'.repeat(48),
    `${summary.summary}`,
    `Priority: ${summary.priority}`,
    '─'.repeat(48),
    `Name:     ${data.name||'—'}`,
    `Business: ${data.business||'—'}`,
    `Email:    ${data.email||'—'}`,
    `Phone:    ${data.phone||'—'}`,
    '',
    `Industry: ${data.industry||'—'}`,
    `Package:  ${data.package||'—'}`,
    '',
    `Notes: ${data.notes||'None provided.'}`,
    '─'.repeat(48),
  ].join('\n');
  GmailApp.sendEmail(CONFIG.NOTIFICATION_EMAIL, subject, body);
}

// ── TEST FUNCTION ─────────────────────────────────────────────────
// Select doTest from the function dropdown and click Run to test.
function doTest() {
  const fakeData = { postData: { contents: JSON.stringify({
    name: 'Sean Burns',
    business: 'Burns Plumbing',
    email: 'sean@burnsplumbing.com',
    phone: '(574) 555-0192',
    industry: 'Plumbing / HVAC / Electrical',
    package: 'Website + Lead Management — $750 setup / $79/mo',
    notes: 'No website currently. Gets most work from word of mouth but thinks he is losing Google searches.',
    submitted_at: new Date().toISOString(),
  })}};
  const result = doPost(fakeData);
  console.log('Test result:', result.getContent());
}
