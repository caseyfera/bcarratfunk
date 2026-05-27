// ============================================================
//  DRAGONFLY v2.96 — Google Apps Script Backend (Complete)
//  Deploy as: Web App → Execute as Me → Anyone
//
//  Uses your "Subscribers" sheet:
//  A: Email   B: Date Subscribed   C: code   D: expiry   E: session_token
// ============================================================

const SHEET_NAME     = 'Subscribers';
const SPREADSHEET_ID = '1fmLHmZ0tzO4RtGS4Zmm8Z2625Wj-i9ytADWud62S16M';
const SS             = SpreadsheetApp.openById(SPREADSHEET_ID);

// ── Router ──────────────────────────────────────────────────
function doGet(e) {
  if (!e || !e.parameter) return text('OK');
  const p = e.parameter;
  if (p.action === 'sendCode' && p.email) return sendLoginCode(p);
  if (p.sessionCheck === 'true')          return checkSession(p);
  if (p.token && !p.action)              return verifyToken(p);
  return text('OK');
}

function doPost(e) {
  if (!e || !e.parameter) return text('Error');
  const p = e.parameter;
  if (p.email) return sendLoginCode(p);
  return text('Error');
}

// ── Send login code ──────────────────────────────────────────
function sendLoginCode(p) {
  const email = (p.email || '').toLowerCase().trim();
  if (!email.includes('@')) return text('InvalidEmail');

  const sheet = SS.getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toLowerCase() === email) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    sheet.appendRow([email, new Date()]);
    rowIndex = sheet.getLastRow();
  }

  try {
    if (MailApp.getRemainingDailyQuota() < 1) return text('QuotaExceeded');
  } catch(e) {}

  const code   = Math.floor(100000 + Math.random() * 900000).toString();
  const expiry = new Date(Date.now() + 5 * 60 * 1000).toISOString();

  sheet.getRange(rowIndex, 3).setValue(code);
  sheet.getRange(rowIndex, 4).setValue(expiry);

  MailApp.sendEmail({
    to: email,
    subject: '🐉 Dragonfly Login Code',
    htmlBody: `
      <div style="font-family:sans-serif;max-width:400px;margin:0 auto;background:#0a1a0f;color:#fff;padding:28px;border-radius:12px;border:1px solid #d4af37;">
        <h2 style="color:#d4af37;letter-spacing:3px;margin:0 0 16px;">DRAGONFLY</h2>
        <p style="color:rgba(255,255,255,0.7);">Your login code:</p>
        <div style="font-size:2.5rem;font-weight:bold;color:#d4af37;letter-spacing:8px;margin:12px 0;">${code}</div>
        <p style="font-size:0.8rem;color:rgba(255,255,255,0.4);">Valid for 5 minutes. Do not share this code.</p>
        <hr style="border-color:rgba(212,175,55,0.2);margin:16px 0;">
        <p style="font-size:0.7rem;color:rgba(255,255,255,0.25);">© 2025 Dragonfly · All Rights Reserved</p>
      </div>`
  });

  return text('CodeSent');
}

// ── Verify code → issue session token ───────────────────────
function verifyToken(p) {
  const email = (p.email || '').toLowerCase().trim();
  const token = (p.token || '').trim();
  const sheet = SS.getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toLowerCase() !== email) continue;

    const storedCode = String(data[i][2] || '').trim();
    const expiry     = data[i][3];

    if (storedCode !== token)          return text('InvalidCode');
    if (new Date() > new Date(expiry)) return text('ExpiredCode');

    const sessionToken  = Utilities.getUuid();
    const sessionExpiry = new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString();
    sheet.getRange(i + 1, 3).setValue('');
    sheet.getRange(i + 1, 4).setValue(sessionExpiry);
    sheet.getRange(i + 1, 5).setValue(sessionToken);

    const sub = JSON.stringify({ status: 'active', expiry: null, handsRemaining: null });
    return text('Verified:' + sessionToken + ':' + sub);
  }
  return text('NotFound');
}

// ── Session check ────────────────────────────────────────────
function checkSession(p) {
  const email = (p.email || '').toLowerCase().trim();
  const token = (p.token || '').trim();
  const sheet = SS.getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toLowerCase() !== email) continue;

    const stored = String(data[i][4] || '').trim();
    const expiry = data[i][3];

    if (stored !== token)                  return text('SessionInvalid');
    if (new Date() > new Date(expiry))     return text('SessionExpired');

    sheet.getRange(i + 1, 4).setValue(
      new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString()
    );
    return text('SessionValid');
  }
  return text('SessionInvalid');
}

// ── Helpers ──────────────────────────────────────────────────
function text(s) { return ContentService.createTextOutput(s); }

// ============================================================
//  NEWSLETTER
//  To send: select sendV295Newsletter from the dropdown → Run
// ============================================================

function buildV295Changelog() {
  return '<p style="color:rgba(255,255,255,0.75);font-size:0.85rem;line-height:1.8;margin:0 0 20px;">Welcome to <strong style="color:#d4af37;">Dragonfly</strong> — a professional baccarat simulator built for players who want to study the game seriously. If you signed up a while back and have not had a chance to explore it yet, here is what it does:</p>'
    + '<div style="color:#d4af37;font-size:0.75rem;font-weight:bold;letter-spacing:1.5px;margin-bottom:10px;">WHAT DRAGONFLY IS</div>'
    + '<ul style="margin:0 0 20px;padding-left:18px;color:rgba(255,255,255,0.8);line-height:2;font-size:0.83rem;">'
    + '<li>A full baccarat shoe simulator with authentic casino dealing rules</li>'
    + '<li>Bet Player, Banker, Tie — plus Dragon 7 and Panda 8 side bets</li>'
    + '<li>Real-time bankroll tracking, win rate, and hand-by-hand P&L</li>'
    + '<li>All five road maps — Bead Plate, Big Road, Big Eye Boy, Small Road, Cockroach Pig</li>'
    + '<li>Full shoe inspector — every card dealt, in order, color coded by seat</li>'
    + '<li>Customizable layout so you can arrange the table the way you think</li>'
    + '<li>Save and reload sessions — pick up any shoe exactly where you left off</li>'
    + '</ul>'
    + '<div style="color:#d4af37;font-size:0.75rem;font-weight:bold;letter-spacing:1.5px;margin-bottom:10px;">NEW IN v2.95</div>'
    + '<ul style="margin:0 0 8px;padding-left:18px;color:rgba(255,255,255,0.8);line-height:2;font-size:0.83rem;">'
    + '<li>Login fixed — email codes now deliver reliably every time</li>'
    + '<li>Game survives browser close — come back days later and pick up right where you left off</li>'
    + '<li>Hand history shows actual cards dealt, in order, color coded Player vs Banker</li>'
    + '<li>Rebet remembers your last wager even if you skipped a hand</li>'
    + '<li>Win/loss result moved to a slim side panel — table stays fully visible while you play</li>'
    + '</ul>';
}

function sendV295Newsletter() {
  // ── TEST MODE: only sends to you ──
  // Change false to true when ready to send to everyone
  var testMode = true;
  if (testMode) {
    sendNewsletterToOne('caseyfera@gmail.com', '2.95', buildV295Changelog());
    return;
  }
  sendNewsletter('2.95', buildV295Changelog());
}

function sendNewsletterToOne(email, version, changelog) {
  Logger.log('Sending test newsletter to: ' + email);
  var subject = 'Dragonfly v' + version + " — What's New";
  MailApp.sendEmail({ to: email, subject: subject, htmlBody: buildNewsletterHtml(version, changelog, email) });
  Logger.log('Test email sent to ' + email);
}

function sendNewsletter(version, changelog) {
  const sheet = SS.getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();

  const recipients = [];
  for (let i = 1; i < data.length; i++) {
    const email = (data[i][0] || '').toLowerCase().trim();
    if (email.includes('@')) recipients.push(email);
  }

  if (recipients.length === 0) { Logger.log('No recipients found.'); return; }

  const quota = MailApp.getRemainingDailyQuota();
  if (quota < recipients.length) {
    Logger.log(`Only ${quota} emails left today, need ${recipients.length}. Aborting.`);
    return;
  }

  Logger.log(`Sending v${version} newsletter to ${recipients.length} recipients...`);

  const subject = `🐉 Dragonfly v${version} — What's New`;
  let sent = 0, failed = 0;

  recipients.forEach((email, idx) => {
    try {
      MailApp.sendEmail({ to: email, subject: subject, htmlBody: buildNewsletterHtml(version, changelog, email) });
      sent++;
    } catch(e) {
      Logger.log(`Failed: ${email} — ${e}`);
      failed++;
    }
    if (idx > 0 && idx % 20 === 0) Utilities.sleep(500);
  });

  Logger.log(`Done: ${sent} sent, ${failed} failed.`);
}

function buildNewsletterHtml(ver, changelog, email) {
  var header = '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#050e07;font-family:Helvetica Neue,Arial,sans-serif;">'
    + '<div style="max-width:480px;margin:0 auto;padding:32px 20px;">'
    + '<div style="text-align:center;margin-bottom:28px;">'
    + '<div style="font-size:2rem;font-weight:bold;letter-spacing:6px;color:#d4af37;">DRAGONFLY</div>'
    + '<div style="font-size:0.8rem;letter-spacing:3px;color:rgba(212,175,55,0.55);margin-top:4px;">BACCARAT TRAINING</div>'
    + '</div>'
    + '<div style="background:#0b2e1a;border:1px solid rgba(212,175,55,0.4);border-radius:12px;padding:24px;margin-bottom:20px;">'
    + '<div style="display:inline-block;background:rgba(212,175,55,0.15);border:1px solid rgba(212,175,55,0.5);border-radius:6px;padding:4px 12px;color:#d4af37;font-weight:bold;font-size:0.85rem;letter-spacing:1px;margin-bottom:16px;">v' + ver + '</div>';

  var footer = '</div>'
    + '<div style="text-align:center;margin-bottom:28px;">'
    + '<a href="https://baccaratfunk.win" style="display:inline-block;background:#d4af37;color:#000;font-weight:bold;letter-spacing:2px;padding:12px 36px;border-radius:50px;text-decoration:none;font-size:0.9rem;">OPEN DRAGONFLY</a>'
    + '</div>'
    + '<div style="border-top:1px solid rgba(212,175,55,0.15);margin-bottom:20px;"></div>'
    + '<div style="font-size:0.68rem;color:rgba(255,255,255,0.2);text-align:center;line-height:1.8;">'
    + '2025 Dragonfly. All Rights Reserved.<br>'
    + 'You received this because you have a Dragonfly account.<br>'
    + 'Sent to: ' + email
    + '</div></div></body></html>';

  return header + changelog + footer;
}
