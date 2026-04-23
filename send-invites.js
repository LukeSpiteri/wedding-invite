/**
 * Wedding Invite Email Sender
 * Usage: node send-invites.js guests.xlsx
 *
 * Spreadsheet format (row 1 = headers):
 *   Column A: Name   (e.g. "Sarah & Tom")
 *   Column B: Email  (e.g. sarah@example.com)
 */

const nodemailer = require('nodemailer');
const XLSX       = require('xlsx');
const path       = require('path');
const fs         = require('fs');

// ─────────────────────────────────────────────────────────────────────────────
//  CONFIG — edit these before running
// ─────────────────────────────────────────────────────────────────────────────
const CONFIG = {
  // Full URL where invite.html is hosted (GitHub Pages, Netlify, etc.)
  inviteBaseUrl: 'https://YOUR_SITE_URL/invite.html',

  smtp: {
    host: 'smtp.gmail.com',
    port: 587,
    secure: false,
    auth: {
      user: 'YOUR_GMAIL@gmail.com',
      pass: 'YOUR_GMAIL_APP_PASSWORD',  // Generate at myaccount.google.com > Security > App Passwords
    },
  },

  fromName:  'Luke & Ayten',
  fromEmail: 'YOUR_GMAIL@gmail.com',
  subject:   "You're Invited — Luke & Ayten's Wedding",

  // Delay between emails (ms) — avoids Gmail rate limits
  delayMs: 400,
};
// ─────────────────────────────────────────────────────────────────────────────

function readGuests(filePath) {
  const ext = path.extname(filePath).toLowerCase();

  if (ext === '.csv') {
    const lines = fs.readFileSync(filePath, 'utf-8').trim().split('\n');
    return lines.slice(1).map(line => {
      const [name, email] = line.split(',').map(s => s.trim().replace(/^"|"$/g, ''));
      return { name, email };
    }).filter(g => g.name && g.email && g.email.includes('@'));
  }

  // .xlsx / .xls
  const wb   = XLSX.readFile(filePath);
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  return rows.slice(1).map(row => ({
    name:  String(row[0] || '').trim(),
    email: String(row[1] || '').trim(),
  })).filter(g => g.name && g.email && g.email.includes('@'));
}

function buildInviteUrl(guestName) {
  return `${CONFIG.inviteBaseUrl}?guest=${encodeURIComponent(guestName)}`;
}

function buildEmail(guest) {
  const url = buildInviteUrl(guest.name);

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${CONFIG.subject}</title>
</head>
<body style="margin:0;padding:0;background:#0E0604;font-family:Georgia,'Times New Roman',serif;">

<table width="100%" cellpadding="0" cellspacing="0" style="background:#0E0604;padding:48px 16px;">
<tr><td align="center">

  <table width="520" cellpadding="0" cellspacing="0"
         style="background:#1C0C06;border:1px solid #8B6A2A;border-radius:3px;max-width:520px;width:100%;">

    <!-- Gold bar top -->
    <tr><td height="3" style="background:linear-gradient(90deg,#6B4A18,#C8A456,#D4B060,#C8A456,#6B4A18);"></td></tr>

    <!-- Names header -->
    <tr>
      <td align="center" style="padding:44px 40px 20px;">
        <p style="margin:0 0 14px;color:#C8A456;font-size:10px;letter-spacing:7px;text-transform:uppercase;">
          Together with their families
        </p>
        <p style="margin:0;color:#F5EDD8;font-size:30px;font-weight:normal;letter-spacing:1px;">
          Luke Charles Spiteri
        </p>
        <p style="margin:6px 0;color:#C8A456;font-size:34px;font-weight:normal;">&amp;</p>
        <p style="margin:0;color:#F5EDD8;font-size:30px;font-weight:normal;letter-spacing:1px;">
          Ayten Kotb Harby
        </p>
      </td>
    </tr>

    <!-- Divider -->
    <tr>
      <td align="center" style="padding:0 40px 0;">
        <table width="60" cellpadding="0" cellspacing="0">
          <tr><td height="1" style="background:#C8A456;"></td></tr>
        </table>
      </td>
    </tr>

    <!-- Personal greeting -->
    <tr>
      <td style="padding:28px 44px 8px;color:#D4B896;font-size:15px;font-style:italic;text-align:center;">
        Dear ${guest.name},
      </td>
    </tr>
    <tr>
      <td style="padding:8px 44px 24px;color:#A89878;font-size:14px;line-height:1.8;text-align:center;">
        We joyfully request the honour of your presence<br>
        as we celebrate our marriage.
      </td>
    </tr>

    <!-- Details block -->
    <tr>
      <td align="center" style="padding:0 44px 32px;">
        <table cellpadding="0" cellspacing="0">
          <tr>
            <td style="padding:0 24px 0 0;text-align:center;vertical-align:top;">
              <p style="margin:0 0 5px;color:#C8A456;font-size:9px;letter-spacing:4px;text-transform:uppercase;">Date</p>
              <p style="margin:0;color:#F5EDD8;font-size:16px;font-style:italic;">8 &bull; 8 &bull; 26</p>
            </td>
            <td width="1" style="background:#4A3A20;">&nbsp;</td>
            <td style="padding:0 0 0 24px;text-align:center;vertical-align:top;">
              <p style="margin:0 0 5px;color:#C8A456;font-size:9px;letter-spacing:4px;text-transform:uppercase;">Time</p>
              <p style="margin:0;color:#F5EDD8;font-size:16px;font-style:italic;">2:00 — 5:00 PM</p>
            </td>
          </tr>
          <tr><td colspan="3" height="20"></td></tr>
          <tr>
            <td colspan="3" style="text-align:center;">
              <p style="margin:0 0 5px;color:#C8A456;font-size:9px;letter-spacing:4px;text-transform:uppercase;">Venue</p>
              <p style="margin:0;color:#F5EDD8;font-size:17px;">Oakview Terrace</p>
              <p style="margin:4px 0 0;color:#8A7A60;font-size:13px;font-style:italic;">
                13256 Leslie St, Richmond Hill, ON L4E 1A2
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- CTA -->
    <tr>
      <td align="center" style="padding:0 40px 44px;">
        <a href="${url}"
           style="display:inline-block;padding:14px 44px;border:1px solid #C8A456;
                  color:#C8A456;font-family:Georgia,serif;font-size:11px;
                  letter-spacing:5px;text-transform:uppercase;text-decoration:none;
                  border-radius:2px;">
          Open Your Invitation
        </a>
        <p style="margin:20px 0 0;color:#5A4A30;font-size:11px;letter-spacing:2px;">
          RSVP BY 5 &bull; 15 &bull; 26
        </p>
      </td>
    </tr>

    <!-- Gold bar bottom -->
    <tr><td height="3" style="background:linear-gradient(90deg,#6B4A18,#C8A456,#D4B060,#C8A456,#6B4A18);"></td></tr>

  </table>

  <!-- Fallback link -->
  <p style="margin:16px 0 0;color:#3A2A18;font-size:11px;font-family:Georgia,serif;">
    If the button doesn't work:
    <a href="${url}" style="color:#6B4A18;">${url}</a>
  </p>

</td></tr>
</table>
</body>
</html>`;

  return {
    from:    `"${CONFIG.fromName}" <${CONFIG.fromEmail}>`,
    to:      guest.email,
    subject: CONFIG.subject,
    html,
    // Plain-text fallback
    text: `Dear ${guest.name},\n\nLuke Charles Spiteri & Ayten Kotb Harby request the honour of your presence at their wedding.\n\n8 • 8 • 26, 2:00–5:00 PM\nOakview Terrace, 13256 Leslie St, Richmond Hill, ON\n\nOpen your invitation: ${url}\n\nRSVP by 5 • 15 • 26`,
  };
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function main() {
  const file = process.argv[2] || 'guests.xlsx';

  if (!fs.existsSync(file)) {
    console.error(`\n  ✗  File not found: ${file}`);
    console.error('  Usage: node send-invites.js guests.xlsx\n');
    process.exit(1);
  }

  // Validate config
  if (CONFIG.inviteBaseUrl.includes('YOUR_SITE_URL')) {
    console.error('\n  ✗  Set CONFIG.inviteBaseUrl in send-invites.js before sending.\n');
    process.exit(1);
  }
  if (CONFIG.smtp.auth.user.includes('YOUR_GMAIL')) {
    console.error('\n  ✗  Set your Gmail credentials in CONFIG.smtp before sending.\n');
    process.exit(1);
  }

  const guests = readGuests(file);
  if (guests.length === 0) {
    console.error('\n  ✗  No guests found. Check your spreadsheet (Name in col A, Email in col B).\n');
    process.exit(1);
  }

  console.log(`\n  Wedding Invite Sender`);
  console.log(`  ${'─'.repeat(40)}`);
  console.log(`  Guests found : ${guests.length}`);
  console.log(`  From         : ${CONFIG.fromName} <${CONFIG.fromEmail}>`);
  console.log(`  Subject      : ${CONFIG.subject}`);
  console.log(`  Invite URL   : ${CONFIG.inviteBaseUrl}`);
  console.log(`  ${'─'.repeat(40)}\n`);

  const transporter = nodemailer.createTransport(CONFIG.smtp);

  try {
    await transporter.verify();
    console.log('  ✓  SMTP connection verified\n');
  } catch (err) {
    console.error(`  ✗  SMTP error: ${err.message}\n`);
    process.exit(1);
  }

  let sent = 0, failed = 0;

  for (const guest of guests) {
    try {
      await transporter.sendMail(buildEmail(guest));
      console.log(`  ✓  ${guest.name.padEnd(30)} → ${guest.email}`);
      sent++;
    } catch (err) {
      console.error(`  ✗  ${guest.name.padEnd(30)} → ${guest.email}  (${err.message})`);
      failed++;
    }
    await sleep(CONFIG.delayMs);
  }

  console.log(`\n  ${'─'.repeat(40)}`);
  console.log(`  Sent   : ${sent}`);
  if (failed > 0) console.log(`  Failed : ${failed}`);
  console.log(`  ${'─'.repeat(40)}\n`);
}

main().catch(err => {
  console.error('\n  Fatal:', err.message, '\n');
  process.exit(1);
});
