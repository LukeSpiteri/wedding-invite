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
  inviteBaseUrl: 'https://lukespiteri.github.io/wedding-invite/invite.html',

  smtp: {
    host: 'smtp.gmail.com',
    port: 587,
    secure: false,
    auth: {
      user: 'lukeayten@gmail.com',
      pass: 'mwsj icay yuaw phxz',
    },
  },

  fromName:  'Luke & Ayten',
  fromEmail: 'lukeayten@gmail.com',
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
<body style="margin:0;padding:0;background:#F5E6A3;font-family:Georgia,'Times New Roman',serif;">

<table width="100%" cellpadding="0" cellspacing="0" style="background:#F5E6A3;padding:48px 16px;">
<tr><td align="center">

  <table width="560" cellpadding="0" cellspacing="0"
         style="background:#FFFDF4;border:2px solid #C8A456;border-radius:6px;max-width:560px;width:100%;box-shadow:0 4px 24px rgba(107,29,42,0.12);">

    <!-- Gold & Maroon bar top -->
    <tr><td height="5" style="background:linear-gradient(90deg,#6B1D2A,#C8A456,#D4B060,#C8A456,#6B1D2A);"></td></tr>

    <!-- Decorative maroon accent strip -->
    <tr><td height="2" style="background:#6B1D2A;"></td></tr>

    <!-- L&A Logo header -->
    <tr>
      <td align="center" style="padding:36px 40px 12px;">
        <img src="https://lukespiteri.github.io/wedding-invite/logo.png"
             alt="L&A" width="90" height="90"
             style="width:90px;height:auto;">
      </td>
    </tr>

    <!-- Names header -->
    <tr>
      <td align="center" style="padding:8px 40px 20px;">
        <p style="margin:0;color:#6B1D2A;font-size:28px;font-weight:normal;letter-spacing:1px;font-style:italic;">
          Luke Charles Spiteri
        </p>
        <p style="margin:6px 0;color:#C8A456;font-size:32px;font-weight:normal;">&amp;</p>
        <p style="margin:0;color:#6B1D2A;font-size:28px;font-weight:normal;letter-spacing:1px;font-style:italic;">
          Ayten Kotb Harby
        </p>
      </td>
    </tr>

    <!-- Gold divider -->
    <tr>
      <td align="center" style="padding:4px 40px 8px;">
        <table width="100" cellpadding="0" cellspacing="0">
          <tr>
            <td width="38" height="1" style="background:linear-gradient(90deg,transparent,#C8A456);"></td>
            <td width="24" align="center" style="color:#C8A456;font-size:10px;line-height:1;">✦</td>
            <td width="38" height="1" style="background:linear-gradient(90deg,#C8A456,transparent);"></td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- Personal greeting -->
    <tr>
      <td style="padding:24px 48px 8px;color:#6B1D2A;font-size:19px;font-style:italic;text-align:center;">
        Dear ${guest.name},
      </td>
    </tr>

    <!-- Message paragraphs -->
    <tr>
      <td style="padding:8px 48px 8px;color:#5A4632;font-size:14px;line-height:1.9;text-align:center;">
        We would be absolutely honoured to have you join us for our wedding ceremony this summer. Please click the link below to view our formal invitation.
      </td>
    </tr>
    <tr>
      <td style="padding:12px 48px 8px;color:#5A4632;font-size:14px;line-height:1.9;text-align:center;">
        We are keeping this ceremony as an intimate gathering. Since our guest list is quite limited, we kindly ask that you RSVP at your earliest convenience to help us finalize our arrangements.
      </td>
    </tr>
    <tr>
      <td style="padding:12px 48px 8px;color:#5A4632;font-size:14px;line-height:1.9;text-align:center;">
        While we are thrilled to share this special milestone with you now, we also look forward to celebrating with everyone at our larger reception, which will be held at a later date in 2027. A separate invitation for the reception will follow down the road!
      </td>
    </tr>
    <tr>
      <td style="padding:12px 48px 20px;color:#5A4632;font-size:14px;line-height:1.9;text-align:center;">
        Upon receiving your RSVP, we will send a follow-up email closer to the date with more detailed information about the ceremony.
      </td>
    </tr>

    <!-- Excitement -->
    <tr>
      <td style="padding:8px 48px 16px;color:#6B1D2A;font-size:16px;font-style:italic;text-align:center;font-weight:bold;">
        We are so excited and hope you can join us!
      </td>
    </tr>

    <!-- Gold divider -->
    <tr>
      <td align="center" style="padding:0 40px 24px;">
        <table width="100" cellpadding="0" cellspacing="0">
          <tr>
            <td width="38" height="1" style="background:linear-gradient(90deg,transparent,#C8A456);"></td>
            <td width="24" align="center" style="color:#C8A456;font-size:10px;line-height:1;">✦</td>
            <td width="38" height="1" style="background:linear-gradient(90deg,#C8A456,transparent);"></td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- Invitation CTA -->
    <tr>
      <td align="center" style="padding:0 40px 20px;">
        <a href="${url}"
           style="display:inline-block;padding:14px 44px;border:2px solid #6B1D2A;
                  color:#6B1D2A;font-family:Georgia,serif;font-size:11px;
                  letter-spacing:5px;text-transform:uppercase;text-decoration:none;
                  border-radius:3px;">
          Open Your Invitation
        </a>
      </td>
    </tr>

    <!-- RSVP CTA -->
    <tr>
      <td align="center" style="padding:0 40px 32px;">
        <a href="https://forms.gle/SmGZT9ihgRDvha988"
           style="display:inline-block;padding:14px 44px;background:#6B1D2A;
                  color:#F5E6A3;font-family:Georgia,serif;font-size:11px;
                  letter-spacing:5px;text-transform:uppercase;text-decoration:none;
                  border-radius:3px;font-weight:bold;">
          RSVP Here
        </a>
      </td>
    </tr>

    <!-- Sign-off -->
    <tr>
      <td style="padding:0 48px 8px;color:#5A4632;font-size:14px;text-align:center;letter-spacing:1px;">
        With love,
      </td>
    </tr>
    <tr>
      <td style="padding:0 48px 28px;color:#6B1D2A;font-size:22px;font-style:italic;text-align:center;letter-spacing:1px;">
        Luke and Ayten
      </td>
    </tr>

    <!-- L&A Logo footer -->
    <tr>
      <td align="center" style="padding:0 40px 32px;">
        <img src="https://lukespiteri.github.io/wedding-invite/logo.png"
             alt="L&A" width="100" height="100"
             style="width:100px;height:auto;opacity:0.9;">
      </td>
    </tr>

    <!-- Maroon accent strip -->
    <tr><td height="2" style="background:#6B1D2A;"></td></tr>

    <!-- Gold & Maroon bar bottom -->
    <tr><td height="5" style="background:linear-gradient(90deg,#6B1D2A,#C8A456,#D4B060,#C8A456,#6B1D2A);"></td></tr>

  </table>

  <!-- Fallback link -->
  <p style="margin:16px 0 0;color:#6B1D2A;font-size:11px;font-family:Georgia,serif;">
    If the button doesn't work:
    <a href="${url}" style="color:#C8A456;">${url}</a>
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
    text: `Dear ${guest.name},\n\nWe would be absolutely honoured to have you join us for our wedding ceremony this summer. Please click the link below to view our formal invitation.\n\nWe are keeping this ceremony as an intimate gathering. Since our guest list is quite limited, we kindly ask that you RSVP at your earliest convenience to help us finalize our arrangements.\n\nWhile we are thrilled to share this special milestone with you now, we also look forward to celebrating with everyone at our larger reception, which will be held at a later date in 2027. A separate invitation for the reception will follow down the road!\n\nUpon receiving your RSVP, we will send a follow-up email closer to the date with more detailed information about the ceremony.\n\nWe are so excited and hope you can join us!\n\nClick the link to experience the invitation: ${url}\n\nRSVP here: https://forms.gle/SmGZT9ihgRDvha988\n\nWith love,\nLuke and Ayten`,
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
