/**
 * RSVP receiver for the Ajay & Edna wedding invitation.
 * Bind this script to a Google Sheet (Extensions -> Apps Script), run setup() once,
 * then Deploy -> New deployment -> Web app (Execute as: Me, Who has access: Anyone).
 * After changing this file: run setup() again, then Deploy -> Manage deployments -> Edit -> New version.
 */
const RSVP_SHEET = 'RSVPs';
const SUMMARY_SHEET = 'Summary';
const HEADERS = ['Received (IST)', 'Full Name', 'Attendance', 'Number of Guests', 'Message', 'Device', 'Email', 'Confirmation email'];
const TZ = 'Asia/Kolkata';
const ATTEND = 'Happily Attending';
const DECLINE = 'Regretfully Decline';

const SITE_URL = 'https://ajayjc47.github.io/';
const SENDER_NAME = 'Ajay & Edna';
const MAIL_COOLDOWN_SECONDS = 21600;   // at most one confirmation per email address every 6 hours

/** Run from the editor: creates/refreshes both tabs, headers and live totals, and asks for permissions. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);                       // timestamps display in IST
  const rs = ss.getSheetByName(RSVP_SHEET) || ss.insertSheet(RSVP_SHEET);
  rs.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  rs.getRange('A2:A').setNumberFormat('dd mmm yyyy, h:mm:ss am/pm');
  rs.setFrozenRows(1);
  rs.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');

  const sm = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET, 0);
  const r = `${RSVP_SHEET}!`;
  sm.clear();
  sm.getRange('A1:B7').setValues([
    ['RSVP summary', ''],
    ['Total replies', `=COUNTA(${r}B2:B)`],
    ['Replies attending', `=COUNTIF(${r}C2:C,"${ATTEND}")`],
    ['Replies declined', `=COUNTIF(${r}C2:C,"${DECLINE}")`],
    ['Total guests attending', `=SUMIF(${r}C2:C,"${ATTEND}",${r}D2:D)`],
    ['Distinct names', `=COUNTUNIQUE(${r}B2:B)`],
    ['Possible duplicates (replies - distinct names)', '=B2-B6'],
  ]);
  MailApp.getRemainingDailyQuota();                   // asks for email permission during setup
  sm.getRange('A1').setFontWeight('bold').setFontSize(14);
  sm.getRange('A5:B5').setFontWeight('bold');
  sm.autoResizeColumn(1);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let row, rsvp;
  try {
    lock.waitLock(10000);
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true });            // honeypot: bots fill it, people can't see it
    const name = clean_(d.name, 120);
    if (!name) return json_({ ok: false, error: 'name required' });
    const attend = d.attend === DECLINE ? DECLINE : ATTEND;
    const guests = Math.max(1, Math.min(20, parseInt(d.guests, 10) || 1));
    const email = validEmail_(d.email);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RSVP_SHEET);
    if (!sheet) return json_({ ok: false, error: 'run setup() first' });
    sheet.appendRow([new Date(), name, attend, guests, clean_(d.msg, 1000), clean_(d.device, 20),
                     clean_(email, 254), email ? 'pending' : '']);
    row = sheet.getLastRow();
    rsvp = { sheet, name: String(d.name).trim().slice(0, 120), attend, guests, email };
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
  // The RSVP is saved. The email is a bonus: any problem here never fails the RSVP.
  if (rsvp.email) {
    let status;
    try { status = sendConfirmation_(rsvp); } catch (err) { status = 'error: ' + String(err).slice(0, 200); }
    try { rsvp.sheet.getRange(row, 8).setValue(status); } catch (_) {}
  }
  return json_({ ok: true });
}

/** Health check: open the /exec URL in a browser and you should see {"ok":true,...}. */
function doGet() {
  return json_({ ok: true, service: 'rsvp' });
}

function sendConfirmation_(r) {
  const cache = CacheService.getScriptCache();
  const key = 'mail:' + r.email.toLowerCase();
  if (cache.get(key)) return 'skipped: already sent recently';
  if (MailApp.getRemainingDailyQuota() < 1) return 'skipped: daily email limit reached';
  MailApp.sendEmail({
    to: r.email,
    subject: "Your RSVP · Ajay & Edna's Wedding · 27 December 2026",
    name: SENDER_NAME,
    htmlBody: emailHtml_(r),
    body: emailText_(r),
  });
  cache.put(key, '1', MAIL_COOLDOWN_SECONDS);
  return 'sent ' + Utilities.formatDate(new Date(), TZ, 'dd MMM, h:mm a');
}

/* ---------------- email content (matches email-preview.html) ---------------- */

const CAL_URL = 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=Wedding%20of%20Ajay%20%26%20Edna&dates=20261227T150000/20261227T170000&location=Marth%20Mariam%20Forane%20Syro-Malabar%20Catholic%20Church%2C%20Thuruthi%2C%20Kerala&details=Wedding%20of%20Dr.%20Ajay%20Choondal%20%26%20Dr.%20Edna%20Chinnu%20Kurian';
const MAP_URL = 'https://www.google.com/maps/search/?api=1&query=Marth%20Mariam%20Forane%20Syro-Malabar%20Catholic%20Church%2C%20Thuruthi%2C%20Kerala';

function emailHtml_(r) {
  const yes = r.attend === ATTEND;
  const SANS = "font-family:Inter,Arial,sans-serif;";
  const SERIF = "font-family:'Cormorant Garamond',Georgia,serif;";
  const LBL = `${SANS}font-size:11px;letter-spacing:.26em;text-transform:uppercase;color:#C8A96B;`;
  const rule = '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr><td style="width:120px;height:1px;background:#C8A96B;font-size:0;line-height:0;">&nbsp;</td></tr></table>';
  const greeting = yes
    ? 'We are delighted that you will be celebrating with us. Here is the reply we received:'
    : 'Thank you for letting us know. We will miss you, and we are grateful for your wishes.';
  const details = yes ? `
      <tr><td style="padding:24px 36px 0;">
        <p style="margin:0;${SERIF}font-style:italic;font-size:24px;color:#2E2E2E;">Holy Matrimony</p>
        <p style="margin:6px 0 0;${SANS}font-weight:300;font-size:15px;line-height:1.7;color:#2E2E2E;">3:00 PM &middot; Marth Mariam Forane Syro-Malabar Catholic Church, Thuruthi, Kerala</p>
        <p style="margin:18px 0 0;${SERIF}font-style:italic;font-size:24px;color:#2E2E2E;">Reception</p>
        <p style="margin:6px 0 0;${SANS}font-weight:300;font-size:15px;line-height:1.7;color:#2E2E2E;">6:00 PM &middot; Church Parish Hall, Thuruthi, Kerala</p>
      </td></tr>
      <tr><td align="center" style="padding:30px 24px 8px;">
        <a href="${CAL_URL}" style="display:inline-block;margin:6px;padding:14px 24px;background:#6F8470;border:1px solid #6F8470;color:#ffffff;${SANS}font-size:12px;letter-spacing:.24em;text-transform:uppercase;text-decoration:none;">Add to Calendar</a>
        <a href="${MAP_URL}" style="display:inline-block;margin:6px;padding:14px 24px;border:1px solid #C8A96B;color:#2E2E2E;${SANS}font-size:12px;letter-spacing:.24em;text-transform:uppercase;text-decoration:none;">View Location</a>
      </td></tr>` : '';
  return `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#F8EDE9;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#F8EDE9;">
<tr><td align="center" style="padding:28px 12px;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;background:#FFFDF9;border:1px solid #C8A96B;">
    <tr><td style="padding:6px;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E4D3AE;">
      <tr><td align="center" style="padding:40px 28px 8px;">
        <p style="margin:0;${LBL}letter-spacing:.38em;">Thank you for your reply</p>
        <h1 style="margin:14px 0 0;font-family:'Playfair Display','Cormorant Garamond',Georgia,serif;font-style:italic;font-weight:400;font-size:38px;line-height:1.2;color:#2E2E2E;">Ajay &amp; Edna</h1>
        <p style="margin:10px 0 22px;${SERIF}font-style:italic;font-size:18px;color:#6F8470;">Sunday, 27 December 2026 &middot; Thuruthi, Kerala</p>
        ${rule}
      </td></tr>
      <tr><td style="padding:22px 36px 6px;${SANS}font-weight:300;font-size:16px;line-height:1.75;color:#2E2E2E;">
        Dear ${esc_(r.name)},<br><br>${greeting}
      </td></tr>
      <tr><td style="padding:10px 36px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#F6E6E1;border:1px solid #E9D3CE;">
          <tr><td style="padding:18px 22px;${SANS}font-weight:300;font-size:15px;line-height:2;color:#2E2E2E;">
            <span style="${LBL}">Attendance</span><br>${r.attend}
            ${yes ? `<br><span style="${LBL}">Number of guests</span><br>${r.guests}` : ''}
          </td></tr>
        </table>
      </td></tr>${details}
      <tr><td align="center" style="padding:22px 36px 40px;">
        <div style="margin:0 0 18px;">${rule}</div>
        <p style="margin:0;${SANS}font-weight:300;font-size:13px;line-height:1.7;color:#6b6b6b;">Plans changed? Simply submit the RSVP form again on the <a href="${SITE_URL}#rsvp" style="color:#6F8470;">invitation page</a>.</p>
        <p style="margin:16px 0 0;${SERIF}font-style:italic;font-size:20px;color:#2E2E2E;">With love, Ajay &amp; Edna</p>
      </td></tr>
    </table>
    </td></tr>
  </table>
</td></tr>
</table>
</body></html>`;
}

function emailText_(r) {
  const yes = r.attend === ATTEND;
  return [
    `Dear ${r.name},`,
    '',
    yes ? 'We are delighted that you will be celebrating with us. Here is the reply we received:'
        : 'Thank you for letting us know. We will miss you, and we are grateful for your wishes.',
    '',
    `Attendance: ${r.attend}` + (yes ? `\nNumber of guests: ${r.guests}` : ''),
    '',
    ...(yes ? [
      'Holy Matrimony: 3:00 PM, Marth Mariam Forane Syro-Malabar Catholic Church, Thuruthi, Kerala',
      'Reception: 6:00 PM, Church Parish Hall, Thuruthi, Kerala',
      '',
      'Add to calendar: ' + CAL_URL,
      'Location: ' + MAP_URL,
      '',
    ] : []),
    `Plans changed? Submit the RSVP form again: ${SITE_URL}#rsvp`,
    '',
    'With love, Ajay & Edna',
  ].join('\n');
}

/* ---------------- helpers ---------------- */

function validEmail_(v) {
  const s = String(v == null ? '' : v).trim().slice(0, 254);
  return /^[^\s@<>()",;:]+@[^\s@<>()",;:]+\.[A-Za-z]{2,}$/.test(s) ? s : '';
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/** Trim, cap length, and stop text being read as a spreadsheet formula. */
function clean_(v, max) {
  const s = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
