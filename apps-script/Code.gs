/**
 * RSVP receiver for the Ajay & Edna wedding invitation.
 * Bind this script to a Google Sheet (Extensions -> Apps Script), run setup() once,
 * then Deploy -> New deployment -> Web app (Execute as: Me, Who has access: Anyone).
 */
const RSVP_SHEET = 'RSVPs';
const SUMMARY_SHEET = 'Summary';
const HEADERS = ['Received (IST)', 'Full Name', 'Attendance', 'Number of Guests', 'Message', 'Device'];
const TZ = 'Asia/Kolkata';
const ATTEND = 'Happily Attending';
const DECLINE = 'Regretfully Decline';

/** Run once from the editor: creates both tabs, headers and live totals. */
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
  sm.getRange('A1').setFontWeight('bold').setFontSize(14);
  sm.getRange('A5:B5').setFontWeight('bold');
  sm.autoResizeColumn(1);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true });            // honeypot: bots fill it, people can't see it
    const name = clean_(d.name, 120);
    if (!name) return json_({ ok: false, error: 'name required' });
    const attend = d.attend === DECLINE ? DECLINE : ATTEND;
    const guests = Math.max(1, Math.min(20, parseInt(d.guests, 10) || 1));
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RSVP_SHEET);
    if (!sheet) return json_({ ok: false, error: 'run setup() first' });
    sheet.appendRow([new Date(), name, attend, guests, clean_(d.msg, 1000), clean_(d.device, 20)]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** Health check: open the /exec URL in a browser and you should see {"ok":true,...}. */
function doGet() {
  return json_({ ok: true, service: 'rsvp' });
}

/** Trim, cap length, and stop text being read as a spreadsheet formula. */
function clean_(v, max) {
  const s = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
