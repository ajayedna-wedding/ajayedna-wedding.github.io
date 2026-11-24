# Ajay & Edna — Wedding Invitation

Single-page invitation (`index.html`) hosted on GitHub Pages. RSVPs are sent to a Google Sheet through a small Google Apps Script web app (`apps-script/Code.gs`), so the couple gets a live headcount.

## How RSVPs work

1. A guest fills in the form and taps **Submit RSVP**.
2. The page sends the name, attendance, number of guests and message to the Apps Script `/exec` URL.
3. The script adds a row to the **RSVPs** tab and replies `{"ok":true}`.
4. The guest sees "Your RSVP has been received ✓". If saving fails (no signal, script not deployed), the form stays filled in with an error message so they can tap **Submit** again. The Google Sheet is the only place RSVPs go.

The **Summary** tab shows total replies, attending/declined, **total guests attending**, and possible duplicates.

## One-time Google setup (about 10 minutes)

1. In the couple's Google account, create a new Google Sheet (e.g. "Wedding RSVPs").
2. **Extensions → Apps Script.** Replace the contents of `Code.gs` with `apps-script/Code.gs` from this repo, then save.
3. Choose `setup` in the function dropdown and click **Run**. Approve the permission prompt. The sheet now has *Summary* and *RSVPs* tabs.
4. **Deploy → New deployment →** type **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone** (not "Anyone with a Google account": guests must not need to sign in)
5. Copy the **Web app URL** (ends in `/exec`). Open it in a browser and check that it shows `{"ok":true,"service":"rsvp"}`.
6. In `index.html`, set `rsvpEndpoint: "<that URL>"` in the `CFG` block, then commit and push.

If you change `Code.gs` later, use **Deploy → Manage deployments → Edit → New version**. This keeps the same URL, so the page doesn't need changing.

## Confirmation emails

If a guest enters an email address (optional), the script emails them a confirmation styled like the invitation (`email-preview.html` shows it). Guests who decline get a shorter version.

- Sent from the Google account that owns the sheet, with the display name "Ajay & Edna".
- **Confirmation email** column: `sent <time>`, `skipped: already sent recently` (one per address every 6 hours; run `resetEmailLimit` in the editor to clear it while testing), `skipped: daily email limit reached` (about 100 a day on a personal Gmail), or `error: …`. The RSVP itself is always saved.
- After updating `Code.gs`: run `setup` once (approve the new "send email" permission), then **Deploy → Manage deployments → Edit → Version: New version → Deploy**. The `/exec` URL stays the same.

## Viewing the sheet

- **Computer:** go to sheets.google.com (or drive.google.com) signed in to the account that owns the sheet, then open "Wedding RSVPs".
- **Phone:** install the Google Sheets app, sign in with the same account, and open it. New RSVPs appear within seconds. Pull down to refresh.
- **Share with family:** in the sheet, tap **Share**, add their Gmail addresses, and choose **Viewer** (read-only) or **Editor**. RSVPs keep working, because the script runs as the owner. Don't use "Anyone with the link": the sheet holds guests' names and messages.
- The **Received (IST)** column shows Indian time (the sheet's time zone is set to Asia/Kolkata by `setup`).

## Notes

- Anyone with the page link can submit. A hidden "honeypot" field filters out simple bots.
- A guest who submits twice appears twice. The *Device* column helps you spot repeat submissions from the same phone. Delete the extra row.
- Links (maps, calendar) are plain links, so they work in WhatsApp's in-app browser and in Safari or Chrome.

## Changelog

- 2026-11-24 — `resetEmailLimit()` editor function to clear the 6-hour email limit while testing.
- 2026-11-24 — Reception time set to 6:00 PM (page, calendar event 6–9 PM, confirmation email, preview).
- 2026-11-24 — Optional email field; themed confirmation email from the Apps Script, with a 6-hour per-address limit.
- 2026-11-24 — RSVPs now go to the sheet owned by the wedding Gmail account.
- 2026-11-24 — Initial version: invitation page, RSVP to Google Sheet via Apps Script (timestamps in IST), email preview for sign-off.
