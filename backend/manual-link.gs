/**
 * Lingel — Installation Manual link (small stand-alone web app)
 *
 * What it does:
 *  1. Team Availability asks it for an order's installation manual and gets the PDF
 *     back, so the 📘 arrow opens the manual directly — also for people who have no
 *     access to the Drive folder.
 *  2. The voice feedback page (manual-feedback.html) sends each feedback here and it
 *     is written into the "Form Responses (voice page)" tab of the Installation Manual Log sheet.
 *
 * Already deployed an older copy? Paste this over it, Save, then
 * Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy (the link stays the same).
 *
 * One-time setup:
 * 1. https://script.google.com → New project → paste this whole file → Save
 * 2. Deploy → New deployment → type "Web app"
 *      Execute as:      Me
 *      Who has access:  Anyone
 *    → Deploy → approve the permission prompt → copy the Web app URL (ends in /exec)
 * 3. Send that URL to Claude (it goes into team-availability.html as MANUAL_API)
 *
 * Test in a browser:  <your /exec URL>?action=ping
 *                     <your /exec URL>?action=manual&offer=85194/4&meta=1
 */

var FOLDER_ID = '1jlBQl_CeFinA5G1jCkMvjnqBp6ZKopdo';   // "Lingel SiteLog Reports" (one sub-folder per offer no.)
var PREFIX = 'Installation_Manual_';
var LOG_SHEET_ID = '1rmB_G-5mi3EA9jZvh3SK0F8Qk1ly48M1KGBDLjJTLDs';   // "Installation Manual Log"
var FEEDBACK_TAB = 'Form Responses (voice page)';   // name starts with "Form Responses" so the feedback review reads it

/* Feedback from the voice feedback page (manual-feedback.html) → one row in the "Form Responses (voice page)" tab */
function doPost(e) {
  var p = (e && e.parameter) || {};
  var res;
  try {
    if (p.action === 'feedback') res = feedback_(p);
    else res = { ok: false, error: 'unknown action' };
  } catch (err) {
    res = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function feedback_(p) {
  var cut = function (v, n) { return String(v == null ? '' : v).replace(/^[=+\-@]/, "'$&").slice(0, n); };
  var offer = cut(p.offer, 40);
  if (!offer) return { ok: false, error: 'Offer number missing' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(LOG_SHEET_ID);
    var sh = ss.getSheetByName(FEEDBACK_TAB);
    if (!sh) {
      sh = ss.insertSheet(FEEDBACK_TAB);
      sh.appendRow(['Timestamp', 'Offer No', 'Name', 'Role', 'Rating (1-5)', 'What was wrong', 'Position', 'What it should be', 'Language', 'Sent from']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), offer, cut(p.name, 80), cut(p.role, 40), cut(p.rating, 2), cut(p.wrong, 400), cut(p.position, 80), cut(p.detail, 2000), cut(p.lang, 10), cut(p.ua, 120)]);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var res;
  try {
    if (p.action === 'ping') res = { ok: true, app: 'manual-link' };
    else if (p.action === 'manual') res = manual_(p.offer, p.meta);
    else res = { ok: false, error: 'unknown action' };
  } catch (err) {
    res = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

/* offer "85194/4" (also 85194-4, 85194 4, or just 85194) → the manual PDF of that order */
function manual_(offer, metaOnly) {
  var m = String(offer || '').trim().match(/^(\d{4,7})\s*(?:[\/\-_ ]\s*(\d{1,4}))?/);
  if (!m) return { ok: false, error: 'Offer number not understood' };
  var base = m[1], sub = m[2] ? String(+m[2]) : '';

  var found = [];
  var folders = DriveApp.getFolderById(FOLDER_ID).getFoldersByName(base);
  while (folders.hasNext()) {
    var files = folders.next().getFiles();
    while (files.hasNext()) {
      var f = files.next(), n = f.getName();
      if (n.indexOf(PREFIX) === 0 && /\.pdf$/i.test(n)) found.push(f);
    }
  }
  if (!found.length) return { ok: false, notFound: true, error: 'No installation manual yet for ' + base };

  var want = PREFIX + base + (sub ? '-' + sub : '');
  var pick = null;
  for (var i = 0; i < found.length; i++) {
    var nm = found[i].getName().replace(/\.pdf$/i, '');
    if (sub && nm === want) { pick = found[i]; break; }
  }
  if (!pick && !sub) pick = newest_(found);   // only the order no. given → newest manual
  if (!pick) {
    return { ok: false, notFound: true,
      error: 'No manual for ' + base + '/' + sub + ' yet',
      others: found.map(function (x) { return x.getName(); }) };
  }

  var out = { ok: true, name: pick.getName(), size: pick.getSize(), updated: pick.getLastUpdated().toISOString() };
  if (!metaOnly) out.data = Utilities.base64Encode(pick.getBlob().getBytes());
  return out;
}

function newest_(list) {
  var best = list[0];
  for (var i = 1; i < list.length; i++) if (list[i].getLastUpdated() > best.getLastUpdated()) best = list[i];
  return best;
}
