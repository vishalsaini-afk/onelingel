/**
 * Lingel — Installation Manual link (small stand-alone web app)
 *
 * What it does: Team Availability asks it for an order's installation manual and
 * gets the PDF back, so the 📘 arrow opens the manual directly — also for people
 * who have no access to the Drive folder.
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
