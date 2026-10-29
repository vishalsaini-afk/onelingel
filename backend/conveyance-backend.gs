/**
 * LINGEL — Plan-linked conveyance backend  (standalone Apps Script web app)
 * ---------------------------------------------------------------------------
 * Rule: no plan → no conveyance.
 *       Site Visit  → approved when a Site Inspection report OR a Handover form exists (logVisit, source SiteLog / Handover).
 *       Measurement → approved automatically, no report needed; km typed by the supervisor (MeasureKm tab).
 *
 * SETUP (once):
 *  1. Create a new Google Sheet "Lingel Conveyance" → Extensions → Apps Script.
 *  2. Paste this whole file as Code.gs. Run setup() once and allow access.
 *  3. Project Settings → Script properties → add HR_KEY = <a password for HR>.
 *  4. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 *  5. Copy the /exec URL into the apps (they ask for it once).
 *
 * Tabs (created by setup): Settings, Supervisors, Plans, SiteVisits, Extras,
 *                          Pins, Decisions, Paid, MeasureKm, Ledger (frozen at month close, for accounts)
 */

var TZ = 'Asia/Kolkata';
var TABS = {
  Settings:    ['Key', 'Value', 'Meaning'],
  Supervisors: ['Name', 'Base address', 'Base lat', 'Base lng', 'Vehicle', 'Active'],
  Plans:       ['Id', 'Date', 'Supervisor', 'Offer', 'Project', 'City', 'Task', 'Salesperson', 'First At', 'Updated At', 'Deleted'],
  SiteVisits:  ['Created At', 'Date', 'Offer', 'Supervisor', 'Lat', 'Lng', 'Accuracy m', 'Client', 'Positions', 'Source'],
  Extras:      ['Id', 'Date', 'Supervisor', 'Offer', 'Type', 'Amount', 'Reason', 'Bill URL', 'Status', 'HR Note', 'Added At'],
  Pins:        ['Offer', 'Lat', 'Lng', 'Source', 'Set At'],
  Decisions:   ['Date', 'Supervisor', 'Status', 'Amount', 'Note', 'By', 'At'],
  Paid:        ['Month', 'Supervisor', 'Amount', 'Paid At', 'By'],
  MeasureKm:   ['Date', 'Supervisor', 'Offer', 'Km', 'Note', 'Updated At'],
  Ledger:      ['Month', 'Date', 'Supervisor', 'Vehicle', 'Sites', 'Route', 'km', 'Rate', 'Travel', 'Bills', 'Total', 'Status', 'Approved By', 'Closed At']
};
var DEFAULTS = [
  ['rate_bike', 4, '₹ per km, own bike'],
  ['rate_car', 9, '₹ per km, own car'],
  ['rate_public', 4, '₹ per km cap when using bus/auto/cab (actual fare via Extras)'],
  ['daily_cap', 600, '₹ maximum regular travel per day'],
  ['cutoff_hour', 21, 'Plan for day D must be made before this hour on D-1 (24h)'],
  ['gps_radius_m', 500, 'Report GPS must be within this many metres of the site pin'],
  ['late_report_hour', 12, 'Report must be created on D (synced by this hour on D+1)'],
  ['extra_flag_amount', 300, 'Any single extra above this ₹ goes to HR review']
];

/* ============================ ENTRY POINTS ============================ */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    var a = p.action || 'ping';
    if (a === 'ping') return out({ ok: true, app: 'lingel-conveyance', at: new Date().toISOString() });
    if (a === 'settings') return out({ ok: true, settings: settings_(), supervisors: supervisors_().map(publicSup_) });
    if (a === 'my') return out(myMonth_(p.sup, p.month));
    if (a === 'hr') { hrGuard_(p.key); return out(hrMonth_(p.month)); }
    if (a === 'hrSettings') { hrGuard_(p.key); return out({ ok: true, settings: settings_(), supervisors: supervisors_() }); }
    return out({ ok: false, error: 'unknown action ' + a });
  } catch (err) { return out({ ok: false, error: String(err.message || err) }); }
}

function doPost(e) {
  var p = readParams_(e);
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var a = p.action || '';
    if (a === 'planUpsert') return out(planUpsert_(p));
    if (a === 'planBulk') return out(planBulk_(p));
    if (a === 'planDelete') return out(planDelete_(p));
    if (a === 'logVisit') return out(logVisit_(p));
    if (a === 'addExtra') return out(addExtra_(p));
    if (a === 'deleteExtra') return out(deleteExtra_(p));
    if (a === 'setMeasureKm') return out(setMeasureKm_(p));
    if (a === 'hrDecide') { hrGuard_(p.key); return out(hrDecide_(p)); }
    if (a === 'hrExtra') { hrGuard_(p.key); return out(hrExtra_(p)); }
    if (a === 'hrSaveSettings') { hrGuard_(p.key); return out(hrSaveSettings_(p)); }
    if (a === 'hrSaveSupervisor') { hrGuard_(p.key); return out(hrSaveSupervisor_(p)); }
    if (a === 'hrPin') { hrGuard_(p.key); return out(hrPin_(p)); }
    if (a === 'hrPaid') { hrGuard_(p.key); return out(hrPaid_(p)); }
    if (a === 'hrCloseMonth') { hrGuard_(p.key); return out(hrCloseMonth_(p)); }
    return out({ ok: false, error: 'unknown action ' + a });
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) });
  } finally { try { lock.releaseLock(); } catch (x) {} }
}

function setup() { Object.keys(TABS).forEach(function (n) { sheet_(n); }); settings_(); return 'ok'; }

/* ============================ HELPERS ============================ */
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function readParams_(e) {
  var p = {};
  if (e && e.parameter) Object.keys(e.parameter).forEach(function (k) { p[k] = e.parameter[k]; });
  if (e && e.postData && e.postData.contents && !Object.keys(p).length) {
    var c = e.postData.contents;
    try { var j = JSON.parse(c); Object.keys(j).forEach(function (k) { p[k] = j[k]; }); }
    catch (x) { c.split('&').forEach(function (kv) { var i = kv.indexOf('='); if (i > 0) p[decodeURIComponent(kv.slice(0, i))] = decodeURIComponent(kv.slice(i + 1).replace(/\+/g, ' ')); }); }
  }
  return p;
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    // plain text everywhere: dates stay "2026-10-28" (no timezone shifts), numbers are parsed by num_()
    sh.getRange(1, 1, sh.getMaxRows(), TABS[name].length).setNumberFormat('@');
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold'); sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(name) {
  var sh = sheet_(name), n = sh.getLastRow();
  if (n < 2) return [];
  var head = TABS[name], vals = sh.getRange(2, 1, n - 1, head.length).getValues();
  return vals.map(function (r, i) { var o = { _row: i + 2 }; head.forEach(function (h, j) { o[h] = r[j]; }); return o; });
}
function append_(name, obj) { var head = TABS[name]; sheet_(name).appendRow(head.map(function (h) { return obj[h] === undefined ? '' : obj[h]; })); }
function update_(name, rowNo, obj) { var head = TABS[name]; sheet_(name).getRange(rowNo, 1, 1, head.length).setValues([head.map(function (h) { return obj[h] === undefined ? '' : obj[h]; })]); }
function norm_(s) { return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase(); }
function cleanOffer_(s) { return String(s || '').trim().replace(/^#/, '').replace(/\s+/g, '').toUpperCase(); }
function dayStr_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = String(v); var m = s.match(/^(\d{4}-\d{2}-\d{2})/); if (m) return m[1];
  var d = new Date(s); return isNaN(d) ? s : Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function addDays_(iso, n) { var p = iso.split('-').map(Number); var d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)); return d.toISOString().slice(0, 10); }
/* An IST wall-clock moment as epoch ms */
function istMs_(iso, hour) { var p = iso.split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2], hour, 0, 0) - 330 * 60000; }
function toMs_(v) { if (!v) return NaN; if (Object.prototype.toString.call(v) === '[object Date]') return v.getTime(); return Date.parse(String(v)); }
function num_(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
function hrGuard_(key) {
  var k = PropertiesService.getScriptProperties().getProperty('HR_KEY');
  if (!k) throw new Error('HR_KEY is not set in Script properties');
  if (String(key || '') !== String(k)) throw new Error('Wrong HR key');
}
function settings_() {
  var rs = rows_('Settings'), have = {}, s = {};
  rs.forEach(function (r) { have[r.Key] = true; s[r.Key] = r.Value; });
  DEFAULTS.forEach(function (d) { if (!have[d[0]]) { append_('Settings', { Key: d[0], Value: d[1], Meaning: d[2] }); s[d[0]] = d[1]; } });
  Object.keys(s).forEach(function (k) { if (k !== 'note') s[k] = num_(s[k]); });
  return s;
}
function supervisors_() {
  return rows_('Supervisors').filter(function (r) { return String(r.Name || '').trim(); }).map(function (r) {
    return { name: String(r.Name).trim(), base: String(r['Base address'] || ''), lat: num_(r['Base lat']), lng: num_(r['Base lng']),
             vehicle: String(r.Vehicle || 'bike').toLowerCase(), active: String(r.Active || 'yes').toLowerCase() !== 'no', _row: r._row };
  });
}
function publicSup_(s) { return { name: s.name, vehicle: s.vehicle, active: s.active, hasBase: !!(s.lat && s.lng) }; }
function supByName_(name) { var n = norm_(name); var all = supervisors_(); for (var i = 0; i < all.length; i++) if (norm_(all[i].name) === n) return all[i]; return null; }
function metres_(aLat, aLng, bLat, bLng) {
  var R = 6371000, toR = Math.PI / 180, dLat = (bLat - aLat) * toR, dLng = (bLng - aLng) * toR;
  var h = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(aLat * toR) * Math.cos(bLat * toR) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return 2 * R * Math.asin(Math.sqrt(h));
}

/* ============================ PLANS (mirrored from Team Availability) ============================ */
function planRowFrom_(p, existing) {
  var now = new Date().toISOString();
  return {
    Id: p.id, Date: dayStr_(p.date), Supervisor: String(p.supervisor || '').trim(), Offer: cleanOffer_(p.offer),
    Project: p.project || '', City: p.city || '', Task: p.task || '', Salesperson: p.salesperson || '',
    'First At': existing ? existing['First At'] : (p.at || now), 'Updated At': now, Deleted: ''
  };
}
function planUpsert_(p) {
  if (!p.id || !/^[sl]\|/.test(String(p.id))) return { ok: true, skipped: 'not a supervisor plan' };
  var all = rows_('Plans'), hit = null;
  for (var i = 0; i < all.length; i++) if (String(all[i].Id) === String(p.id)) { hit = all[i]; break; }
  var row = planRowFrom_(p, hit);
  if (hit) {
    // a plan that changes its Offer No after the cut-off is treated as planned late
    if (cleanOffer_(hit.Offer) !== row.Offer && Date.now() > istMs_(addDays_(row.Date, -1), settings_().cutoff_hour)) row['First At'] = row['Updated At'];
    update_('Plans', hit._row, row);
  } else append_('Plans', row);
  return { ok: true, id: p.id };
}
function planBulk_(p) {
  var list = []; try { list = JSON.parse(p.plans || '[]'); } catch (x) {}
  var all = rows_('Plans'), byId = {}; all.forEach(function (r) { byId[String(r.Id)] = r; });
  var added = 0;
  list.forEach(function (e) { if (!e || !e.id || !/^[sl]\|/.test(String(e.id)) || byId[String(e.id)]) return; append_('Plans', planRowFrom_(e, null)); added++; });
  return { ok: true, added: added };
}
function planDelete_(p) {
  var all = rows_('Plans');
  for (var i = 0; i < all.length; i++) if (String(all[i].Id) === String(p.id)) { all[i].Deleted = new Date().toISOString(); update_('Plans', all[i]._row, all[i]); }
  return { ok: true };
}

/* ============================ VISITS (from SiteLog) ============================ */
function logVisit_(p) {
  var offer = cleanOffer_(p.offer), sup = String(p.supervisor || '').trim(), date = dayStr_(p.date || p.createdAt);
  if (!offer || !sup) return { ok: false, error: 'offer and supervisor are required' };
  var created = p.createdAt || new Date().toISOString();
  // same report generated again the same day = one visit (keep the first, fill GPS if it was missing)
  var all = rows_('SiteVisits');
  for (var i = 0; i < all.length; i++) {
    var r = all[i];
    if (cleanOffer_(r.Offer) === offer && norm_(r.Supervisor) === norm_(sup) && dayStr_(r.Date) === date) {
      if ((!num_(r.Lat) || !num_(r.Lng)) && num_(p.lat) && num_(p.lng)) { r.Lat = num_(p.lat); r.Lng = num_(p.lng); r['Accuracy m'] = num_(p.acc); update_('SiteVisits', r._row, r); }
      return { ok: true, duplicate: true };
    }
  }
  append_('SiteVisits', { 'Created At': created, Date: date, Offer: offer, Supervisor: sup, Lat: num_(p.lat) || '', Lng: num_(p.lng) || '',
    'Accuracy m': num_(p.acc) || '', Client: p.client || '', Positions: p.positions || '', Source: /handover/i.test(String(p.source || '')) ? 'Handover' : (p.source || 'SiteLog') });
  // first GPS for an offer becomes its (unconfirmed) site pin
  if (num_(p.lat) && num_(p.lng) && !pinOf_(offer)) append_('Pins', { Offer: offer, Lat: num_(p.lat), Lng: num_(p.lng), Source: 'first visit', 'Set At': created });
  return { ok: true };
}
function isMeasurement_(task) { return /measure/i.test(String(task || '')); }
function pinOf_(offer) {
  var o = cleanOffer_(offer), best = null;
  rows_('Pins').forEach(function (r) { if (cleanOffer_(r.Offer) === o && num_(r.Lat) && num_(r.Lng)) { if (!best || String(r.Source) === 'HR') best = r; } });
  return best ? { lat: num_(best.Lat), lng: num_(best.Lng), source: String(best.Source) } : null;
}

/* ============================ EXTRAS ============================ */
function addExtra_(p) {
  var sup = String(p.supervisor || '').trim(), amount = num_(p.amount);
  if (!sup || !amount) return { ok: false, error: 'supervisor and amount are required' };
  var url = '';
  if (p.photo) {
    var m = String(p.photo).match(/^data:([^;]+);base64,(.*)$/);
    if (m) {
      var folder = billFolder_();
      var name = dayStr_(p.date) + '_' + sup.replace(/\W+/g, '-') + '_' + (p.type || 'extra') + '_' + Date.now() + '.jpg';
      var f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      url = f.getUrl();
    }
  }
  var id = 'x' + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36);
  append_('Extras', { Id: id, Date: dayStr_(p.date), Supervisor: sup, Offer: cleanOffer_(p.offer), Type: p.type || 'Other', Amount: amount,
    Reason: p.reason || '', 'Bill URL': url, Status: url ? 'pending' : 'no bill', 'HR Note': '', 'Added At': new Date().toISOString() });
  return { ok: true, id: id, url: url };
}
function deleteExtra_(p) {
  var all = rows_('Extras');
  for (var i = 0; i < all.length; i++) if (String(all[i].Id) === String(p.id) && norm_(all[i].Supervisor) === norm_(p.supervisor)) {
    if (String(all[i].Status) === 'approved' || String(all[i].Status) === 'paid') return { ok: false, error: 'already approved' };
    sheet_('Extras').deleteRow(all[i]._row); return { ok: true };
  }
  return { ok: false, error: 'not found' };
}
/* supervisor types the km for a Measurement visit (0 clears it) */
function setMeasureKm_(p) {
  var date = dayStr_(p.date), sup = String(p.supervisor || '').trim(), offer = cleanOffer_(p.offer), km = num_(p.km);
  if (!date || !sup || !offer) return { ok: false, error: 'date, supervisor and Offer No are required' };
  if (date > today_()) return { ok: false, error: 'km can be entered only on the day or later' };
  if (km < 0 || km > 500) return { ok: false, error: 'km must be between 0 and 500' };
  var paid = rows_('Paid').some(function (r) { return String(r.Month).slice(0, 7) === date.slice(0, 7) && norm_(r.Supervisor) === norm_(sup); });
  if (paid) return { ok: false, error: 'this month is already paid' };
  var hit = null; rows_('MeasureKm').forEach(function (r) { if (norm_(r.Supervisor) === norm_(sup) && dayStr_(r.Date) === date && cleanOffer_(r.Offer) === offer) hit = r; });
  if (!km) { if (hit) sheet_('MeasureKm').deleteRow(hit._row); return { ok: true, cleared: true }; }
  var row = { Date: date, Supervisor: sup, Offer: offer, Km: km, Note: p.note || '', 'Updated At': new Date().toISOString() };
  if (hit) update_('MeasureKm', hit._row, row); else append_('MeasureKm', row);
  return { ok: true, km: km };
}
function billFolder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('BILL_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (x) {} }
  var f = DriveApp.createFolder('Lingel Conveyance Bills'); props.setProperty('BILL_FOLDER_ID', f.getId()); return f;
}

/* ============================ THE CHECKS ============================ */
/* Builds one supervisor-day: stops (plan ↔ report), route km, amount, status. */
function buildDay_(sup, date, ctx) {
  var S = ctx.settings, now = Date.now(), n = norm_(sup);
  var supRec = ctx.supByNorm[n] || { name: sup, vehicle: 'bike', lat: 0, lng: 0 };
  var plans = ctx.plans.filter(function (r) { return norm_(r.Supervisor) === n && dayStr_(r.Date) === date && !r.Deleted && String(r.Task) !== 'Leave'; });
  var visits = ctx.visits.filter(function (r) { return norm_(r.Supervisor) === n && dayStr_(r.Date) === date; });
  var extras = ctx.extras.filter(function (r) { return norm_(r.Supervisor) === n && dayStr_(r.Date) === date; });
  var cutoff = istMs_(addDays_(date, -1), S.cutoff_hour), lateBy = istMs_(addDays_(date, 1), S.late_report_hour);
  var past = date < ctx.today, used = {}, stops = [];

  plans.forEach(function (pl) {
    var offer = cleanOffer_(pl.Offer), st = { offer: offer, project: String(pl.Project || ''), city: String(pl.City || ''), task: String(pl.Task || ''), planned: true, checks: [] };
    var chk = function (n, res, msg) { st.checks.push({ n: n, res: res, msg: msg }); };
    chk(1, toMs_(pl['First At']) <= cutoff ? 'ok' : 'warn', toMs_(pl['First At']) <= cutoff ? 'Planned in time' : 'Planned after the cut-off');
    if (!offer) { chk(2, 'bad', 'No Offer No in the plan'); st.result = 'bad'; stops.push(st); return; }
    var offerOk = /^\d{3,7}(\/\d+[A-Z]?)?$/.test(offer);   // 75165 or 75057/14
    chk(2, offerOk ? 'ok' : 'warn', offerOk ? 'Offer No valid' : 'Offer No looks unusual');
    var v = null; visits.forEach(function (x) { if (!v && cleanOffer_(x.Offer) === offer) v = x; });
    if (isMeasurement_(pl.Task)) {
      used[offer] = true;
      chk(3, 'ok', 'Measurement — approved automatically');
      var mk = ctx.measKm[n + '|' + date + '|' + offer];
      st.manualKm = mk ? num_(mk.Km) : null;
      if (mk) chk(4, 'ok', 'Km entered by supervisor: ' + num_(mk.Km));
      else if (!past) { chk(4, 'wait', 'Enter the km for this measurement'); st.result = 'wait'; stops.push(st); return; }
      else chk(4, 'ok', 'No km entered — ₹0 travel for this visit');
      st.result = 'ok'; stops.push(st); return;
    }
    if (!v) { chk(3, past ? 'bad' : 'wait', past ? 'No Site Inspection report or Handover form' : 'Site report / Handover pending'); st.result = past ? 'bad' : 'wait'; stops.push(st); return; }
    used[offer] = true; addVisitChecks_(st, v, offer, ctx, S, date, lateBy, chk); stops.push(st);
  });
  // reports with no plan behind them
  visits.forEach(function (v) {
    var offer = cleanOffer_(v.Offer); if (used[offer]) return; used[offer] = true;
    var st = { offer: offer, project: String(v.Client || ''), city: '', task: 'Unplanned', planned: false, checks: [] };
    var chk = function (n, res, msg) { st.checks.push({ n: n, res: res, msg: msg }); };
    chk(1, 'warn', 'Not in the plan (urgent visit?)'); chk(2, 'ok', 'Offer No from report');
    addVisitChecks_(st, v, offer, ctx, S, date, lateBy, chk); stops.push(st);
  });
  stops.forEach(function (st) { if (!st.result) st.result = st.checks.some(function (c) { return c.res === 'bad'; }) ? 'bad' : (st.checks.some(function (c) { return c.res === 'warn'; }) ? 'warn' : 'ok'); });

  // route: base → paid-eligible stops in plan order → base
  var pts = stops.filter(function (s) { return (s.result === 'ok' || s.result === 'warn') && s.lat && s.lng; });
  var km = 0, kmNote = '', legs = [];
  var names = ['Base'].concat(pts.map(function (s) { return '#' + s.offer; })).concat(['Base']);
  if (pts.length) {
    if (supRec.lat && supRec.lng) {
      var r = routeKm_([{ lat: supRec.lat, lng: supRec.lng }].concat(pts).concat([{ lat: supRec.lat, lng: supRec.lng }]), ctx); km = r.km; kmNote = r.how;
      (r.legs || []).forEach(function (k, i) { legs.push({ from: names[i], to: names[i + 1], km: k }); });
    }
    else kmNote = 'No base location set for this supervisor';
  }
  // measurement visits: km typed by the supervisor, added to the route km
  stops.forEach(function (st) { if (st.manualKm) { km += st.manualKm; legs.push({ from: 'Measurement', to: '#' + st.offer, km: st.manualKm }); } });
  var mNames = stops.filter(function (st) { return st.manualKm; }).map(function (st) { return '#' + st.offer + ' (meas. ' + st.manualKm + ' km)'; });
  if (mNames.length && kmNote === 'No base location set for this supervisor') kmNote += ' (measurement km still counted)';
  var route = (pts.length ? names.join(' → ') : '') + (mNames.length ? (pts.length ? ' + ' : '') + mNames.join(' + ') : '');
  var rate = supRec.vehicle === 'car' ? S.rate_car : (supRec.vehicle === 'public' ? S.rate_public : S.rate_bike);
  var travel = Math.min(Math.round(km * rate), S.daily_cap);
  var capped = Math.round(km * rate) > S.daily_cap;
  var ex = extras.map(function (x) {
    var res = String(x.Status) === 'approved' ? 'ok' : (String(x.Status) === 'rejected' ? 'bad' : (!x['Bill URL'] ? 'bad' : (num_(x.Amount) > S.extra_flag_amount ? 'warn' : 'ok')));
    return { id: String(x.Id), type: String(x.Type), amount: num_(x.Amount), offer: String(x.Offer), reason: String(x.Reason), url: String(x['Bill URL']), status: String(x.Status), note: String(x['HR Note'] || ''), result: res };
  });
  var exOk = ex.filter(function (x) { return x.result === 'ok' || x.result === 'warn'; }).reduce(function (a, x) { return a + x.amount; }, 0);

  var dec = ctx.decByKey[n + '|' + date] || null;
  var anyWarn = stops.some(function (s) { return s.result === 'warn'; }) || ex.some(function (x) { return x.result === 'warn'; }) || capped || (pts.length && !km);
  var anyPaid = stops.some(function (s) { return s.result === 'ok' || s.result === 'warn'; });
  var status, amount = travel + exOk, reason = '';
  if (dec) { status = String(dec.Status); if (dec.Amount !== '' && dec.Amount !== null && !isNaN(parseFloat(dec.Amount))) amount = num_(dec.Amount); reason = String(dec.Note || ''); if (status === 'rejected') amount = 0; }
  else if (!past && stops.some(function (s) { return s.result === 'wait'; })) { status = 'in progress'; }
  else if (!anyPaid && !exOk) { status = stops.length ? 'not eligible' : 'no plan'; amount = 0; reason = stops.length ? 'No planned site has a report' : ''; }
  else if (anyWarn) status = 'review';
  else status = 'auto-approved';
  if (ctx.paidKey[n + '|' + date.slice(0, 7)] && (status === 'approved' || status === 'auto-approved')) status = 'paid';

  var approvedBy = dec ? (String(dec.By || 'HR') + (dec.At ? ' · ' + dayStr_(dec.At) : '')) : (status === 'auto-approved' || status === 'paid' ? 'Auto (all checks passed)' : '');
  return { date: date, supervisor: supRec.name || sup, vehicle: supRec.vehicle, rate: rate, km: Math.round(km * 10) / 10, kmNote: kmNote, route: route, legs: legs, approvedBy: approvedBy, base: supRec.base || '',
           travel: travel, capped: capped, extras: ex, extrasOk: exOk, amount: Math.round(amount), status: status, reason: reason,
           stops: stops, planned: plans.length, reported: stops.filter(function (s) { return s.result === 'ok' || s.result === 'warn'; }).length };
}
function addVisitChecks_(st, v, offer, ctx, S, date, lateBy, chk) {
  st.lat = num_(v.Lat); st.lng = num_(v.Lng); st.reportAt = String(v['Created At']);
  chk(3, 'ok', /handover/i.test(String(v.Source || '')) ? 'Handover form filed' : 'Site report filed');
  var pin = ctx.pinByOffer[offer];
  if (!st.lat || !st.lng) chk(4, 'warn', 'Report has no GPS');
  else if (!pin) chk(4, 'ok', 'First visit — GPS becomes the site pin');
  else {
    var d = Math.round(metres_(st.lat, st.lng, pin.lat, pin.lng)); st.distM = d;
    chk(4, d <= S.gps_radius_m ? 'ok' : 'warn', d <= S.gps_radius_m ? ('At site (' + d + ' m)') : ('Report ' + (d >= 1000 ? (Math.round(d / 100) / 10 + ' km') : (d + ' m')) + ' from the site pin'));
  }
  var c = toMs_(v['Created At']); var cDay = isNaN(c) ? '' : Utilities.formatDate(new Date(c), TZ, 'yyyy-MM-dd');
  chk(5, cDay === date && c <= lateBy ? 'ok' : 'warn', cDay === date ? 'Report made the same day' : 'Report made on ' + (cDay || 'unknown date'));
}

/* Driving km with Google Maps; straight-line × 1.3 if Maps fails. Cached per route. */
function routeKm_(pts, ctx) {
  var key = 'rt:' + pts.map(function (p) { return p.lat.toFixed(4) + ',' + p.lng.toFixed(4); }).join(';');
  var cache = CacheService.getScriptCache(), hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  var res = null;
  try {
    var df = Maps.newDirectionFinder().setOrigin(pts[0].lat, pts[0].lng).setDestination(pts[pts.length - 1].lat, pts[pts.length - 1].lng).setMode(Maps.DirectionFinder.Mode.DRIVING);
    for (var i = 1; i < pts.length - 1; i++) df.addWaypoint(pts[i].lat, pts[i].lng);
    var d = df.getDirections();
    if (d && d.routes && d.routes.length) {
      var legs = d.routes[0].legs.map(function (l) { return Math.round(l.distance.value / 100) / 10; });
      var m = d.routes[0].legs.reduce(function (a, l) { return a + l.distance.value; }, 0);
      res = { km: Math.round(m / 100) / 10, legs: legs, how: 'Google Maps driving distance' };
    }
  } catch (x) {}
  if (!res) {
    var s = 0, lg = [];
    for (var j = 1; j < pts.length; j++) { var mm = metres_(pts[j - 1].lat, pts[j - 1].lng, pts[j].lat, pts[j].lng); s += mm; lg.push(Math.round(mm * 1.3 / 100) / 10); }
    res = { km: Math.round(s * 1.3 / 100) / 10, legs: lg, how: 'Estimated (straight line × 1.3)' };
  }
  cache.put(key, JSON.stringify(res), 21600);
  return res;
}

function ctx_() {
  var S = settings_(), supByNorm = {}; supervisors_().forEach(function (s) { supByNorm[norm_(s.name)] = s; });
  var measKm = {}; rows_('MeasureKm').forEach(function (r) { measKm[norm_(r.Supervisor) + '|' + dayStr_(r.Date) + '|' + cleanOffer_(r.Offer)] = r; });
  var pinByOffer = {}; rows_('Pins').forEach(function (r) { var o = cleanOffer_(r.Offer); if (num_(r.Lat) && num_(r.Lng) && (!pinByOffer[o] || String(r.Source) === 'HR')) pinByOffer[o] = { lat: num_(r.Lat), lng: num_(r.Lng), source: String(r.Source) }; });
  var decByKey = {}; rows_('Decisions').forEach(function (r) { decByKey[norm_(r.Supervisor) + '|' + dayStr_(r.Date)] = r; });
  var paidKey = {}; rows_('Paid').forEach(function (r) { paidKey[norm_(r.Supervisor) + '|' + String(r.Month).slice(0, 7)] = r; });
  return { settings: S, measKm: measKm, supByNorm: supByNorm, pinByOffer: pinByOffer, decByKey: decByKey, paidKey: paidKey,
           plans: rows_('Plans'), visits: rows_('SiteVisits'), extras: rows_('Extras'), today: today_() };
}
function monthDays_(month) {
  var m = /^\d{4}-\d{2}$/.test(String(month || '')) ? String(month) : today_().slice(0, 7);
  var p = m.split('-').map(Number), days = [], last = new Date(Date.UTC(p[0], p[1], 0)).getUTCDate();
  for (var d = 1; d <= last; d++) days.push(m + '-' + (d < 10 ? '0' : '') + d);
  return { month: m, days: days };
}
function activeDays_(sup, md, ctx) {
  var n = norm_(sup), set = {};
  ctx.plans.forEach(function (r) { var d = dayStr_(r.Date); if (norm_(r.Supervisor) === n && !r.Deleted && String(r.Task) !== 'Leave' && d.slice(0, 7) === md.month) set[d] = 1; });
  ctx.visits.forEach(function (r) { var d = dayStr_(r.Date); if (norm_(r.Supervisor) === n && d.slice(0, 7) === md.month) set[d] = 1; });
  ctx.extras.forEach(function (r) { var d = dayStr_(r.Date); if (norm_(r.Supervisor) === n && d.slice(0, 7) === md.month) set[d] = 1; });
  return Object.keys(set).sort();
}

/* ============================ READ VIEWS ============================ */
function myMonth_(sup, month) {
  if (!String(sup || '').trim()) return { ok: false, error: 'supervisor name missing' };
  var ctx = ctx_(), md = monthDays_(month);
  var days = activeDays_(sup, md, ctx).map(function (d) { return buildDay_(sup, d, ctx); }).reverse();
  var total = days.filter(function (d) { return d.status === 'approved' || d.status === 'auto-approved' || d.status === 'paid'; }).reduce(function (a, d) { return a + d.amount; }, 0);
  var pending = days.filter(function (d) { return d.status === 'review' || d.status === 'in progress'; }).reduce(function (a, d) { return a + d.amount; }, 0);
  var sr = ctx.supByNorm[norm_(sup)];
  return { ok: true, month: md.month, supervisor: sr ? sr.name : sup, known: !!sr, hasBase: !!(sr && sr.lat && sr.lng), settings: ctx.settings,
           today: buildDay_(sup, ctx.today, ctx), days: days, total: total, pending: pending };
}
function hrMonth_(month) {
  var ctx = ctx_(), md = monthDays_(month), names = {};
  Object.keys(ctx.supByNorm).forEach(function (k) { if (ctx.supByNorm[k].active) names[k] = ctx.supByNorm[k].name; });
  ctx.plans.concat(ctx.visits).forEach(function (r) { var d = dayStr_(r.Date); if (d.slice(0, 7) === md.month && String(r.Supervisor || '').trim()) { var k = norm_(r.Supervisor); if (!names[k]) names[k] = String(r.Supervisor).trim(); } });
  var people = Object.keys(names).sort().map(function (k) {
    var sup = names[k], days = activeDays_(sup, md, ctx).map(function (d) { return buildDay_(sup, d, ctx); });
    var sum = function (f) { return days.filter(f).reduce(function (a, d) { return a + d.amount; }, 0); };
    return { supervisor: sup, known: !!ctx.supByNorm[k], hasBase: !!(ctx.supByNorm[k] && ctx.supByNorm[k].lat), vehicle: ctx.supByNorm[k] ? ctx.supByNorm[k].vehicle : '',
             days: days, approved: sum(function (d) { return d.status === 'approved' || d.status === 'auto-approved'; }), paid: sum(function (d) { return d.status === 'paid'; }),
             review: days.filter(function (d) { return d.status === 'review'; }).length, km: Math.round(days.reduce(function (a, d) { return a + (d.status === 'rejected' ? 0 : d.km); }, 0)),
             visits: days.reduce(function (a, d) { return a + d.reported; }, 0), planned: days.reduce(function (a, d) { return a + d.planned; }, 0),
             isPaid: !!ctx.paidKey[k + '|' + md.month] };
  });
  return { ok: true, month: md.month, days: md.days, today: ctx.today, settings: ctx.settings, people: people };
}

/* ============================ HR ACTIONS ============================ */
function hrDecide_(p) {
  var date = dayStr_(p.date), sup = String(p.supervisor || '').trim(), status = String(p.status || '');
  if (!date || !sup || ['approved', 'rejected', 'clear'].indexOf(status) < 0) return { ok: false, error: 'date, supervisor and status are required' };
  var all = rows_('Decisions'), hit = null;
  all.forEach(function (r) { if (norm_(r.Supervisor) === norm_(sup) && dayStr_(r.Date) === date) hit = r; });
  if (status === 'clear') { if (hit) sheet_('Decisions').deleteRow(hit._row); return { ok: true }; }
  if (status === 'rejected' && !String(p.note || '').trim()) return { ok: false, error: 'a reason is required to reject' };
  var row = { Date: date, Supervisor: sup, Status: status, Amount: p.amount === undefined || p.amount === '' ? '' : num_(p.amount), Note: p.note || '', By: p.by || 'HR', At: new Date().toISOString() };
  if (hit) update_('Decisions', hit._row, row); else append_('Decisions', row);
  // approving a day also approves its pending extras; rejecting rejects them
  rows_('Extras').forEach(function (x) { if (norm_(x.Supervisor) === norm_(sup) && dayStr_(x.Date) === date && ['pending', 'no bill'].indexOf(String(x.Status)) >= 0 && x['Bill URL']) { x.Status = status; update_('Extras', x._row, x); } });
  return { ok: true };
}
function hrExtra_(p) {
  var all = rows_('Extras');
  for (var i = 0; i < all.length; i++) if (String(all[i].Id) === String(p.id)) { all[i].Status = p.status; all[i]['HR Note'] = p.note || ''; update_('Extras', all[i]._row, all[i]); return { ok: true }; }
  return { ok: false, error: 'not found' };
}
function hrSaveSettings_(p) {
  var vals = {}; try { vals = JSON.parse(p.settings || '{}'); } catch (x) {}
  settings_(); var rs = rows_('Settings');
  rs.forEach(function (r) { if (vals[r.Key] !== undefined) { r.Value = num_(vals[r.Key]); update_('Settings', r._row, r); } });
  return { ok: true, settings: settings_() };
}
function hrSaveSupervisor_(p) {
  var name = String(p.name || '').trim(); if (!name) return { ok: false, error: 'name required' };
  var row = { Name: name, 'Base address': p.base || '', 'Base lat': num_(p.lat) || '', 'Base lng': num_(p.lng) || '', Vehicle: p.vehicle || 'bike', Active: p.active === 'no' ? 'no' : 'yes' };
  var hit = null; supervisors_().forEach(function (s) { if (norm_(s.name) === norm_(p.oldName || name)) hit = s; });
  if (hit) update_('Supervisors', hit._row, row); else append_('Supervisors', row);
  return { ok: true };
}
function hrPin_(p) {
  var offer = cleanOffer_(p.offer); if (!offer || !num_(p.lat) || !num_(p.lng)) return { ok: false, error: 'offer, lat and lng required' };
  append_('Pins', { Offer: offer, Lat: num_(p.lat), Lng: num_(p.lng), Source: 'HR', 'Set At': new Date().toISOString() });
  return { ok: true };
}
function hrPaid_(p) {
  var month = String(p.month || '').slice(0, 7), sup = String(p.supervisor || '').trim();
  if (!month || !sup) return { ok: false, error: 'month and supervisor required' };
  var exists = rows_('Paid').some(function (r) { return String(r.Month).slice(0, 7) === month && norm_(r.Supervisor) === norm_(sup); });
  if (!exists) append_('Paid', { Month: month, Supervisor: sup, Amount: num_(p.amount), 'Paid At': new Date().toISOString(), By: p.by || 'HR' });
  return { ok: true };
}

/* Freeze the month for accounts: every approved / paid day → Ledger tab (re-closing replaces that month's rows). */
function hrCloseMonth_(p) {
  var data = hrMonth_(p.month), month = data.month, now = new Date().toISOString();
  var sh = sheet_('Ledger'), old = rows_('Ledger').filter(function (r) { return String(r.Month).slice(0, 7) === month; });
  old.map(function (r) { return r._row; }).sort(function (a, b) { return b - a; }).forEach(function (n) { sh.deleteRow(n); });
  var out = [], review = 0, total = 0;
  data.people.forEach(function (pp) {
    pp.days.forEach(function (d) {
      if (d.status === 'review' || d.status === 'in progress') { review++; return; }
      if (['approved', 'auto-approved', 'paid'].indexOf(d.status) < 0) return;
      var sites = d.stops.filter(function (s) { return s.result === 'ok' || s.result === 'warn'; }).map(function (s) { return '#' + s.offer; }).join(', ');
      out.push([month, d.date, d.supervisor, d.vehicle, sites, d.route, d.km, d.rate, d.travel, d.extrasOk, d.amount, d.status, d.approvedBy, now]);
      total += d.amount;
    });
  });
  out.sort(function (a, b) { return a[2] === b[2] ? (a[1] < b[1] ? -1 : 1) : (a[2] < b[2] ? -1 : 1); });
  if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, TABS.Ledger.length).setValues(out);
  return { ok: true, month: month, rows: out.length, total: total, stillInReview: review, closedAt: now };
}