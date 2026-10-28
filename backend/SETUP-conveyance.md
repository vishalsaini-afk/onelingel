# Plan-linked conveyance — setup (about 15 minutes)

## What's in this pack
| File | What it is | Where it goes |
|---|---|---|
| `conveyance-backend.gs` | New, separate backend. Your invoice and SiteLog backends are not touched. | A new Google Sheet's Apps Script |
| `team-availability.html` | Your current app + a new **₹ Conveyance** tab for supervisors | Replace the file in the hub folder |
| `site-inspection-portal.html` | Your current SiteLog + records the visit (Offer No, date, live GPS) on every report | Replace the file in the hub folder |
| `conveyance-hr.html` | New HR dashboard: Review, Board, Month close, Settings | Add to the hub folder |

## 1. Backend (once)
1. Create a new Google Sheet named **Lingel Conveyance** → Extensions → Apps Script.
2. Delete the sample code, paste all of `conveyance-backend.gs`, save.
3. Choose the function **setup** → Run → allow access (Sheets, Drive for bill photos, Maps for km).
4. Project Settings → Script properties → Add: `HR_KEY` = a password for HR.
5. Deploy → New deployment → Web app → Execute as **Me**, Who has access **Anyone** → Deploy → copy the **/exec** link.

## 2. Apps
1. Upload the three HTML files to the same folder as `index.html` (the hub), replacing the old ones.
2. Optional: paste the /exec link into `const CONV_DEFAULT=''` in `team-availability.html`, so supervisors are never asked for it.

## 3. HR, first time
1. Open `conveyance-hr.html`, paste the /exec link and the HR key.
2. **Settings → Rate card**: set ₹/km for bike and car, daily cap, cut-off hour, at-site radius.
3. **Settings → Supervisors**: one row each. The **name must match** the name used in Team Availability plans. Base location = paste "lat, lng" or a Google Maps link.

## How it runs every day
- **Planning** (Team Availability → Board → Supervisors, as today): each plan needs an **Offer No.**, made before 9 pm the day before.
- **Supervisor**: at the site, files the SiteLog report as usual. The PDF/Share/Drive button also records the visit with GPS. The ₹ Conveyance tab shows today's sites, the "Open SiteLog" button, bills, and the month.
- **Bills**: toll, parking, auto/cab, bus/train, stay, with a photo, from ₹ Conveyance → Add bill.
- **System**: matches plan ↔ report, works out km (Google Maps, base → sites → base) × rate, and runs the 6 checks. Clean days approve themselves.
- **HR**: Review tab = only days that need a human. Approve (amount can be changed) or Reject with a reason. Month close → Payroll CSV → Mark paid.

## Notes
- The first GPS for an Offer No. becomes its site pin; HR can replace it with "Use as site pin".
- Plans made in the last 35 days on a phone are sent to the backend the first time that phone opens the ₹ Conveyance tab.
- If a phone is offline, visits and bills are saved on the phone and sent when it's back online.
