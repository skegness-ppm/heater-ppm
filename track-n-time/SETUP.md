# Track’n’Time – setup (about 5 minutes, best done on a computer)

The Google Sheet **"Track’n’Time – Butlin's Skegness"** is already in your Google Drive. It's the database: every job, timesheet entry, material and note is a row in it.

## 1. Put the code in the sheet
1. Open the sheet → **Extensions → Apps Script**.
2. In `Code.gs`, delete what's there and paste everything from `apps-script/Code.gs`.
3. Click **+ → HTML**, name it `Index` (exactly), and paste everything from `apps-script/Index.html`.
4. Click **Save** (disk icon).
5. Pick `setup` in the function drop-down and click **Run**. Google asks for permission → **Review permissions → your account → Advanced → Go to (unsafe) → Allow**. It asks for access to your spreadsheets because it reads the village units and venues from your *Wall Heaters CBRE – Mobile Friendly V1* checklist. The sheet now has its tabs and a **Locations** tab filled from the heater list.

## 2. Publish it
1. **Deploy → New deployment** → gear icon → **Web app**.
2. Execute as: **Me**. Who has access: **Anyone**.
3. **Deploy**, then copy the **Web app URL** (ends in `/exec`) and send it to Claude.

## 3. First login
Open the link, enter your name and choose a PIN. You're the first admin. Then **Team → + Add engineer** for each person.

## Updating later
Paste the new code, then **Deploy → Manage deployments → pencil → Version: New version → Deploy**. The link stays the same.

## How it works
The **timesheet** is the home screen. Raising a job books it onto the timesheet: you pick the day, start time and how long you think it needs (your quote). Actual time is logged by clocking on, confirming the booked block, or adding time.

**Job stages:** Scheduled → In progress → To Admin → Complete

| | Admin | Engineer |
|---|---|---|
| Raise and book jobs | For anyone, or unassigned | Onto their own timesheet |
| Assign / reassign | ✓ (moves the booking) | – |
| Scheduled ↔ In progress | ✓ | Own jobs |
| Send to Admin (with report) | ✓ | Own jobs |
| Mark Complete (after paperwork check) | ✓ | – |
| Send back to engineer | ✓ | – |
| Add people, reset PINs, refresh locations | ✓ | – |

Locations: villages and units, and venues with floors, come from the heater PPM checklist. Admins can tap **Team → Refresh from heater list** after the checklist changes. Extra areas (plant rooms, car parks…) live in the **Areas** tab, and the quick buttons on Log time live in **Presets**.
