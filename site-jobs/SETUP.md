# Site Jobs – setup (about 5 minutes, best done on a computer)

The Google Sheet **"Site Jobs – Butlin's Skegness"** is already in your Google Drive. It's the database: every job, timesheet entry, material and note is a row in it.

## 1. Put the code in the sheet
1. Open the sheet → **Extensions → Apps Script**.
2. In `Code.gs`, delete what's there and paste everything from `apps-script/Code.gs` in this folder.
3. Click **+ → HTML**, name it `Index` (exactly), and paste everything from `apps-script/Index.html`.
4. Click **Save** (disk icon).
5. Pick `setup` in the function drop-down and click **Run**. Google asks for permission → **Review permissions → your account → Advanced → Go to (unsafe) → Allow**. This is your own script; Google shows that warning for any script that isn't published. The sheet now has its tabs.

## 2. Publish it
1. **Deploy → New deployment** → gear icon → **Web app**.
2. Execute as: **Me**. Who has access: **Anyone**.
3. **Deploy**, then copy the **Web app URL** (ends in `/exec`) and send it to Claude. Claude puts it into the home-screen app.

## 3. First login
Open the link, enter your name and choose a PIN. You're the first admin. Then go to **Team → + Add engineer** for each person and give them their PIN. Tick "Admin" for anyone else who should assign jobs and sign off.

## Updating later
When the code changes: paste the new code, then **Deploy → Manage deployments → pencil → Version: New version → Deploy**. The link stays the same.

## How it works
| | Admin | Engineer |
|---|---|---|
| Raise jobs | For anyone, or unassigned | Assigned to themselves |
| Assign / reassign jobs | ✓ | – |
| Work a job (time, materials, status) | Any job | Their own jobs |
| Complete a job | ✓ | ✓ → goes to *Awaiting sign-off* |
| Sign off / send back | ✓ | – |
| Schedule time (e.g. PPM emergency lights) | Anyone's | Their own |
| Add people, reset PINs | ✓ | – |

Areas and the quick "What" buttons on the timesheet (PPM – Emergency lights etc.) come from the **Areas** and **Presets** tabs in the sheet. Edit those tabs to change them.
