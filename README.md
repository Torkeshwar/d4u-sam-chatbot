# D4U Tracker | SAM — Issue Chatbot (export-based)

A local chatbot that answers questions about your **RD_DATA4YOU_L2_COGNIZANT**
incidents. It reads a file you **export from ServiceNow** — no login, cookies,
OAuth, or admin rights needed. Everything runs locally on your machine.

## Setup (one-time)

```powershell
cd c:\Users\AShukl77\data4you\d4u-sam-chatbot
Copy-Item config.example.json config.json
pip install -r requirements.txt
python app.py
```

Open http://localhost:8766.

## Export your incidents from ServiceNow

1. In ServiceNow, open your incident list filtered to your group
   (`RD_DATA4YOU_L2_COGNIZANT`). The dashboard's "Open Assigned Incidents" /
   "Assigned to me" lists work too — open the list view.
2. Right-click the list header row → **Export → CSV** (CSV is most reliable;
   Excel also works).
3. The file lands in your Downloads folder.

## Load it into the chatbot

Either:
- Click **Upload CSV/Excel** and pick the file, or
- Drag the file onto the page, or
- Click **Load latest from Downloads** to auto-pick the newest export.

## Ask questions

- `summary` — totals, open count, breakdown by priority
- `open incidents`
- `open P1` / `critical`
- `unassigned`
- `assigned to <name>`
- `resolved` / `closed`
- `list all`
- `status of SINC0492904`
- any other text → searches descriptions / numbers / assignees

## Refreshing data

Data is a snapshot of your last export. To refresh, re-export from ServiceNow
and load the new file (or click **Load latest from Downloads** again).

## Notes

- Works offline; no connection to ServiceNow is ever made.
- Column headers vary between exports; the app maps common ServiceNow labels
  (Number, Short description, Priority, State, Assigned to, etc.) automatically.
- Excel support needs `openpyxl` (included in requirements). CSV needs nothing extra.
