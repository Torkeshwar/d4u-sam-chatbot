# D4U Tracker | SAM — Incident Chatbot

A lightweight, local assistant for exploring **RD_DATA4YOU_L2_COGNIZANT**
incidents exported from ServiceNow (`sam1.service-now.com`). You ask plain
questions, it returns counts, filtered tables, breakdowns, and monthly charts.

---

## 1. What it is (and is not)

- **Is:** a deterministic reporting/query tool over an exported incident file.
  All numbers are computed in code (Python), so they are exact and repeatable.
- **Is not:** an AI/LLM model, and not a live ServiceNow integration. It does not
  guess or estimate; it filters and counts the rows you loaded.

---

## 2. How it works (architecture)

```
ServiceNow  --(you export CSV/Excel)-->  incident file
                                               |
                             you Upload / Load in the browser UI
                                               |
            Local Python app (FastAPI)  <-- holds rows in memory only
                                               |
     You type a question  ->  rule-based parser  ->  deterministic filter/count
                                               |
                 Browser renders: tables, counts, breakdowns, charts
```

- **Backend:** Python + FastAPI (`app.py`), served by uvicorn on
  `http://localhost:8766`.
- **Frontend:** a single HTML page (embedded in `app.py`) with a chat UI.
  Charts use Chart.js (loaded from a CDN).
- **Data:** kept in memory for the session only. Nothing is written to disk and
  nothing is sent to any external service.

---

## 3. Why export-based (not live API)

The ServiceNow instance uses corporate SSO, and the account available did not
have rights to create an OAuth app or use the API directly. Cookie/session reuse
also failed (CSRF-protected endpoints returned 401). The reliable, no-admin path
was therefore to read a **manual export**, which needs no special access.

---

## 4. Setup & run

```powershell
cd c:\Users\AShukl77\data4you\d4u-sam-chatbot
Copy-Item config.example.json config.json   # first time only
pip install -r requirements.txt
python app.py
```

Open `http://localhost:8766`.

### Export from ServiceNow
Open your incident list filtered to `RD_DATA4YOU_L2_COGNIZANT`, right-click the
column header → **Export → CSV**. Then in the app: **Upload CSV/Excel**,
drag-drop, or **Load latest from Downloads**.

---

## 5. Features / what you can ask

| Category | Examples |
|---|---|
| Greetings | `hi`, `hello`, `thanks`, `bye` |
| Summary | `summary` (totals, open/closed, unassigned, priority mix) |
| Status filters | `open incidents`, `resolved`, `closed`, `on hold`/`onhold`, `in progress`/`inprogress`, `new`, `canceled` |
| Priority | `open P2`, `critical`/`high`/`moderate`/`low`/`planning`, `P3 incidents` |
| People | `assigned to <name>`, `unassigned`, `how many worked by <name>` |
| Therapeutic area | `incidents in Oncology`, `how many cases in Immunology` |
| Study | `incidents for study 77242113UCO3001` |
| Dates / months | `cases in november`, `opened in october` |
| Counts | `how many open incidents`, `count unassigned` |
| Combined filters | `open P2 in Oncology assigned to <name>`, `on hold in study X` |
| Breakdowns | `group by study`, `breakdown by assignee`, `by category`, `by state` |
| Monthly charts | `monthly chart by status` / `by assignee` / `by study` / `by priority` |
| Specific lookup | `status of SINC0492904` |
| Free-text search | any text → searches description / number / assignee / category |
| Listing | `list all` (removes the default 100-row display cap) |

Type `help` in the app for a full, data-aware guide (it fills in real example
values from your loaded export).

---

## 6. Accuracy & data handling

- **Accuracy:** every count/filter/group/chart is computed by Python iterating
  over the loaded rows. No language model is involved in any calculation.
- **Scope:** results reflect only the rows in the loaded export. If the export
  was filtered or capped in ServiceNow, the numbers mirror that export.
- **Freshness:** data is a snapshot as of the last export. Re-export and reload
  to refresh.
- **Privacy:** runs entirely on the local machine; the session data stays in
  memory and is never persisted or transmitted. `config.json` is gitignored.

---

## 7. Known limitations

- Not live — manual export required to refresh.
- Single-user, local, no authentication or audit log.
- "Opened in <month>" uses the incident **opened date**, not status history.
- Charts need the Chart.js CDN; if the network blocks it, charts won't render
  (the rest still works). Can be bundled offline if needed.

---

## 8. Possible next steps

- Suggested-question and follow-up chips driven by the loaded data.
- Data-driven insight cards (e.g., aging open incidents, unassigned backlog).
- Optional local LLM (via Ollama) to interpret free-form phrasing — with Python
  still doing all calculations for accuracy. (Would fall under AI governance.)
- Proper ServiceNow API integration for live data (needs a service account /
  OAuth client from an admin) and a reviewed, hosted deployment.

---

## 9. File overview

| File | Purpose |
|---|---|
| `app.py` | The entire application: server, parser, and UI |
| `config.json` | Settings (assignment group, downloads dir); gitignored |
| `config.example.json` | Template for `config.json` |
| `requirements.txt` | Python dependencies |
| `start.ps1` | Convenience launcher |
| `sample_export.csv` | Small sample file to try the app |
| `README.md` | Quick start |
| `DOCUMENTATION.md` | This document |
