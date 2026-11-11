"""
D4U Tracker | SAM - issue chatbot (export-based, no auth required).

How it works:
  1. In ServiceNow, open your incident list (filtered to RD_DATA4YOU_L2_COGNIZANT),
     then right-click the header > Export > CSV (or Excel).
  2. Either upload that file here, or drop it in your Downloads folder and let the
     app auto-detect the newest ServiceNow export.
  3. The chatbot answers questions about that data entirely locally. No ServiceNow
     connection, cookies, or credentials involved.

Data freshness = your last export. Re-export and re-load to refresh.
"""

import csv
import glob
import io
import json
import os
import re
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.json"


def load_config() -> dict:
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
            return json.load(fh)
    return {}


CONFIG = load_config()
GROUP = CONFIG.get("assignment_group", "RD_DATA4YOU_L2_COGNIZANT")
DOWNLOADS = Path(os.path.expanduser(CONFIG.get("downloads_dir", "~/Downloads")))

# In-memory dataset: list of dict rows, plus metadata about the loaded file.
DATA = {"rows": [], "columns": [], "source": None, "loaded_at": None, "count": 0}

app = FastAPI(title="D4U SAM Chatbot (export)")


# --------------------------------------------------------------------------- #
# Column normalization
# --------------------------------------------------------------------------- #
# ServiceNow export headers vary (display labels vs field names). Map common ones
# to canonical keys so the chatbot works regardless of the exact export format.
COLUMN_ALIASES = {
    "number": ["number", "task", "incident"],
    "short_description": ["short description", "short_description", "description", "summary"],
    "priority": ["priority", "pri"],
    "state": ["state", "status", "incident state"],
    "assigned_to": ["assigned to", "assigned_to", "assignee", "owner"],
    "assignment_group": ["assignment group", "assignment_group", "group"],
    "opened_at": ["opened", "opened at", "opened_at", "created", "sys_created_on"],
    "updated": ["updated", "updated at", "sys_updated_on", "last updated"],
    "category": ["category"],
    "caller": ["caller", "caller id", "requested for"],
    "study": ["u_study", "study", "study id", "study number"],
    "therapeutic_area": ["u_study.u_therapeutic_area", "therapeutic area", "therapeutic_area", "ta"],
    "idp": ["u_study.u_idp", "idp"],
}


def _canon(header: str) -> str:
    h = (header or "").strip().lower()
    for canonical, variants in COLUMN_ALIASES.items():
        if h in variants:
            return canonical
    return h  # keep unknown columns under their own lowercased name


def _normalize_rows(raw_rows: list[dict]) -> list[dict]:
    out = []
    for r in raw_rows:
        row = {}
        for k, v in r.items():
            row[_canon(k)] = (v or "").strip() if isinstance(v, str) else v
        out.append(row)
    return out


# --------------------------------------------------------------------------- #
# File loading (CSV, and best-effort Excel/HTML exports)
# --------------------------------------------------------------------------- #
def _sniff_delimiter(sample: str) -> str:
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t").delimiter
    except Exception:
        return ","


def load_csv_text(text: str) -> list[dict]:
    # Strip a possible BOM.
    if text and text[0] == "\ufeff":
        text = text[1:]
    sample = text[:4096]
    delim = _sniff_delimiter(sample)
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)
    return [dict(r) for r in reader]


def load_bytes(name: str, data: bytes) -> tuple[list[dict], str | None]:
    lower = name.lower()
    # Try text/CSV first (ServiceNow "Export > CSV").
    if lower.endswith((".csv", ".txt")) or b"," in data[:2000] or b";" in data[:2000]:
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                text = data.decode(enc)
                rows = load_csv_text(text)
                if rows:
                    return rows, None
            except Exception:
                continue
    # Excel (.xlsx) - optional dependency.
    if lower.endswith((".xlsx", ".xls")):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            ws = wb.active
            it = ws.iter_rows(values_only=True)
            headers = [str(c) if c is not None else "" for c in next(it)]
            rows = []
            for vals in it:
                rows.append({headers[i]: ("" if v is None else str(v)) for i, v in enumerate(vals) if i < len(headers)})
            return rows, None
        except ModuleNotFoundError:
            return [], "This is an Excel file. Install support with: pip install openpyxl — or export as CSV instead."
        except Exception as exc:
            return [], f"Could not read Excel file: {exc}"
    return [], "Unrecognized file format. Please export from ServiceNow as CSV."


def set_dataset(rows: list[dict], source: str) -> dict:
    norm = _normalize_rows(rows)
    DATA["rows"] = norm
    DATA["columns"] = sorted({k for r in norm for k in r.keys()})
    DATA["source"] = source
    DATA["loaded_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    DATA["count"] = len(norm)
    return DATA


def find_latest_download() -> Path | None:
    """Find the newest likely ServiceNow incident export in Downloads."""
    if not DOWNLOADS.exists():
        return None
    candidates = []
    for pat in ("*.csv", "*.xlsx", "*.xls"):
        candidates.extend(DOWNLOADS.glob(pat))
    # Prefer files whose name hints at incidents/servicenow; else newest overall.
    def score(p: Path):
        n = p.name.lower()
        hint = any(k in n for k in ("incident", "sn_", "servicenow", "task", "d4u", "sam"))
        return (1 if hint else 0, p.stat().st_mtime)
    if not candidates:
        return None
    candidates.sort(key=score, reverse=True)
    return candidates[0]


# --------------------------------------------------------------------------- #
# Query / filtering over the in-memory dataset
# --------------------------------------------------------------------------- #
OPEN_STATES = {"new", "in progress", "on hold", "open", "active", "work in progress", "assigned"}
CLOSED_STATES = {"resolved", "closed", "cancelled", "canceled", "closed complete", "closed/resolved"}


def _is_open(row) -> bool:
    st = str(row.get("state", "")).strip().lower()
    if not st:
        return True
    if st in CLOSED_STATES:
        return False
    return True


def _month_of(row) -> str:
    """Extract the 2-digit month from opened_at (handles YYYY-MM-DD and DD-MM-YYYY-ish)."""
    val = str(row.get("opened_at", "")).strip()
    if not val:
        return ""
    m = re.search(r"^\s*(\d{4})-(\d{2})-(\d{2})", val)
    if m:
        return m.group(2)
    # Fallback: dd/mm/yyyy or dd-mm-yyyy
    m2 = re.search(r"^\s*\d{1,2}[/-](\d{1,2})[/-]\d{2,4}", val)
    if m2:
        return m2.group(1).zfill(2)
    return ""


def _priority_rank(row) -> str:
    p = str(row.get("priority", "")).strip().lower()
    # Normalize "1 - Critical" -> "1"
    m = re.match(r"\s*(\d)", p)
    return m.group(1) if m else p


def _fmt(rows, limit=30) -> str:
    if not rows:
        return "No matching incidents in the loaded export."
    lines = []
    for r in rows[:limit]:
        num = r.get("number", "?")
        pri = r.get("priority", "")
        st = r.get("state", "")
        desc = r.get("short_description", "")
        who = r.get("assigned_to", "") or "Unassigned"
        tag = " / ".join([x for x in (pri, st) if x])
        lines.append(f"• {num}  [{tag}]  {desc}  — {who}")
    extra = f"\n… and {len(rows) - limit} more." if len(rows) > limit else ""
    return f"{len(rows)} match(es):\n" + "\n".join(lines) + extra


def _rows_payload(rows, limit=50) -> dict:
    """Structured result the UI renders as a table of incident cards."""
    items = []
    for r in rows[:limit]:
        items.append({
            "number": r.get("number", ""),
            "priority": r.get("priority", ""),
            "priority_rank": _priority_rank(r),
            "state": r.get("state", ""),
            "open": _is_open(r),
            "short_description": r.get("short_description", ""),
            "assigned_to": r.get("assigned_to", "") or "Unassigned",
            "study": r.get("study", ""),
            "therapeutic_area": r.get("therapeutic_area", ""),
            "opened_at": r.get("opened_at", ""),
        })
    return {
        "type": "incidents",
        "total": len(rows),
        "shown": len(items),
        "items": items,
    }


def _text_payload(text: str) -> dict:
    return {"type": "text", "text": text}


def _count_payload(label: str, rows) -> dict:
    """A headline count plus the matching rows (so UI shows the number AND the list)."""
    payload = _rows_payload(rows)
    payload["type"] = "count"
    payload["label"] = label
    payload["value"] = len(rows)
    return payload


# Fields the user can group by, mapped to friendly labels.
GROUPABLE = {
    "study": "Study",
    "therapeutic_area": "Therapeutic area",
    "assigned_to": "Assigned to",
    "category": "Category",
    "priority": "Priority",
    "state": "State",
}


MONTH_NAMES = {
    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun",
    "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec",
}


def _ym_of(row) -> str:
    """Return YYYY-MM from opened_at for correct chronological ordering."""
    val = str(row.get("opened_at", "")).strip()
    m = re.search(r"^\s*(\d{4})-(\d{2})-\d{2}", val)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m2 = re.search(r"^\s*(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", val)
    if m2:
        year = m2.group(3)
        if len(year) == 2:
            year = "20" + year
        return f"{year}-{m2.group(2).zfill(2)}"
    return ""


def _monthly_chart_payload(field: str, rows, top_n=8) -> dict:
    """Build a stacked-bar dataset: X = month, series = values of `field`."""
    # Collect month buckets (chronological) and series keys.
    months = sorted({_ym_of(r) for r in rows if _ym_of(r)})
    # Count per (month, series).
    series_totals = {}
    cell = {}  # (month, key) -> count
    for r in rows:
        ym = _ym_of(r)
        if not ym:
            continue
        key = (r.get(field) or "").strip() or "(none)"
        series_totals[key] = series_totals.get(key, 0) + 1
        cell[(ym, key)] = cell.get((ym, key), 0) + 1

    # Keep only the top-N series by total; bucket the rest as "Other".
    ranked = sorted(series_totals, key=lambda k: -series_totals[k])
    keep = ranked[:top_n]
    has_other = len(ranked) > top_n
    keys = keep + (["Other"] if has_other else [])

    labels = [f"{MONTH_NAMES.get(ym.split('-')[1], ym)} {ym.split('-')[0]}" for ym in months]
    datasets = []
    for key in keys:
        data = []
        for ym in months:
            if key == "Other" and has_other:
                total = sum(cell.get((ym, k), 0) for k in ranked[top_n:])
            else:
                total = cell.get((ym, key), 0)
            data.append(total)
        datasets.append({"label": key, "data": data})

    return {
        "type": "chart",
        "chart": "bar",
        "stacked": True,
        "field": field,
        "title": f"Incidents per month by {GROUPABLE.get(field, field)}",
        "labels": labels,
        "datasets": datasets,
        "total": len([r for r in rows if _ym_of(r)]),
    }


def _group_payload(field: str, rows, blank_label="(none)") -> dict:
    counts = {}
    for r in rows:
        key = (r.get(field) or "").strip() or blank_label
        counts[key] = counts.get(key, 0) + 1
    # Sort by count desc, then name.
    ordered = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0].lower()))
    return {
        "type": "group",
        "field": field,
        "label": GROUPABLE.get(field, field.title()),
        "total": len(rows),
        "distinct": len(counts),
        "groups": [{"key": k, "count": v} for k, v in ordered],
    }


def known_assignees() -> list[str]:
    names = set()
    for r in DATA["rows"]:
        a = (r.get("assigned_to") or "").strip()
        if a:
            names.add(a)
    return sorted(names)


def known_values(field: str) -> list[str]:
    vals = set()
    for r in DATA["rows"]:
        v = (r.get(field) or "").strip()
        if v and v.lower() != "n/a":
            vals.add(v)
    return sorted(vals)


def _match_therapeutic_area(low: str):
    """Match a therapeutic area mentioned anywhere in the query."""
    for ta in known_values("therapeutic_area"):
        if re.search(r"\b" + re.escape(ta.lower()) + r"\b", low):
            return ta
    return None


def _match_study(msg: str):
    """Match a study id. Study ids look like 77242113UCO3001 (digits+letters+digits)."""
    m = re.search(r"\b(\d{6,}[A-Za-z]{2,}\d{2,})\b", msg)
    if m:
        token = m.group(1).lower()
        for s in known_values("study"):
            if s.lower() == token:
                return s
        return m.group(1)  # return as typed even if not an exact known match
    # Also allow "study <id>" phrasing with partial match.
    m2 = re.search(r"\bstudy\s+([A-Za-z0-9]{5,})\b", msg, re.IGNORECASE)
    if m2:
        token = m2.group(1).lower()
        for s in known_values("study"):
            if token in s.lower():
                return s
    return None


def _match_person(low: str):
    """Find a person from the dataset mentioned anywhere in the query.
    Matches on any name token (e.g. 'akash' matches 'Akash Shukla')."""
    best = None
    best_len = 0
    for full in known_assignees():
        fl = full.lower()
        # Strip trailing employee id in parentheses for token matching.
        base = re.sub(r"\s*\(.*?\)\s*$", "", fl).strip()
        tokens = [t for t in re.split(r"[\s,]+", base) if len(t) > 2]
        for t in tokens:
            if re.search(r"\b" + re.escape(t) + r"\b", low) and len(t) > best_len:
                best = full
                best_len = len(t)
    return best


def help_text() -> str:
    """Build a full capability guide. When data is loaded, use real values from the
    export as examples so the suggestions are directly runnable."""
    loaded = DATA["count"] > 0

    # Pick live examples from the data when available, else sensible placeholders.
    def first(field, fallback):
        if loaded:
            vals = known_values(field)
            if vals:
                return vals[0]
        return fallback

    example_name = "a name"
    if loaded and known_assignees():
        # Use just the first token of a real assignee (e.g. "Gaurav").
        full = known_assignees()[0]
        example_name = re.sub(r"\s*\(.*?\)\s*$", "", full).split()[0]
    example_study = first("study", "77242113UCO3001")
    example_ta = first("therapeutic_area", "Oncology")
    example_number = ""
    if loaded:
        for r in DATA["rows"]:
            if r.get("number"):
                example_number = r["number"]
                break
    example_number = example_number or "SINC0492904"

    header = (
        f"I'm your D4U ({GROUP}) issue assistant. "
        + (f"I have {DATA['count']} incidents loaded from {DATA['source']}.\n"
           if loaded else "Load a ServiceNow CSV/Excel export first (Upload, drag-drop, "
                          "or 'Load latest from Downloads').\n")
    )

    sections = [
        ("GREETINGS & BASICS",
         ["hi / hello / good morning  — I greet back",
          "help  — this guide",
          "summary  — totals, open/closed, unassigned, priority breakdown"]),

        ("STATUS FILTERS",
         ["open incidents  (open = New / In Progress / On Hold)",
          "resolved   |   closed   |   on hold (or 'onhold')   |   in progress (or 'inprogress')",
          "new   |   canceled"]),

        ("PRIORITY FILTERS (P1-P5 or words)",
         ["open P2   |   critical / high / moderate / low / planning",
          "P3 incidents   |   priority 4",
          "(if a priority has no incidents, I'll tell you which ones exist)"]),

        ("PEOPLE / ASSIGNMENT",
         [f"assigned to {example_name}   |   {example_name}'s incidents",
          "unassigned   (open incidents with no owner)",
          f"how many incidents worked by {example_name}"]),

        ("THERAPEUTIC AREA",
         [f"incidents in {example_ta}   |   how many cases in {example_ta}",
          f"open P2 in {example_ta}"]),

        ("STUDY",
         [f"incidents for study {example_study}",
          f"open incidents in study {example_study}"]),

        ("DATES / MONTHS (by opened date)",
         ["cases in november   |   opened in october",
          "works with full or short month names (jan, feb, ...)"]),

        ("COUNTS",
         ["how many open incidents   |   number of P2   |   count unassigned",
          f"how many in {example_ta}   |   count of {example_name}'s incidents"]),

        ("COMBINE FILTERS (mix freely)",
         [f"open P2 in {example_ta} assigned to {example_name}",
          f"unassigned in {example_ta}",
          f"on hold in study {example_study}",
          "resolved in november"]),

        ("BREAKDOWNS / GROUP BY",
         ["group by study   |   cases per study",
          "breakdown by assignee   |   group by category",
          "by therapeutic area   |   group by priority   |   by state"]),

        ("MONTHLY TREND CHARTS",
         ["monthly chart by status",
          "monthly chart by assignee   |   ...by study   |   ...by priority   |   ...by therapeutic area",
          "(also triggers on: trend / graph / plot / over time)"]),

        ("SPECIFIC INCIDENT & SEARCH",
         [f"status of {example_number}   (exact lookup)",
          "any free text  — searches description / number / assignee / category",
          "e.g. 'timeout', 'pipeline', 'refresh'"]),

        ("LISTING",
         ["list all   (full list, no 100-row cap)",
          "add 'list all' to any filter, e.g. 'list all in " + example_ta + "'"]),
    ]

    out = [header]
    for title, items in sections:
        out.append("\n" + title)
        for it in items:
            out.append("  • " + it)
    out.append("\nTip: data is a snapshot of your last export. Re-export from "
               "ServiceNow and reload to refresh.")
    return "\n".join(out)


def summary() -> dict:
    rows = DATA["rows"]
    if not rows:
        return _text_payload("No data loaded yet.")
    total = len(rows)
    open_n = sum(1 for r in rows if _is_open(r))
    closed_n = total - open_n
    by_pri = {}
    for r in rows:
        rank = _priority_rank(r) or "?"
        by_pri[rank] = by_pri.get(rank, 0) + 1
    unassigned = sum(1 for r in rows if _is_open(r) and not (r.get("assigned_to") or "").strip())
    return {
        "type": "summary",
        "source": DATA["source"],
        "loaded_at": DATA["loaded_at"],
        "total": total,
        "open": open_n,
        "closed": closed_n,
        "unassigned_open": unassigned,
        "by_priority": [{"rank": k, "count": v} for k, v in sorted(by_pri.items())],
    }


GREETINGS = {
    "hi", "hii", "hiii", "hello", "hey", "heya", "yo", "hola",
    "good morning", "good afternoon", "good evening", "morning",
    "namaste", "hi there", "hello there", "greetings", "sup", "whats up",
    "what's up", "howdy",
}
THANKS = {"thanks", "thank you", "thx", "ty", "thankyou", "great", "awesome", "cool", "nice"}
BYES = {"bye", "goodbye", "see you", "see ya", "cya", "later"}


def _greeting_reply() -> dict:
    loaded = DATA["count"] > 0
    if loaded:
        extra = f"I've got {DATA['count']} incidents loaded from {DATA['source']}. "
    else:
        extra = "Load a ServiceNow export first (Upload or 'Load latest from Downloads'). "
    return _text_payload(
        "Hi! I'm your D4U (" + GROUP + ") issue assistant. " + extra +
        "Ask me things like 'open incidents', 'how many P2', 'unassigned', "
        "'assigned to <name>', 'summary', or 'status of SINC...'. Type 'help' for more."
    )


def handle_message(msg: str) -> dict:
    low = msg.lower().strip()
    rows = DATA["rows"]

    # --- Social: greetings / thanks / bye (work with or without data) ---
    clean = re.sub(r"[!.?,]+$", "", low).strip()
    if clean in GREETINGS:
        return _greeting_reply()
    if clean in THANKS:
        return _text_payload("You're welcome! Anything else about your D4U incidents?")
    if clean in BYES:
        return _text_payload("Goodbye! Come back anytime to check on D4U incidents.")

    if not msg or low in ("help", "?", "commands", "menu", "options",
                          "what can you do", "what can i ask", "capabilities", "guide"):
        return _text_payload(help_text())
    if not rows:
        return _text_payload(
            "No export loaded yet. Upload a ServiceNow CSV export or click 'Load latest from Downloads'."
        )

    # --- Summary / overview ---
    if low in ("count", "summary", "stats", "overview", "total"):
        return summary()

    # --- Charts / graphs / trends (monthly) ---
    if re.search(r"\b(chart|graph|plot|trend|trends|monthly|per month|over time)\b", low):
        cfield = None
        if re.search(r"\b(status|state)\b", low):
            cfield = "state"
        elif re.search(r"\b(assignee|assigned|person|people|owner)\b", low):
            cfield = "assigned_to"
        elif re.search(r"\b(study|studies)\b", low):
            cfield = "study"
        elif re.search(r"\b(priority|severity)\b", low):
            cfield = "priority"
        elif re.search(r"\b(therapeutic|\bta\b)\b", low):
            cfield = "therapeutic_area"
        elif re.search(r"\b(category|type)\b", low):
            cfield = "category"
        else:
            cfield = "state"  # sensible default for a monthly trend
        if cfield not in DATA["columns"]:
            avail = ", ".join(c for c in GROUPABLE if c in DATA["columns"])
            return _text_payload(f"No '{cfield}' column to chart. Available: {avail}.")
        return _monthly_chart_payload(cfield, rows)

    # --- Group by / breakdown ---
    if re.search(r"\b(group by|groupby|breakdown|break down|per|by study|by assignee|by category|by priority|by state|by ta|by therapeutic)\b", low) \
       or low.startswith("by "):
        # Determine the grouping field from keywords.
        field = None
        if re.search(r"\b(study|studies)\b", low):
            field = "study"
        elif re.search(r"\b(therapeutic|\bta\b)\b", low):
            field = "therapeutic_area"
        elif re.search(r"\b(assignee|assigned|person|people|owner)\b", low):
            field = "assigned_to"
        elif re.search(r"\b(category|type|kind)\b", low):
            field = "category"
        elif re.search(r"\b(priority|severity)\b", low):
            field = "priority"
        elif re.search(r"\b(state|status)\b", low):
            field = "state"
        if field:
            # Only group by columns that actually exist in the data.
            if field not in DATA["columns"]:
                avail = ", ".join(c for c in GROUPABLE if c in DATA["columns"])
                return _text_payload(
                    f"The loaded export has no '{field}' column. "
                    f"I can group by: {avail or 'nothing groupable'}."
                )
            return _group_payload(field, rows)

    # --- Specific incident number (highest priority) ---
    m = re.search(r"\b((?:S?INC)\d{3,})\b", msg, re.IGNORECASE)
    if m:
        num = m.group(1).upper()
        match = [r for r in rows if str(r.get("number", "")).upper() == num]
        if not match:
            match = [r for r in rows if num in str(r.get("number", "")).upper()]
        return _rows_payload(match)

    # --- Detect optional filters that can combine ---
    wants_count = bool(re.search(r"\b(how many|no of|number of|count of|count|total)\b", low))
    wants_all = bool(re.search(r"\b(list all|show all|all of them|everything|full list|list|show me all)\b", low)) \
                or "all" in low.split()
    # Use word boundaries so "opened" does not trigger the "open" filter.
    wants_open = bool(re.search(r"\b(open|active|pending)\b", low))
    wants_closed = bool(re.search(r"\b(resolved|closed|done)\b", low))
    wants_unassigned = "unassigned" in low or "no owner" in low

    # Specific state words (match with or without spaces, e.g. "on hold"/"onhold").
    STATE_WORDS = {
        "On Hold": ["on hold", "onhold", "hold"],
        "In Progress": ["in progress", "inprogress", "progress", "wip"],
        "New": ["new"],
        "Resolved": ["resolved"],
        "Closed": ["closed"],
        "Canceled": ["canceled", "cancelled"],
    }
    wanted_state = None
    for state_name, words in STATE_WORDS.items():
        if any(re.search(r"\b" + re.escape(w) + r"\b", low) for w in words):
            wanted_state = state_name
            break

    # Month / date detection (e.g. "cases in november", "opened in oct").
    MONTHS = {
        "january": "01", "jan": "01", "february": "02", "feb": "02",
        "march": "03", "mar": "03", "april": "04", "apr": "04",
        "may": "05", "june": "06", "jun": "06", "july": "07", "jul": "07",
        "august": "08", "aug": "08", "september": "09", "sep": "09", "sept": "09",
        "october": "10", "oct": "10", "november": "11", "nov": "11",
        "december": "12", "dec": "12",
    }
    wanted_month = None
    for name, num in MONTHS.items():
        if re.search(r"\b" + name + r"\b", low):
            wanted_month = num
            break

    # Priority detection across P1-P5 and the word labels ServiceNow uses.
    PRIORITY_WORDS = {
        "1": ["p1", "priority 1", "critical"],
        "2": ["p2", "priority 2", "high"],
        "3": ["p3", "priority 3", "moderate", "medium"],
        "4": ["p4", "priority 4", "low"],
        "5": ["p5", "priority 5", "planning"],
    }
    wanted_priority = None
    for rank, words in PRIORITY_WORDS.items():
        if any(re.search(r"\b" + re.escape(w) + r"\b", low) for w in words):
            wanted_priority = rank
            break

    # Person detection: prefer the smart matcher (handles "Akash" -> "Akash Shukla
    # (702531994)"). Fall back to literal text after "assigned to".
    person = _match_person(low)
    person_label = None
    if person:
        person_label = person
    elif "assigned to" in low:
        tail = low.split("assigned to", 1)[1].strip()
        if tail and tail != "me":
            person = tail
            person_label = tail

    # Therapeutic area and study filters (only if those columns exist).
    ta = _match_therapeutic_area(low) if "therapeutic_area" in DATA["columns"] else None
    study = _match_study(msg) if "study" in DATA["columns"] else None

    # Start from all rows and apply whichever filters were detected.
    filtered = rows
    applied = []

    if person:
        pl = person.lower()
        filtered = [r for r in filtered if pl in str(r.get("assigned_to", "")).lower()]
        applied.append(f"assigned to '{person_label}'")
    if wanted_state:
        filtered = [r for r in filtered if str(r.get("state", "")).strip().lower() == wanted_state.lower()]
        applied.append(wanted_state)
    if wanted_month:
        filtered = [r for r in filtered if _month_of(r) == wanted_month]
        applied.append(f"opened in month {wanted_month}")
    if ta:
        tal = ta.lower()
        filtered = [r for r in filtered if tal == str(r.get("therapeutic_area", "")).lower()]
        applied.append(f"TA '{ta}'")
    if study:
        sl = study.lower()
        filtered = [r for r in filtered if sl == str(r.get("study", "")).lower()]
        applied.append(f"study '{study}'")
    if wants_unassigned:
        filtered = [r for r in filtered if not (r.get("assigned_to") or "").strip()]
        applied.append("unassigned")
    if wanted_priority:
        # If this priority doesn't exist in the data at all, say so clearly.
        present = {_priority_rank(r) for r in rows}
        if wanted_priority not in present:
            have = ", ".join("P" + p for p in sorted(x for x in present if x))
            return _text_payload(
                f"There are no P{wanted_priority} incidents in this export. "
                f"The priorities present are: {have}."
            )
        filtered = [r for r in filtered if _priority_rank(r) == wanted_priority]
        applied.append("P" + wanted_priority)
    # Generic open/closed only when a specific state wasn't already chosen.
    if not wanted_state:
        if wants_open and not wants_closed:
            filtered = [r for r in filtered if _is_open(r)]
            applied.append("open")
        elif wants_closed and not wants_open:
            filtered = [r for r in filtered if not _is_open(r)]
            applied.append("resolved/closed")

    any_filter = bool(person or ta or study or wanted_state or wanted_month or wants_unassigned or wanted_priority or wants_open or wants_closed)

    if any_filter:
        label = "Incidents " + (", ".join(applied) if applied else "")
        if wants_count:
            return _count_payload(label.strip(), filtered)
        limit = len(filtered) if wants_all else 100
        return _rows_payload(filtered, limit=limit)

    # --- List everything ---
    if wants_all or low.startswith("list") or "recent" in low or "show me" in low:
        if wants_count:
            return _count_payload("Total incidents", rows)
        limit = len(rows) if wants_all else 100
        return _rows_payload(rows, limit=limit)

    # --- Fallback: free-text search ---
    res = [r for r in rows if low in " ".join(
        str(r.get(f, "")) for f in ("short_description", "number", "assigned_to", "category")
    ).lower()]
    if res:
        if wants_count:
            return _count_payload(f"Incidents matching '{msg}'", res)
        return _rows_payload(res, limit=100)

    # Nothing matched: help the user instead of a dead end.
    sample = ", ".join(known_assignees()[:8])
    return _text_payload(
        "I couldn't match that to a filter. Try things like:\n"
        "• how many incidents worked by Akash\n"
        "• open P1\n"
        "• unassigned\n"
        "• resolved\n"
        "• summary\n"
        + (f"\nPeople in this export include: {sample}…" if sample else "")
    )


# --------------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------------- #
@app.get("/", response_class=HTMLResponse)
async def index():
    return HTMLResponse(INDEX_HTML.replace("__GROUP__", GROUP))


@app.get("/status")
async def status():
    return {
        "loaded": DATA["count"] > 0,
        "count": DATA["count"],
        "source": DATA["source"],
        "loaded_at": DATA["loaded_at"],
        "group": GROUP,
        "downloads": str(DOWNLOADS),
    }


@app.post("/upload")
async def upload(file: UploadFile = File(...)):
    data = await file.read()
    rows, err = load_bytes(file.filename, data)
    if err:
        return JSONResponse({"ok": False, "error": err}, status_code=400)
    if not rows:
        return JSONResponse({"ok": False, "error": "No rows found in the file."}, status_code=400)
    set_dataset(rows, file.filename)
    return {"ok": True, "count": DATA["count"], "columns": DATA["columns"]}


@app.post("/load-latest")
async def load_latest():
    latest = find_latest_download()
    if not latest:
        return JSONResponse(
            {"ok": False, "error": f"No CSV/Excel files found in {DOWNLOADS}."},
            status_code=404,
        )
    try:
        data = latest.read_bytes()
    except Exception as exc:
        return JSONResponse({"ok": False, "error": f"Could not read {latest.name}: {exc}"}, status_code=400)
    rows, err = load_bytes(latest.name, data)
    if err:
        return JSONResponse({"ok": False, "error": err}, status_code=400)
    if not rows:
        return JSONResponse({"ok": False, "error": f"No rows found in {latest.name}."}, status_code=400)
    set_dataset(rows, latest.name)
    return {"ok": True, "count": DATA["count"], "file": latest.name}


@app.post("/chat")
async def chat(request: Request):
    body = await request.json()
    message = (body.get("message") or "").strip()
    return JSONResponse({"result": handle_message(message)})


# --------------------------------------------------------------------------- #
# HTML
# --------------------------------------------------------------------------- #
INDEX_HTML = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>D4U Tracker | SAM - Issue Chatbot</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root{--bg:#0f172a;--panel:#1e293b;--accent:#2563eb;--muted:#64748b;--text:#e2e8f0;}
  *{box-sizing:border-box;}
  body{margin:0;font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--text);height:100vh;display:flex;flex-direction:column;}
  header{padding:12px 20px;background:var(--panel);display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #334155;gap:12px;flex-wrap:wrap;}
  header h1{font-size:15px;margin:0;font-weight:600;}
  .pill{font-size:12px;padding:4px 10px;border-radius:999px;}
  .pill.on{background:#14532d;color:#86efac;}
  .pill.off{background:#450a0a;color:#fca5a5;}
  button{cursor:pointer;border:none;border-radius:8px;font-size:13px;}
  .bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:10px 16px;background:#0b1220;border-bottom:1px solid #334155;}
  .bar button{background:var(--accent);color:#fff;padding:8px 12px;}
  .bar button.ghost{background:#1e293b;border:1px solid #334155;color:var(--text);}
  .bar label{font-size:12px;color:var(--muted);}
  #drop{border:1px dashed #475569;border-radius:8px;padding:8px 12px;font-size:12px;color:var(--muted);}
  #drop.hot{border-color:var(--accent);color:var(--text);}
  #log{flex:1;overflow-y:auto;padding:18px;display:flex;flex-direction:column;gap:12px;}
  .msg{max-width:92%;padding:10px 14px;border-radius:12px;white-space:pre-wrap;line-height:1.45;font-size:14px;}
  .user{align-self:flex-end;background:var(--accent);color:#fff;border-bottom-right-radius:3px;max-width:82%;}
  .bot{align-self:flex-start;background:var(--panel);border-bottom-left-radius:3px;width:92%;}
  .count{font-size:12px;color:var(--muted);margin-bottom:8px;}
  table.inc{width:100%;border-collapse:collapse;font-size:13px;}
  table.inc th{text-align:left;color:var(--muted);font-weight:600;padding:6px 8px;border-bottom:1px solid #334155;position:sticky;top:0;background:var(--panel);}
  table.inc td{padding:7px 8px;border-bottom:1px solid #22304a;vertical-align:top;}
  table.inc tr:hover td{background:#17233b;}
  .num{font-family:monospace;color:#93c5fd;white-space:nowrap;}
  .badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;white-space:nowrap;}
  .p1{background:#7f1d1d;color:#fecaca;}
  .p2{background:#7c2d12;color:#fed7aa;}
  .p3{background:#713f12;color:#fde68a;}
  .p4{background:#064e3b;color:#a7f3d0;}
  .p5{background:#1e3a5f;color:#bae6fd;}
  .pX{background:#334155;color:#cbd5e1;}
  .st{font-size:12px;color:#cbd5e1;white-space:nowrap;}
  .st.closed{color:#64748b;}
  .tablewrap{max-height:52vh;overflow:auto;border:1px solid #334155;border-radius:8px;}
  .tiles{display:flex;gap:10px;flex-wrap:wrap;}
  .tile{background:#0b1220;border:1px solid #334155;border-radius:10px;padding:10px 16px;min-width:96px;}
  .tile .n{font-size:22px;font-weight:700;}
  .tile .l{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;}
  .meta{font-size:11px;color:var(--muted);margin-top:8px;}
  .chips{display:flex;gap:8px;flex-wrap:wrap;padding:0 14px 10px;background:var(--panel);}
  .chip{background:#0b1220;border:1px solid #334155;color:var(--text);padding:6px 10px;border-radius:999px;font-size:12px;}
  #composer{display:flex;gap:10px;padding:14px;background:var(--panel);border-top:1px solid #334155;}
  #composer input{flex:1;padding:12px 14px;border-radius:8px;border:1px solid #334155;background:#0b1220;color:var(--text);font-size:14px;}
  #composer button{background:var(--accent);color:#fff;padding:0 20px;}
  .err{color:#fca5a5;font-size:12px;}
  a{color:#60a5fa;}
</style>
</head>
<body>
<header>
  <h1>D4U Tracker | SAM — __GROUP__</h1>
  <span id="state" class="pill off">No data loaded</span>
</header>

<div class="bar">
  <button onclick="document.getElementById('fileInput').click()">Upload CSV/Excel</button>
  <input id="fileInput" type="file" accept=".csv,.xlsx,.xls,.txt" style="display:none" onchange="uploadFile(this.files[0])"/>
  <button class="ghost" onclick="loadLatest()">Load latest from Downloads</button>
  <span id="drop">…or drag a ServiceNow export here</span>
  <span id="err" class="err"></span>
</div>

<div id="log"></div>

<div class="chips">
  <button class="chip" onclick="quick('summary')">summary</button>
  <button class="chip" onclick="quick('open incidents')">open incidents</button>
  <button class="chip" onclick="quick('open P1')">open P1</button>
  <button class="chip" onclick="quick('unassigned')">unassigned</button>
  <button class="chip" onclick="quick('list all')">list all</button>
  <button class="chip" onclick="quick('monthly chart by status')">chart: status</button>
  <button class="chip" onclick="quick('monthly chart by assignee')">chart: assignee</button>
  <button class="chip" onclick="quick('monthly chart by study')">chart: study</button>
  <button class="chip" onclick="quick('help')">help</button>
</div>

<form id="composer" onsubmit="return send(event)">
  <input id="box" autocomplete="off" placeholder="Ask about your exported incidents… (type 'help')"/>
  <button type="submit">Send</button>
</form>

<script>
function esc(s){ const d=document.createElement('div'); d.textContent = s==null?'':String(s); return d.innerHTML; }

function priBadge(rank, label){
  const cls = ['1','2','3','4','5'].includes(String(rank)) ? 'p'+rank : 'pX';
  return '<span class="badge '+cls+'">'+esc(label||('P'+rank))+'</span>';
}

function addMsg(role, text){
  const d = document.createElement('div');
  d.className = 'msg ' + (role==='user'?'user':'bot');
  d.textContent = text;
  const log = document.getElementById('log');
  log.appendChild(d); log.scrollTop = log.scrollHeight;
  return d;
}

function addNode(html){
  const d = document.createElement('div');
  d.className = 'msg bot';
  d.innerHTML = html;
  const log = document.getElementById('log');
  log.appendChild(d); log.scrollTop = log.scrollHeight;
  return d;
}

function renderIncidents(res){
  if(res.total === 0) return addMsg('bot', 'No matching incidents in the loaded export.');
  let rows = res.items.map(it => {
    const stCls = it.open ? 'st' : 'st closed';
    return '<tr>'
      + '<td class="num">'+esc(it.number)+'</td>'
      + '<td>'+priBadge(it.priority_rank, it.priority)+'</td>'
      + '<td><span class="'+stCls+'">'+esc(it.state)+'</span></td>'
      + '<td>'+esc(it.short_description)+'</td>'
      + '<td>'+esc(it.study||'')+'</td>'
      + '<td>'+esc(it.therapeutic_area||'')+'</td>'
      + '<td>'+esc(it.assigned_to)+'</td>'
      + '</tr>';
  }).join('');
  const more = res.shown < res.total ? ' (showing first '+res.shown+')' : '';
  const html = '<div class="count">'+res.total+' incident(s)'+more+'</div>'
    + '<div class="tablewrap"><table class="inc">'
    + '<thead><tr><th>Number</th><th>Priority</th><th>State</th><th>Short description</th><th>Study</th><th>TA</th><th>Assigned to</th></tr></thead>'
    + '<tbody>'+rows+'</tbody></table></div>';
  return addNode(html);
}

function renderSummary(res){
  const pri = res.by_priority.map(p => '<div class="tile"><div class="n">'+p.count+'</div><div class="l">P'+esc(p.rank)+'</div></div>').join('');
  const html =
    '<div class="tiles">'
    + '<div class="tile"><div class="n">'+res.total+'</div><div class="l">Total</div></div>'
    + '<div class="tile"><div class="n">'+res.open+'</div><div class="l">Open</div></div>'
    + '<div class="tile"><div class="n">'+res.closed+'</div><div class="l">Closed</div></div>'
    + '<div class="tile"><div class="n">'+res.unassigned_open+'</div><div class="l">Unassigned</div></div>'
    + '</div>'
    + '<div class="meta">By priority</div><div class="tiles">'+pri+'</div>'
    + '<div class="meta">Source: '+esc(res.source||'')+' · loaded '+esc(res.loaded_at||'')+'</div>';
  return addNode(html);
}

function renderCount(res){
  let head = '<div class="tiles"><div class="tile"><div class="n">'+res.value+'</div><div class="l">'+esc(res.label||'Count')+'</div></div></div>';
  if(res.total > 0){
    const sub = renderIncidents(res);  // appends its own node
    // prepend the count tile into that same node
    sub.insertAdjacentHTML('afterbegin', head);
    return sub;
  }
  return addNode(head + '<div class="count" style="margin-top:8px">No matching incidents.</div>');
}

function renderGroup(res){
  const max = Math.max(1, ...res.groups.map(g => g.count));
  const rows = res.groups.map(g => {
    const pct = Math.round((g.count / max) * 100);
    return '<tr>'
      + '<td>'+esc(g.key)+'</td>'
      + '<td style="width:90px"><div style="background:#2563eb;height:14px;border-radius:3px;width:'+pct+'%;min-width:4px"></div></td>'
      + '<td style="text-align:right;width:60px">'+g.count+'</td>'
      + '</tr>';
  }).join('');
  const html = '<div class="count">'+res.label+' — '+res.distinct+' group(s), '+res.total+' incidents total</div>'
    + '<div class="tablewrap"><table class="inc">'
    + '<thead><tr><th>'+esc(res.label)+'</th><th></th><th style="text-align:right">Count</th></tr></thead>'
    + '<tbody>'+rows+'</tbody></table></div>';
  return addNode(html);
}

const CHART_COLORS = ['#2563eb','#16a34a','#f59e0b','#dc2626','#9333ea','#0891b2','#db2777','#65a30d','#ea580c','#64748b'];

function renderChart(res){
  const node = addNode('<div class="count">'+esc(res.title)+' ('+res.total+' incidents)</div><div style="background:#0b1220;border:1px solid #334155;border-radius:8px;padding:10px"><canvas></canvas></div>');
  const canvas = node.querySelector('canvas');
  if(typeof Chart === 'undefined'){
    node.innerHTML += '<div class="err">Chart library not loaded (no internet for CDN). I can switch to an offline chart if needed.</div>';
    return node;
  }
  const datasets = res.datasets.map((d, i) => ({
    label: d.label,
    data: d.data,
    backgroundColor: CHART_COLORS[i % CHART_COLORS.length],
    borderWidth: 0,
  }));
  new Chart(canvas, {
    type: res.chart || 'bar',
    data: { labels: res.labels, datasets: datasets },
    options: {
      responsive: true,
      plugins: {
        legend: { labels: { color: '#e2e8f0', boxWidth: 12, font: { size: 11 } } },
        tooltip: { mode: 'index', intersect: false },
      },
      scales: {
        x: { stacked: !!res.stacked, ticks: { color: '#cbd5e1' }, grid: { color: '#22304a' } },
        y: { stacked: !!res.stacked, beginAtZero: true, ticks: { color: '#cbd5e1' }, grid: { color: '#22304a' } },
      },
    },
  });
  return node;
}

function renderResult(res){
  if(!res) return addMsg('bot', '(no response)');
  if(res.type === 'incidents') return renderIncidents(res);
  if(res.type === 'summary') return renderSummary(res);
  if(res.type === 'count') return renderCount(res);
  if(res.type === 'group') return renderGroup(res);
  if(res.type === 'chart') return renderChart(res);
  return addMsg('bot', res.text || '');
}

async function refreshStatus(){
  const j = await (await fetch('/status')).json();
  const pill = document.getElementById('state');
  if(j.loaded){
    pill.textContent = j.count + ' incidents · ' + (j.source||'') ;
    pill.className = 'pill on';
  } else {
    pill.textContent = 'No data loaded';
    pill.className = 'pill off';
  }
}

function setErr(t){ document.getElementById('err').textContent = t || ''; }

async function uploadFile(file){
  if(!file) return;
  setErr('');
  const fd = new FormData(); fd.append('file', file);
  const r = await fetch('/upload', {method:'POST', body: fd});
  const j = await r.json();
  if(j.ok){ await refreshStatus(); addMsg('bot', 'Loaded ' + j.count + ' incidents from ' + file.name + ". Ask me anything, or type 'summary'."); }
  else setErr(j.error || 'Upload failed.');
}

async function loadLatest(){
  setErr('');
  const r = await fetch('/load-latest', {method:'POST'});
  const j = await r.json();
  if(j.ok){ await refreshStatus(); addMsg('bot', 'Loaded ' + j.count + ' incidents from ' + j.file + '.'); }
  else setErr(j.error || 'No file found.');
}

function quick(t){ document.getElementById('box').value = t; send(new Event('x')); }

async function send(ev){
  if(ev && ev.preventDefault) ev.preventDefault();
  const box = document.getElementById('box');
  const text = box.value.trim();
  if(!text) return false;
  addMsg('user', text); box.value='';
  const pending = addMsg('bot', '…');
  try{
    const j = await (await fetch('/chat',{method:'POST',headers:{'Content-Type':'application/json'},
      body: JSON.stringify({message:text})})).json();
    pending.remove();
    renderResult(j.result);
  }catch(e){ pending.textContent = 'Error: ' + e; }
  return false;
}

// Drag & drop
const drop = document.getElementById('drop');
['dragover','dragenter'].forEach(e=>document.body.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('hot');}));
['dragleave','drop'].forEach(e=>document.body.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('hot');}));
document.body.addEventListener('drop', ev=>{ const f = ev.dataTransfer.files[0]; if(f) uploadFile(f); });

addMsg('bot', "Welcome. Load a ServiceNow export (Upload or 'Load latest from Downloads'), then ask about your __GROUP__ incidents. Type 'help' for options.");
refreshStatus();
setInterval(refreshStatus, 6000);
</script>
</body>
</html>
"""


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "8766"))
    uvicorn.run(app, host="127.0.0.1", port=port)
