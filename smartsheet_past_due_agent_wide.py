import os
import re
import requests
from datetime import datetime, date
import pytz

SMARTSHEET_TOKEN = os.getenv("SMARTSHEET_TOKEN")
SHEET_ID = os.getenv("SMARTSHEET_SHEET_ID")
TIMEZONE = os.getenv("TIMEZONE", "America/Denver")
SLACK_WEBHOOK_URL = os.getenv("SLACK_WEBHOOK_URL", "").strip()

# Treat any of these as "completed"
COMPLETED_VALUES = {"completed", "complete", "done", "closed", "100%"}

BASE_URL = "https://api.smartsheet.com/2.0"


def smartsheet_headers():
    return {
        "Authorization": f"Bearer {SMARTSHEET_TOKEN}",
        "Content-Type": "application/json",
    }


def get_today_local() -> date:
    tz = pytz.timezone(TIMEZONE)
    return datetime.now(tz).date()


def fetch_sheet(sheet_id: str) -> dict:
    url = f"{BASE_URL}/sheets/{sheet_id}"
    r = requests.get(url, headers=smartsheet_headers(), timeout=30)
    r.raise_for_status()
    return r.json()


def build_column_maps(sheet: dict):
    title_to_id = {}
    norm_to_id = {}
    primary_col_id = None

    for c in sheet.get("columns", []):
        title = c["title"]
        cid = c["id"]
        title_to_id[title] = cid
        norm_to_id[title.strip().lower()] = cid
        if c.get("primary"):
            primary_col_id = cid

    return title_to_id, norm_to_id, primary_col_id


def cell_value(row: dict, col_id: int):
    for cell in row.get("cells", []):
        if cell.get("columnId") == col_id:
            if "value" in cell:
                return cell["value"]
            if "displayValue" in cell:
                return cell["displayValue"]
            return None
    return None


def parse_date(val):
    if not val:
        return None
    try:
        s = str(val).replace("Z", "+00:00")
        d = datetime.fromisoformat(s)
        return d.date()
    except Exception:
        return None


def normalize_status(val) -> str:
    if val is None:
        return ""
    return str(val).strip().lower()


def detect_milestone_sets(norm_to_id: dict):
    base_re = re.compile(r"^m(\d+)$")
    sets = []

    for norm_title in norm_to_id.keys():
        m = base_re.match(norm_title)
        if not m:
            continue
        n = m.group(1)
        base = f"m{n}"
        date_key = f"{base} date"
        status_key = f"{base} status"

        if date_key in norm_to_id and status_key in norm_to_id:
            sets.append({
                "label": f"M{n}",
                "title_id": norm_to_id[base],
                "date_id": norm_to_id[date_key],
                "status_id": norm_to_id[status_key],
            })

    sets.sort(key=lambda x: int(x["label"][1:]))
    return sets


def find_past_due_milestones(sheet: dict):
    _, norm_to_id, primary_col_id = build_column_maps(sheet)
    today = get_today_local()

    milestone_sets = detect_milestone_sets(norm_to_id)
    if not milestone_sets:
        raise RuntimeError("No milestone columns found (expected M1 / M1 date / M1 Status pattern).")

    results = []

    for row in sheet.get("rows", []):
        project = cell_value(row, primary_col_id) or "(Unnamed Project)"
        hits = []

        for ms in milestone_sets:
            title = cell_value(row, ms["title_id"]) or ms["label"]
            due_val = cell_value(row, ms["date_id"])
            status_val = cell_value(row, ms["status_id"])

            due_dt = parse_date(due_val)
            if not due_dt:
                continue

            if due_dt < today and normalize_status(status_val) not in COMPLETED_VALUES:
                hits.append({
                    "milestone": title,
                    "label": ms["label"],
                    "due": due_dt.isoformat(),
                    "days": (today - due_dt).days,
                    "status": status_val
                })

        if hits:
            hits.sort(key=lambda x: x["days"], reverse=True)
            results.append({"project": project, "hits": hits})

    results.sort(key=lambda r: r["hits"][0]["days"], reverse=True)
    return results


def format_summary(results):
    if not results:
        return "No past-due incomplete milestones found 🎉"

    total = sum(len(r["hits"]) for r in results)
    lines = [
        f"Past-due milestones (not completed): {total}",
        f"Projects impacted: {len(results)}",
        ""
    ]

    for r in results:
        lines.append(f"• {r['project']}")
        for h in r["hits"]:
            lines.append(
                f"   - {h['milestone']} ({h['label']}) | Due: {h['due']} | "
                f"Overdue: {h['days']}d | Status: {h['status']}"
            )
        lines.append("")

    return "\n".join(lines).strip()


def post_to_slack(text: str):
    if not SLACK_WEBHOOK_URL:
        return
    requests.post(SLACK_WEBHOOK_URL, json={"text": text}, timeout=20)


def main():
    if not SMARTSHEET_TOKEN or not SHEET_ID:
        raise RuntimeError("Missing SMARTSHEET_TOKEN or SMARTSHEET_SHEET_ID.")

    sheet = fetch_sheet(SHEET_ID)
    results = find_past_due_milestones(sheet)
    summary = format_summary(results)

    print(summary)
    post_to_slack(summary)


if __name__ == "__main__":
    main()
