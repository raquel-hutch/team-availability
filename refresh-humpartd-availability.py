"""
HUMPartD OOO Availability Grid Generator

Queries Outlook calendars for the HUMPartD distribution list, generates an
HTML availability grid, and uploads it to the Humana Bids Teams channel.

Prerequisites: m365 CLI authenticated (m365 login), Python 3.9+
Usage: python refresh-humpartd-availability.py
"""
import json, datetime, subprocess, sys, os, tempfile

# ── Configuration ────────────────────────────────────────────────────────────
DL_GROUP_ID = "a9269596-444b-4c06-9a00-b027a9980a88"  # HUMPartD DL
TEAM_DRIVE_ID = "b!V0AwNwlRpEW-fDOGOiiSuqsm8B-xxNVKoARQvmDfdE-9LAogW_-bQoiNmmhyq7-g"
UPLOAD_FOLDER = "General/team calendar"
UPLOAD_FILENAME = "humpartd-availability.html"

# Explicit display order
NAME_ORDER = [
    "adam.barnhart@milliman.com",
    "jake.klaisner@milliman.com",
    "Peter.Heinen@milliman.com",
    "ben.corrao@milliman.com",
    "kara.noethiger@milliman.com",
    "Amanda.Howell@milliman.com",
    "Aanika.Mathew@milliman.com",
    "caroline.oconnor@milliman.com",
    "lily.vanhorn@milliman.com",
    "raquel.hutchinson@milliman.com",
]

HOLIDAYS = {
    datetime.date(2026, 9, 7): "Labor Day",
    datetime.date(2026, 11, 26): "Thanksgiving",
    datetime.date(2026, 11, 27): "Day after Thanksgiving",
}

# Humana/Milliman Lessons Learned Check-In dates
MEETING_DAYS = {
    datetime.date(2026, 8, 28): "Lessons Learned Check-In",
    datetime.date(2026, 9, 23): "Lessons Learned Check-In",
    datetime.date(2026, 10, 28): "Lessons Learned Check-In",
    datetime.date(2026, 11, 18): "Lessons Learned Check-In",
}

# Special event spans (box + overlay label)
EVENT_SPANS = [
    (datetime.date(2026, 10, 5), datetime.date(2026, 10, 9), "FORUM"),
]

# Manual OOO overrides (dates not on Outlook calendars)
MANUAL_OOO = {
    "jake.klaisner@milliman.com": [
        ("2026-08-20", "2026-08-23"),
        ("2026-08-27", "2026-08-28"),
        ("2026-09-17", "2026-09-20"),
        ("2026-09-24", "2026-09-25"),
        ("2026-10-15", "2026-10-16"),
        ("2026-10-21", "2026-10-22"),
    ],
}

START = datetime.date(2026, 8, 3)
END = datetime.date(2026, 11, 30)

# ── Helpers ──────────────────────────────────────────────────────────────────
def m365(args: str) -> str:
    result = subprocess.run(
        f"m365 {args}", capture_output=True, text=True, shell=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"m365 command failed: {result.stderr.strip()}")
    return result.stdout.strip()

def m365_json(args: str):
    return json.loads(m365(args + " --output json"))

def biz_days(start, end):
    d = start
    while d <= end:
        if d.weekday() < 5:
            yield d
        d += datetime.timedelta(days=1)

# ── Step 1: Get DL members ──────────────────────────────────────────────────
def get_members():
    print("Fetching HUMPartD distribution list members...")
    members = m365_json(f'entra group member list --groupId "{DL_GROUP_ID}"')
    people = {}
    for m in members:
        email = m["userPrincipalName"]
        people[email] = m["displayName"]
    print(f"  Found {len(people)} members")
    return people

# ── Step 2: Query OOO calendars ─────────────────────────────────────────────
def query_ooo(people):
    print("Querying Outlook calendars for OOO events...")
    emails = list(people.keys())
    ooo = {email: set() for email in emails}

    # Graph getSchedule has a 42-day max, so chunk the date range
    chunk_start = START
    chunk_num = 0
    while chunk_start < END:
        chunk_end = min(chunk_start + datetime.timedelta(days=41), END + datetime.timedelta(days=1))
        chunk_num += 1
        print(f"  Chunk {chunk_num}: {chunk_start} to {chunk_end}")

        body = json.dumps({
            "schedules": emails,
            "startTime": {"dateTime": f"{chunk_start}T00:00:00", "timeZone": "America/Chicago"},
            "endTime": {"dateTime": f"{chunk_end}T00:00:00", "timeZone": "America/Chicago"},
            "availabilityViewInterval": 1440,
        })
        body_file = os.path.join(tempfile.gettempdir(), f"sched_chunk{chunk_num}.json")
        with open(body_file, "w", encoding="utf-8") as f:
            f.write(body)

        raw = m365(
            f'request --url "https://graph.microsoft.com/v1.0/me/calendar/getSchedule"'
            f' --method post --body "@{body_file}" --content-type "application/json" --output json'
        )
        data = json.loads(raw)

        for person in data.get("value", []):
            email = person["scheduleId"]
            if email not in people:
                continue
            for item in person.get("scheduleItems", []):
                if item["status"] != "oof":
                    continue
                s = datetime.datetime.fromisoformat(item["start"]["dateTime"]).date()
                e = datetime.datetime.fromisoformat(item["end"]["dateTime"]).date()
                dur_h = (datetime.datetime.fromisoformat(item["end"]["dateTime"]) -
                         datetime.datetime.fromisoformat(item["start"]["dateTime"])).total_seconds() / 3600
                if dur_h < 20:
                    continue
                d = s
                while d < e:
                    if d.weekday() < 5 and START <= d <= END:
                        ooo[email].add(d)
                    d += datetime.timedelta(days=1)

        chunk_start = chunk_end

    for email in sorted(people.keys(), key=lambda e: people[e]):
        print(f"  {people[email]}: {len(ooo[email])} OOO days")

    # Merge manual overrides
    for email, ranges in MANUAL_OOO.items():
        if email not in ooo:
            continue
        for start_str, end_str in ranges:
            d = datetime.date.fromisoformat(start_str)
            end = datetime.date.fromisoformat(end_str)
            while d <= end:
                if d.weekday() < 5 and START <= d <= END:
                    ooo[email].add(d)
                d += datetime.timedelta(days=1)
        print(f"  {people[email]}: {len(ooo[email])} OOO days (after manual overrides)")

    return ooo

# ── Step 3: Generate HTML ───────────────────────────────────────────────────
def generate_html(people, ooo):
    print("Generating HTML grid...")
    days = list(biz_days(START, END))
    ordered_emails = [e for e in NAME_ORDER if e in people]
    for e in people:
        if e not in ordered_emails:
            ordered_emails.append(e)

    month_spans = []
    current_month = None
    for d in days:
        m = d.strftime("%B %Y")
        if m != current_month:
            month_spans.append([m, 1])
            current_month = m
        else:
            month_spans[-1][1] += 1

    DOW = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    rows = []

    # Precompute event span info
    event_dates = {}  # date -> (label, is_first, is_last)
    for evt_start, evt_end, label in EVENT_SPANS:
        d = evt_start
        while d <= evt_end:
            if d.weekday() < 5 and START <= d <= END:
                event_dates[d] = (label, d == evt_start, d == evt_end)
            d += datetime.timedelta(days=1)

    for idx, email in enumerate(ordered_emails):
        name = people[email]
        is_first_row = (idx == 0)
        is_last_row = (idx == len(ordered_emails) - 1)
        cells = []
        for d in days:
            evt_classes = ""
            if d in event_dates:
                _, is_first_col, is_last_col = event_dates[d]
                if is_first_row:
                    evt_classes += " event-top"
                if is_last_row:
                    evt_classes += " event-bottom"
                if is_first_col:
                    evt_classes += " event-left"
                if is_last_col:
                    evt_classes += " event-right"
            if d in HOLIDAYS and d in ooo[email]:
                cells.append(f'<td class="holiday-ooo{evt_classes}" title="{HOLIDAYS[d]}">OOO</td>')
            elif d in HOLIDAYS:
                cells.append(f'<td class="holiday{evt_classes}" title="{HOLIDAYS[d]}"></td>')
            elif d in MEETING_DAYS and d in ooo[email]:
                cells.append(f'<td class="meeting-ooo{evt_classes}" title="{MEETING_DAYS[d]}">OOO</td>')
            elif d in MEETING_DAYS:
                cells.append(f'<td class="meeting{evt_classes}" title="{MEETING_DAYS[d]}"></td>')
            elif d in ooo[email]:
                cells.append(f'<td class="ooo{evt_classes}">OOO</td>')
            else:
                cells.append(f'<td class="avail{evt_classes}"></td>')
        rows.append(f'<tr><td class="name">{name}</td>{"".join(cells)}</tr>')

    # OOO count row
    count_cells = []
    total = len(people)
    for d in days:
        if d in HOLIDAYS:
            count_cells.append('<td class="holiday">\u2014</td>')
        elif d in MEETING_DAYS:
            c = sum(1 for email in people if d in ooo[email])
            cls = ' class="meeting-ooo"' if c >= total * 0.4 else ' class="meeting"'
            count_cells.append(f'<td{cls}>{c}</td>')
            continue
        else:
            c = sum(1 for email in people if d in ooo[email])
            cls = ' class="hot"' if c >= total * 0.4 else ''
            count_cells.append(f'<td{cls}>{c}</td>')

    rows_html = "\n".join(rows)
    count_html = "".join(count_cells)
    month_headers = "".join(f'<th colspan="{span}" class="month-header">{m}</th>' for m, span in month_spans)
    day_headers = "".join(f'<th>{d.day}<span class="dow">{DOW[d.weekday()]}</span></th>' for d in days)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>HUMPartD Team \u2013 OOO Calendar</title>
<style>
  body {{ font-family: 'Segoe UI', Tahoma, sans-serif; margin: 20px; background: #f9f9f9; }}
  h1 {{ color: #1a3c6e; font-size: 1.4rem; margin-bottom: 4px; }}
  .subtitle {{ color: #666; font-size: 0.85rem; margin-bottom: 16px; }}
  .legend {{ display: flex; gap: 18px; margin-bottom: 12px; font-size: 0.82rem; }}
  .legend span {{ display: inline-flex; align-items: center; gap: 5px; }}
  .swatch {{ width: 16px; height: 16px; border-radius: 3px; display: inline-block; }}
  .container {{ overflow-x: auto; }}
  table {{ border-collapse: collapse; font-size: 0.72rem; }}
  th, td {{ border: 1px solid #ccc; text-align: center; padding: 2px 4px; min-width: 24px; }}
  th {{ background: #1a3c6e; color: #fff; font-weight: 600; position: sticky; top: 0; z-index: 2; }}
  th.name {{ text-align: left; min-width: 140px; }}
  td.name {{ text-align: left; font-weight: 600; white-space: nowrap; }}
  .month-header {{ background: #344e7a; color: #fff; font-weight: 700; font-size: 0.8rem; }}
  .avail   {{ background: #c6efce; }}
  .ooo     {{ background: #f4a4a4; }}
  .meeting {{ background: #ffe699; }}
  .meeting-ooo {{ background: #ffe699; color: #a04000; font-weight: 700; }}
  .holiday {{ background: repeating-linear-gradient(45deg, #c6efce, #c6efce 3px, #a8e6a0 3px, #a8e6a0 6px); }}
  .holiday-ooo {{ background: #f4a4a4; color: #600; font-weight: 700; }}
  .dow {{ font-size: 0.65rem; color: #b0c4de; display: block; }}
  .count-row td {{ font-weight: 700; background: #eef2f7; }}
  .count-row td.hot {{ background: #f4a4a4; color: #600; }}
  .event-box {{ position: relative; }}
  .event-overlay {{
    position: absolute; top: 0; left: 0; right: 0; bottom: 0;
    display: flex; align-items: center; justify-content: center;
    font-size: 2.5rem; font-weight: 900; color: rgba(26,60,110,0.18);
    pointer-events: none; letter-spacing: 0.15em; z-index: 1;
  }}
  .event-left {{ border-left: 3px solid #1a3c6e; }}
  .event-right {{ border-right: 3px solid #1a3c6e; }}
  .event-top {{ border-top: 3px solid #1a3c6e; }}
  .event-bottom {{ border-bottom: 3px solid #1a3c6e; }}
</style>
</head>
<body>
<h1>HUMPartD Team \u2013 OOO Calendar</h1>
<p class="subtitle">Generated from Outlook calendars &bull; {datetime.date.today().strftime("%B %d, %Y")}</p>
<div class="legend">
  <span><span class="swatch" style="background:#c6efce"></span> Available</span>
  <span><span class="swatch" style="background:#f4a4a4"></span> OOO</span>
  <span><span class="swatch" style="background:#ffe699"></span> Lessons Learned Check-In</span>
</div>
<div class="container">
<table>
<thead>
<tr><th class="name" rowspan="2"></th>{month_headers}</tr>
<tr>{day_headers}</tr>
</thead>
<tbody>
{rows_html}
<tr class="count-row"><td class="name" style="color:#c00">OOO count</td>{count_html}</tr>
</tbody>
</table>
</div>
<p style="font-size:0.8rem; color:#888; margin-top:16px;">
  Only full-day OOO calendar events are shown. Hover cells for holiday names.
</p>
<script>
// Overlay labels on event spans
document.addEventListener('DOMContentLoaded', function() {{
  const table = document.querySelector('table');
  const rect = table.getBoundingClientRect();
  const container = document.querySelector('.container');
  container.style.position = 'relative';
  const spans = {json.dumps([(evt_start.isoformat(), evt_end.isoformat(), label) for evt_start, evt_end, label in EVENT_SPANS])};
  const days = {json.dumps([d.isoformat() for d in days])};
  const headerRow = table.querySelectorAll('thead tr:last-child th');
  spans.forEach(function(span) {{
    const startIdx = days.indexOf(span[0]);
    const endIdx = days.indexOf(span[1]);
    if (startIdx < 0 || endIdx < 0) return;
    const firstTh = headerRow[startIdx + 1];
    const lastTh = headerRow[endIdx + 1];
    const tRect = table.getBoundingClientRect();
    const fRect = firstTh.getBoundingClientRect();
    const lRect = lastTh.getBoundingClientRect();
    const rows = table.querySelectorAll('tbody tr');
    const firstRow = rows[0].getBoundingClientRect();
    const lastRow = rows[rows.length - 1].getBoundingClientRect();
    const overlay = document.createElement('div');
    overlay.textContent = span[2];
    const boxH = lastRow.bottom - firstRow.top;
    overlay.style.cssText = 'position:absolute;pointer-events:none;z-index:1;' +
      'font-size:1.6rem;font-weight:900;color:rgba(26,60,110,0.18);' +
      'letter-spacing:0.15em;white-space:nowrap;' +
      'transform-origin:left center;transform:rotate(-55deg);' +
      'left:' + (fRect.left - tRect.left) + 'px;' +
      'top:' + (firstRow.top - tRect.top + boxH * 0.5) + 'px;';
    container.appendChild(overlay);
  }});
}});
</script>
</body>
</html>"""
    return html

# ── Step 4: Upload to SharePoint ────────────────────────────────────────────
def upload_to_sharepoint(html_content):
    print("Uploading to Humana Bids Teams channel...")
    local_path = os.path.join(tempfile.gettempdir(), UPLOAD_FILENAME)
    with open(local_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    upload_url = (
        f"https://graph.microsoft.com/v1.0/drives/{TEAM_DRIVE_ID}"
        f"/root:/{UPLOAD_FOLDER}/{UPLOAD_FILENAME}:/content"
    )
    result = m365(
        f'request --url "{upload_url}" --method put'
        f' --body "@{local_path}" --content-type "text/html"'
    )
    print(f"  Uploaded: {UPLOAD_FILENAME}")
    # Get sharing link
    web_url = (
        f"https://graph.microsoft.com/v1.0/drives/{TEAM_DRIVE_ID}"
        f"/root:/{UPLOAD_FOLDER}/{UPLOAD_FILENAME}?$select=webUrl"
    )
    info = m365_json(f'request --url "{web_url}"')
    print(f"  SharePoint URL: {info.get('webUrl', '(check Teams Files tab)')}")

# ── Main ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    people = get_members()
    ooo = query_ooo(people)
    html = generate_html(people, ooo)

    # Also save locally
    local_out = os.path.join(os.path.dirname(__file__), "..", "projects", UPLOAD_FILENAME)
    with open(local_out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Saved locally: {local_out}")

    upload_to_sharepoint(html)
    print("Done!")
