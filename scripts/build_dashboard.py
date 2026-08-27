#!/usr/bin/env python3
"""Build the SPH Activity Dashboard.

Reads the master list workbook, normalises it, and injects the result into
src/dashboard.template.html to produce a single self-contained HTML file.

    python scripts/build_dashboard.py

Re-run this whenever the spreadsheet changes.
"""

from __future__ import annotations

import base64
import json
import re
import sys
from functools import lru_cache
from datetime import date, datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "Data"
LOGO = ROOT / "Design Assets" / "Bristol and Beyond SPH Early Years Logo Medium - Transparent.png"
TEMPLATE = ROOT / "src" / "dashboard.template.html"
# Sevalla serves this directory as the static site's publish directory.
PUBLIC = ROOT / "public"
OUTPUT = PUBLIC / "index.html"

# Reporting period date ranges, per the master list reference document.
PERIODS = {
    "RP1": ("RP1-26-27", "1 Sep 2026", "30 Nov 2026"),
    "RP2": ("RP2-26-27", "1 Dec 2026", "28 Feb 2027"),
    "RP3": ("RP3-26-27", "1 Mar 2027", "31 May 2027"),
    "RP4": ("RP4-26-27", "1 Jun 2027", "31 Aug 2027"),
}

# Event Name cell fill -> publishing pipeline status.
STATUS_BY_FILL = {
    "FFFFFF00": ("new", "Not yet on website"),
    "FF00B0F0": ("built", "Built, not public"),
    "FF92D050": ("live", "Live on website"),
    "FFFFA500": ("later", "Later session"),
    "FFFFC000": ("later", "Later session"),
}

# Mailing List Subscriptions is free text and some values contain commas
# ("Communication, Language and Literacy") while others are comma-separated
# compounds ("Leadership, PSED"). Match greedily against the known vocabulary.
MAILING_VOCAB = [
    "Communication, Language and Literacy",
    "Leadership and Staff Development",
    "EYFS Learning Community",
    "Physical Development",
    "Working with Babies",
    "Spotlight on Twos",
    "Childminders",
    "Leadership",
    "Maths",
    "PSED",
    "SEND",
    "EDI",
    "Local",
]

# Same physical venue, typed three different ways across rows.
VENUE_ALIASES = {
    "st pauls nursery community room": "St Paul's Nursery School",
    "st. paul's nursery school": "St Paul's Nursery School",
    "st paul's nursery school": "St Paul's Nursery School",
    "community room": "St Paul's Nursery School",
}

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]


@lru_cache(maxsize=1)
def find_workbook() -> Path:
    """Locate the master list, tolerating a renamed file.

    The spreadsheet is often replaced by uploading a new one through the
    GitHub website, where the name may not match exactly.
    """
    preferred = DATA_DIR / "2026-2027 SPH Activity Master List.xlsx"
    if preferred.exists():
        return preferred

    candidates = [f for f in DATA_DIR.glob("*.xlsx") if not f.name.startswith("~$")]
    if not candidates:
        sys.exit(f"No .xlsx file found in {DATA_DIR}.")

    newest = max(candidates, key=lambda f: f.stat().st_mtime)
    if len(candidates) > 1:
        print(f"note: {len(candidates)} spreadsheets in Data/, using the newest")
    print(f"reading {newest.name}")
    return newest


def clean(value) -> str:
    """Trim a cell to a string, collapsing whitespace; None becomes ''."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def strip_tag(value: str, prefix: str) -> str:
    """Turn a HubSpot list tag into a plain label.

    'PD Communication and Language RP3-26-27' -> 'Communication and Language'
    """
    text = clean(value)
    if not text:
        return ""
    if text.startswith(prefix + " "):
        text = text[len(prefix) + 1:]
    return re.sub(r"\s*RP[1-4]-\d{2}-\d{2}$", "", text).strip()


def split_mailing(value: str) -> list[str]:
    text = clean(value)
    out: list[str] = []
    while text:
        for known in sorted(MAILING_VOCAB, key=len, reverse=True):
            if text.lower().startswith(known.lower()):
                out.append(known)
                text = text[len(known):].lstrip(", ").strip()
                break
        else:
            # Unrecognised value: keep it whole rather than mangling it.
            out.append(text)
            break
    return out


def title_key(name: str) -> str:
    """Normalise a title so sessions of one programme share a key."""
    stem = re.sub(r"\s*[-–]\s*Session\s*[123]\s*$", "", name, flags=re.I)
    return re.sub(r"[^a-z0-9]+", " ", stem.lower()).strip()


def parse_time(value: str) -> tuple[str, int | None]:
    """Return (display, start minutes past midnight) for sorting."""
    text = clean(value)
    match = re.match(r"(\d{1,2})(?::(\d{2}))?\s*(am|pm)", text, flags=re.I)
    if not match:
        return text, None
    hour = int(match.group(1)) % 12
    minute = int(match.group(2) or 0)
    if match.group(3).lower() == "pm":
        hour += 12
    return text, hour * 60 + minute


def read_rows() -> list[dict]:
    workbook = openpyxl.load_workbook(find_workbook())
    rows: list[dict] = []

    for sheet in workbook.worksheets:
        period = sheet.title.strip().upper()
        headers = [clean(c.value) for c in next(sheet.iter_rows(min_row=1, max_row=1))]
        index = {h: i for i, h in enumerate(headers)}

        def cell(row, header):
            position = index.get(header)
            return row[position] if position is not None else None

        for excel_row in sheet.iter_rows(min_row=2):
            name_cell = cell(excel_row, "Event Name")
            name = clean(name_cell.value if name_cell else None)
            if not name:
                continue

            fill = name_cell.fill
            rgb = None
            if fill is not None and fill.fill_type == "solid" and fill.fgColor.type == "rgb":
                rgb = fill.fgColor.rgb
            status, status_label = STATUS_BY_FILL.get(rgb, ("unknown", "Not set"))

            raw_date = cell(excel_row, "Activity Date").value
            if isinstance(raw_date, datetime):
                activity = raw_date.date()
            elif isinstance(raw_date, date):
                activity = raw_date
            else:
                activity = None

            session_match = re.search(r"[-–]\s*Session\s*([123])\s*$", name, flags=re.I)
            session = int(session_match.group(1)) if session_match else None

            location = clean(cell(excel_row, "Location").value)
            venue = VENUE_ALIASES.get(location.lower(), location)
            if location.lower() == "online":
                fmt, venue = "Online", ""
            elif location.upper() == "TBC" or not location:
                fmt, venue = "Venue TBC", ""
            else:
                fmt = "Face-to-face"

            tickets_raw = cell(excel_row, "Number of Tickets").value
            try:
                tickets = int(tickets_raw)
            except (TypeError, ValueError):
                tickets = None

            time_display, time_sort = parse_time(cell(excel_row, "Time").value)
            brochure = clean(cell(excel_row, "Brochure Category").value)
            # 'Working With Babies' and 'Working with Babies' are one group.
            if brochure.lower() == "working with babies":
                brochure = "Working with Babies"

            rows.append({
                "rp": period,
                "name": name,
                "stem": re.sub(r"\s*[-–]\s*Session\s*[123]\s*$", "", name, flags=re.I),
                "key": title_key(name),
                "session": session,
                "date": activity.isoformat() if activity else "",
                # Day of Week is a formula cell, so derive the label from the date.
                "day": activity.strftime("%A") if activity else "",
                "month": f"{activity.year}-{activity.month:02d}" if activity else "",
                "monthLabel": f"{MONTHS[activity.month - 1]} {activity.year}" if activity else "Date TBC",
                "time": time_display,
                "timeSort": time_sort if time_sort is not None else 9999,
                "brochure": brochure,
                "location": location,
                "venue": venue,
                "format": fmt,
                "tickets": tickets,
                "type": strip_tag(cell(excel_row, "Activity Type").value, "AT"),
                "pd": strip_tag(cell(excel_row, "Professional Development Category").value, "PD"),
                "network": strip_tag(cell(excel_row, "Network").value, "NW"),
                "mailing": split_mailing(cell(excel_row, "Mailing List Subscriptions").value),
                "cpd": clean(cell(excel_row, "CPD Bundle for Survey").value),
                "workflows": clean(cell(excel_row, "Workflows Configured").value),
                "notes": clean(cell(excel_row, "Notes").value),
                "url": clean(cell(excel_row, "Website URL").value),
                "summary": clean(cell(excel_row, "Short Description").value),
                "status": status,
                "statusLabel": status_label,
                "tags": {
                    "registered": clean(cell(excel_row, "Period Registered").value),
                    "attended": clean(cell(excel_row, "Period Attended").value),
                    "type": clean(cell(excel_row, "Activity Type").value),
                    "pd": clean(cell(excel_row, "Professional Development Category").value),
                    "network": clean(cell(excel_row, "Network").value),
                },
            })

    rows.sort(key=lambda r: (r["date"] or "9999", r["timeSort"], r["name"]))

    # Mark multi-part programmes: a title stem carrying "Session N" rows.
    parts: dict[str, int] = {}
    for row in rows:
        if row["session"]:
            parts[row["key"]] = max(parts.get(row["key"], 0), row["session"])
    for i, row in enumerate(rows, start=1):
        row["id"] = f"a{i:03d}"
        row["parts"] = parts.get(row["key"], 0)
        # A "distinct activity" is every row except sessions 2 and 3 of a
        # multi-part programme, which are covered by the session 1 booking.
        row["counts"] = not (row["session"] and row["session"] > 1)
    return rows


def build() -> None:
    rows = read_rows()
    payload = {
        "generated": datetime.now().strftime("%d %B %Y"),
        "year": datetime.now().year,
        "periods": [
            {"id": key, "tag": tag, "from": start, "to": end}
            for key, (tag, start, end) in PERIODS.items()
        ],
        "activities": rows,
    }

    logo = base64.b64encode(LOGO.read_bytes()).decode("ascii")
    html = TEMPLATE.read_text(encoding="utf-8")
    html = html.replace("__SPH_LOGO__", f"data:image/png;base64,{logo}")
    html = html.replace(
        '"__SPH_DATA__"',
        json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
    )
    if "__SPH_" in html:
        sys.exit("Template still contains an unreplaced placeholder.")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(html, encoding="utf-8")

    # Keep the dashboard out of search results. Three overlapping signals,
    # because each can be ignored independently: the noindex meta tag in the
    # page, robots.txt, and an X-Robots-Tag header via Sevalla's _headers file.
    (PUBLIC / "robots.txt").write_text(
        "User-agent: *\nDisallow: /\n", encoding="utf-8")
    (PUBLIC / "_headers").write_text(
        "/*\n  X-Robots-Tag: noindex, nofollow, noarchive\n", encoding="utf-8")

    counted = sum(1 for r in rows if r["counts"])
    size = OUTPUT.stat().st_size / 1024
    print(f"{len(rows)} rows -> {counted} distinct activities")
    print(f"wrote {OUTPUT.relative_to(ROOT)} ({size:.0f} KB)")


if __name__ == "__main__":
    build()
