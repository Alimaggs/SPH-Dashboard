# SPH Activity Master List — Reference Document

Background and data dictionary for `2026-2027 SPH Activity Master List.xlsx`, the Bristol & Beyond Stronger Practice Hub's (SPH) event/activity tracker for the 2026–2027 reporting year. Written to brief a dashboard build — covers what the file is, how it's structured, and what every column and tag means.

## 1. What this file is

The Stronger Practice Hub runs CPD events, webinars and practitioner networks for Early Years settings across the Bristol & Beyond region. Every activity — one-off webinars, recurring network sessions, multi-part CPD programmes, face-to-face study days — is logged as one row in this workbook so the team can track what's running, tag it correctly for HubSpot (email marketing / list membership), and report on it by reporting period.

The workbook has four sheets, one per reporting period:

| Sheet | Reporting Period | Date range |
|---|---|---|
| RP1 | RP1-26-27 | 1 Sep 2026 – 30 Nov 2026 |
| RP2 | RP2-26-27 | 1 Dec 2026 – 28 Feb 2027 |
| RP3 | RP3-26-27 | 1 Mar 2027 – 31 May 2027 |
| RP4 | RP4-26-27 | 1 Jun 2027 – 31 Aug 2027 |

An activity is filed on the sheet matching the calendar date it actually runs on — a multi-part programme that spans two reporting periods (e.g. sessions in November and December) has its rows split across the two relevant sheets.

Each sheet is a proper Excel Table (not just a formatted range), so it supports native sort/filter and auto-extends as rows are added. As of the last update (2 Oct 2026) there are **129 activity rows** across the four sheets (RP1: 33, RP2: 39, RP3: 37, RP4: 20). Counting each multi-part CPD programme once rather than per-session, that's **115 distinct activities/programmes**. These counts include postponed/cancelled rows — filter on Status (column V) to exclude them.

## 2. Column reference (A–V)

| Col | Header | Type | Notes |
|---|---|---|---|
| A | Event Name | Text | Public-facing title. Title Case. See §5 for status colour-coding on this cell. |
| B | Activity Date | Date | UK format `dd/mm/yyyy`. One row per date the activity runs on (a 3-part programme = 3 rows). |
| C | Day of Week | Formula | `=IF(B{n}="","",TEXT(B{n},"dddd"))` — auto-derived, never entered manually. |
| D | Time | Text | Free text, e.g. `"4:00 pm - 6:00 pm"`. |
| E | Brochure Category | Text | Free-text grouping category used for the public brochure/website listing. See §4. |
| F | Location | Text | `"Online"` for webinars, or a venue name for face-to-face. |
| G | Number of Tickets | Number | Capacity. Typical defaults: 80 for online webinars, 24 for face-to-face events at partner venues — but varies by venue/room (e.g. 6, 16, 20, 40, 100, 500 all appear). **Can also contain the text `"External"`** for externally-booked events with unknown capacity, or `"TBC"` — so don't assume this column is always numeric. |
| H | Period Registered | Text | HubSpot list tag, pattern `"Registered Events {RP}-26-27"`. |
| I | Period Attended | Text | HubSpot list tag, pattern `"Attended Events {RP}-26-27"`. |
| J | Activity Type | Text | HubSpot list tag. See §3. |
| K | Professional Development Category | Text | HubSpot list tag ("PD Category"). See §3. |
| L | Network | Text | HubSpot list tag — **only populated for Network-type activities** (see §3). Blank for webinars/events. |
| M | Mailing List Subscriptions | Text | Free-text grouping, usually mirrors Brochure Category (E) but not always identical in wording/casing — see §6. |
| N | CPD Bundle for Survey | Text | `"Yes"` / `"No"` / blank. Whether the activity counts toward a CPD bundle. |
| O | Workflows Configured | Text | `"Yes"` / `"No"`. Whether the HubSpot workflow automation has been set up for this row yet. |
| P | Notes | Text | Free text — usually presenter/trainer credit, delivery partner, or logistics notes. |
| Q | Website URL | Text (hyperlinked) | Link to the live course page on beyth.co.uk. Blank until the event page/venue is confirmed. |
| R | Short Description | Text | 2–4 sentence public-facing summary, written from the full event brief. |
| S | Regional/Local | Dropdown | `"Regional"` (SPH-wide) or `"Local"` (one area only, e.g. Bristol or North Somerset). |
| T | Registration | Dropdown | `"Website"` (booked via beyth.co.uk) or `"External"` (booked on a third-party platform, e.g. Eventbrite, iLearn, Bristol Early Years calendar). |
| U | Processed for Reports | Boolean | `TRUE`/`FALSE` checkbox, default `FALSE`. Ticked manually once the event's data has been pulled into reporting. May be stored as `=FALSE()`/`=TRUE()` formulas — treat as booleans. |
| V | Status | Dropdown | `"Scheduled"` / `"Postponed"` / `"Cancelled"`. Default `Scheduled`. Postponed/cancelled rows stay in the sheet (with a dated note in column P) rather than being deleted — **exclude them from activity counts and "upcoming" views**. |

## 3. HubSpot tagging system (columns H–L)

These columns exist to drive HubSpot list membership and email segmentation. Every tag is suffixed with the reporting period, e.g. `RP1-26-27`, so the same category can be filtered by period.

**Activity Type (J)** — one of three, always prefixed `AT`:
- `AT Event {RP}-26-27` — face-to-face session
- `AT Webinar {RP}-26-27` — online session
- `AT Network {RP}-26-27` — a recurring practitioner network meeting (online or face-to-face)

**PD Category (K)** — the personal development / CPD subject area, always prefixed `PD`. Values in use this year: `PD Mathematics`, `PD Communication and Language`, `PD Early Literacy`, `PD PSED`, `PD Physical Development`, `PD SEND`, `PD Other`.

**Network (L)** — only set when Activity Type is `AT Network`; blank otherwise. Prefixed `NW`. Values in use this year: `NW Working with Babies`, `NW Spotlight on Twos`, `NW Childminders`, `NW SEND`, `NW Equality, Diversity and Inclusion`, `NW Communication, Language and Literacy`, `NW PSED`, `NW Maths`, `NW Leadership and Staff Development`, `NW EYFS Learning Community` *(defined in the tag list but not currently used on any row)*, `NW Physical Development` *(defined but not currently used)*.

**Period Registered / Period Attended (H/I)** — simple per-period tags marking that someone registered for or attended an event; not activity-specific.

## 4. Brochure Category & Mailing List Subscriptions (columns E & M)

Unlike columns H–L, these are **not** structured HubSpot tags — they're free-text grouping labels used for the public brochure listing (E) and for deciding which mailing list segment receives promotion (M). They usually match each other but are typed independently, so wording and casing can drift (see §6).

Values currently in use:
`Childminders` · `Communication, Language and Literacy` · `EDI` · `EYFS Learning Community` · `Leadership and Staff Development` · `Local` · `Maths` · `PSED` · `Physical Development` · `SEND` · `Spotlight on Twos` · `Working With Babies` / `Working with Babies` (casing varies) · plus two compound values that appear only in Mailing List: `Leadership, PSED` and `Working with Babies, Leadership` — used when an activity should reach two mailing segments at once.

Note: "Leadership" appears as a Mailing List value on its own (distinct from "Leadership and Staff Development", which is the paired Brochure Category value for most leadership webinars).

## 5. Event Name status colour-coding (column A cell fill)

The fill colour on the Event Name cell tracks where the activity is in the publishing pipeline — it's a manual signal, not derived from any other column:

| Colour | Hex | Meaning |
|---|---|---|
| 🟡 Yellow | `FFFF00` | Just added to the spreadsheet — not yet processed onto the website. |
| 🔵 Blue | `00B0F0` | Page built/processed in WordPress but not yet made public (hidden). |
| 🟢 Green | `92D050` | Live on the website. |
| 🟣 Purple | `7030A0` | Live on the website, but tickets not yet released. |
| 🟠 Orange | `FFC000` | Session 2 or 3 of a multi-part programme — see §7. |

Colour only ever moves forward as the event is processed — yellow → blue → green, or yellow → blue → purple where the page goes live before tickets are released, then purple → green once they are. It's never set backward or overwritten automatically.

Orange sits outside that flow: it marks sessions 2 and 3 of a multi-part programme, which are never published as separate pages, so they never reach green.

## 6. Known data quality quirks (relevant for a dashboard build)

Worth normalising/handling defensively when filtering programmatically:

- **Casing inconsistency**: `"Working With Babies"` vs `"Working with Babies"` both appear as Brochure Category / Mailing List values for the same underlying group — a case-sensitive filter will split what should be one group in two.
- **Venue naming inconsistency**: the same physical venue appears under slightly different strings across rows, e.g. `"St Pauls Nursery Community Room"` vs `"St. Paul's Nursery School"` vs plain `"Community Room"`. Not a single canonical venue field.
- **Brochure Category vs Mailing List aren't always identical**: e.g. Brochure `"Leadership and Staff Development"` commonly pairs with Mailing List `"Leadership"` (shortened), not an exact string match — don't assume E == M.
- **Network column (L)** is genuinely blank (not "N/A" or similar) for all non-Network activity types — treat blank as "not a network event", not as missing data.
- **`NW EYFS Learning Community`** and **`NW Physical Development`** are defined tags (per the source HubSpot list file) that currently have zero activities tagged with them this year — they exist in the taxonomy but aren't in the live data.

## 7. Multi-part programmes vs. recurring named series

Two patterns look similar (same title stem, multiple dates) but are handled differently — important for a dashboard to distinguish:

**Genuine multi-part CPD programmes** — one booking covers all sessions; attendees must attend all of them. Session 1 is Yellow with `CPD Bundle = Yes` and `Workflows = Yes`; Sessions 2 and 3 are Orange with `CPD Bundle = Yes` (it tracks the whole programme) and `Workflows = No` (workflow automation only needs configuring once, on session 1). Title suffix `" - Session 1/2/3"`. Examples: *CPD Programme: Getting it Right for Twos/Babies*, *Tuning in to Babies/Twos CPD Programme*, *Behaviour CPD Programme*, *CPD Programme: Using Evidence to Strengthen Your Early Years Leadership*, *EAL in the Early Years (CPD Programme)*.

**Recurring named series** — each date is a standalone, independently-bookable session under a shared series name; no requirement to attend more than one. Every occurrence is Yellow (or further along the pipeline) individually, with its own full tagging. Title suffix is a date/month, not "Session N". Examples: *PSED Network*, *Childminder Network*, *Communication and Language Network*, *EDI Network*, *Maths Network*, *SALT Network sessions*, *Talking About Babies/Twos* sessions, *Leadership Matters Network*, *SEND Network*, *Quality Interactions Toolkit Roadshow*.

When counting "how many activities are we running," the 129-vs-115 distinction in §1 comes directly from this split: 129 = every bookable row; 115 = every row *except* Session 2/3 rows of genuine multi-part programmes (i.e. counting each multi-part programme once).

## 8. Ticket counts — typical defaults (not fixed rules)

- Online webinar: 80 tickets is the common default, but ranges up to 500 for high-profile panel webinars and down to 40 for smaller/niche online sessions.
- Face-to-face at a partner nursery/school: commonly 6–25 tickets depending on room capacity (small "immersion day" visits are as low as 6; study days at larger venues run 20–25).
- Always venue/format-specific — never assume a fixed number without checking what was actually specified for that event.

## 9. Row lifecycle notes

- Rows are always kept in date order within each sheet.
- "Save the date" placeholders (venue not yet confirmed) use Location `"TBC"` and Time `"TO ADD"`, with tickets held but not yet released.
- A cross-reference file, `Things to Check with Anna.txt`, sits alongside the workbook logging open questions/follow-ups per event that don't belong in the spreadsheet itself (missing bios, photos, booking-cutoff exceptions, etc.).
