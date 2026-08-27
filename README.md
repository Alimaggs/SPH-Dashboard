# SPH Activity Planner

A dashboard for the Bristol & Beyond Stronger Practice Hub team to view and filter
every activity running in 2026–2027.

The build reads `Data/2026-2027 SPH Activity Master List.xlsx` and produces one
self-contained HTML file at `public/index.html` — no server, no build tooling, no
internet dependency except web fonts. Sevalla serves that file; you can also open
it locally, email it, or drop it on a shared drive.

## Deploying

The site is hosted as a **Sevalla static site** connected to this repo:

| Setting | Value |
|---|---|
| Build command | *(leave empty)* |
| Publish directory | `public` |

`public/index.html` is committed to the repo already built, so Sevalla needs no
build step and no Python in its build image. Every push to `main` triggers a
redeploy.

## Rebuilding after the spreadsheet changes

```bash
python scripts/build_dashboard.py
```

Requires `openpyxl` (`pip install -r requirements.txt`). Output lands in
`public/`. The script prints the row count so you can sanity-check it against
the workbook.

```
Data/…xlsx  ─┐
             ├─ scripts/build_dashboard.py ──▶ public/index.html
src/dashboard.template.html ─┘                        public/robots.txt
```

Edit `src/dashboard.template.html` to change the dashboard itself; never edit
anything in `public/`, it is overwritten on every build.

## Publishing an update

**The easy way, no software needed.** On GitHub, open the `Data` folder, use
*Add file → Upload files* to drop in the new spreadsheet, and commit. A GitHub
Action rebuilds `public/` and commits it back; Sevalla then redeploys within a
minute or two. Watch the rebuild under the repo's *Actions* tab.

The upload does not have to keep the exact filename — the build uses the most
recently modified `.xlsx` in `Data/` if the expected name is missing. Do delete
the old one afterwards, so it is obvious which list is live.

**Locally**, if you have Python:

```bash
python scripts/build_dashboard.py
git add -A
git commit -m "Update activity list"
git push
```

Either way the URL never changes, so a link you have shared with the team keeps
working.

## Keeping it out of search results

The dashboard is not meant to be found by the public. Two measures ship with the
build: a `noindex, nofollow, noarchive` meta tag in the page, and
`public/robots.txt` disallowing all crawlers.

Both are requests, which well-behaved crawlers honour and others ignore. If the
content genuinely must not be seen, put it behind Sevalla's password protection
rather than relying on these. Adding an `X-Robots-Tag: noindex` response header
in Sevalla is a further belt-and-braces step.

### If the Action fails to push

It needs write access to the repo. Check *Settings → Actions → General →
Workflow permissions* is set to **Read and write permissions** — that setting is
a ceiling the workflow cannot raise on its own.

## What the dashboard does

**Views** — Schedule (grouped by month, date in the gutter) and Table (dense,
sortable by any column). Click any activity for the full record: description,
notes, tickets, every HubSpot tag, and a link to the course page.

**Filters** — reporting period, activity type, format, website status, PD
category, brochure category, network, venue, mailing list, plus a "needs
attention" group (workflows not configured, no website link, venue/time to
confirm, CPD bundles, multi-part programmes). Option counts stay live as you
filter. The month bar chart doubles as a month filter.

**Counting** — the header shows both numbers: bookable sessions and distinct
activities. Sessions 2 and 3 of a multi-part CPD programme are one booking, so
they don't count as separate activities. "Count programmes once" at the bottom of
the filter rail hides those rows entirely.

**Sharing a view** — filters are written into the URL, so a filtered view can be
copied out of the address bar and sent to a colleague. (Inside sandboxed
previews the URL stays fixed; the filters still work.)

**Export CSV** — downloads exactly the rows currently on screen.

**Other** — light/dark theme toggle, `/` to jump to search, `Esc` to close the
detail panel, and a print stylesheet that drops the controls and prints the list.

## How spreadsheet columns are interpreted

| Dashboard | Source |
|---|---|
| Website status | Fill colour on the Event Name cell — green `92D050` Live, blue `00B0F0` built but hidden, yellow `FFFF00` not yet processed, orange `FFA500`/`FFC000` later session |
| Activity type / PD category / Network | Columns J/K/L with the `AT`/`PD`/`NW` prefix and the `RP…-26-27` suffix stripped for display; the full tag is shown in the detail panel |
| Format | Derived from Location — `Online`, `TBC` → "Venue TBC", anything else → face-to-face |
| Day of week | Recomputed from the date, because column C is a formula |
| Session number | Parsed from a `- Session N` title suffix |

### Data quirks it handles

- `Working With Babies` and `Working with Babies` are merged into one category.
- Mailing List values are split on commas against a known vocabulary, so
  `Communication, Language and Literacy` stays whole while `Leadership, PSED`
  splits into two.
- The venue filter merges `St Pauls Nursery Community Room`, `Community Room` and
  `St. Paul's Nursery School` into one venue (per the master list reference doc).
  The detail panel always shows the original Location text as typed.
- A blank Network column means "not a network activity", not missing data.

Two things to know when reading the numbers: the venue merge above is a
judgement call baked into `VENUE_ALIASES` in the build script — if `Community
Room` ever means a different building, correct it there. And "Ticket capacity"
sums every session, so a three-part programme counts its capacity three times.
