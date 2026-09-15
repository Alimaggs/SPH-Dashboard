# SPH Activity Dashboard

A dashboard for the Bristol & Beyond Stronger Practice Hub team to view and filter
every activity running in 2026–2027.

The build reads `Data/2026-2027 SPH Activity Master List.xlsx` and produces one
self-contained HTML file at `public/index.html` — no server, no build tooling and
no network requests at all. Sevalla serves that file; you can also open it
locally, email it, or drop it on a shared drive.

## Deploying

The site is hosted as a **Sevalla static site** connected to this repo. Under
*Build strategy*:

| Setting | Value |
|---|---|
| Build site before publishing | **off** |
| Root directory | `.` |
| Publish directory | `public` |
| Index file | `index.html` |
| Build command / Node version | *(ignored while building is off)* |

`public/index.html` is committed already built, so Sevalla has nothing to build
and needs no Python in its image. Publish directory is relative to the root
directory, and its contents become the site root — so `public/index.html` is
served at `/`, and there is **no need for an index file at the repo root**.
Every push to `main` triggers a redeploy.

## Rebuilding after the spreadsheet changes

```bash
python scripts/build_dashboard.py
```

Requires `openpyxl` (`pip install -r requirements.txt`). Output lands in
`public/`. The script prints the row count so you can sanity-check it against
the workbook.

```
Data/…xlsx  ──────────────┐
src/dashboard.template.html ──├─ scripts/build_dashboard.py ──▶ public/
src/images/*.png ───────────┘
```

Edit `src/dashboard.template.html` to change the dashboard itself; never edit
anything in `public/`, it is overwritten on every build.

The build is deterministic: the same inputs always produce a byte-identical
`public/index.html`, so a rebuild that changes nothing commits nothing. The
"Data updated" date in the footer is the workbook's own last-saved timestamp,
read from its document properties — not the time of the build, which would
otherwise make every rebuild on a new day look like a change.

### Changing the logos

The images are embedded in the page, pre-sized in `src/images/`. Replace the
artwork in `Design Assets`, then regenerate them:

```bash
pip install Pillow
python scripts/prepare_images.py
python scripts/build_dashboard.py
```

Pillow is deliberately **not** in `requirements.txt`. Resizing during the build
made PNG output depend on the installed Pillow and zlib versions, so CI produced
a byte-different file from a local build every time and committed a pointless
rebuild on every push. Sizing happens once, here; the build only base64-encodes
the result.

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

## Security

The live site is **read-only by construction**. It is one static HTML file: no
server-side code, no database, no API, no login, no forms, and no upload path.
Nothing a visitor does can change what anyone else sees — a visitor can edit the
page in their own browser's dev tools, but that lives only in their tab and is
gone on refresh.

Verified against the live site:

| Check | Result |
|---|---|
| `PUT` / `POST` / `DELETE` / `PATCH` | Return 200 but write nothing — the host serves `index.html` whatever the method. A `PUT` to a new path 404s and creates no resource. |
| Spreadsheet, source, scripts, README | Not served — all 404. Only `public/` is published. |
| `/.git/` | Not served. |
| Directory listing | Not available. |
| Outbound requests from the page | None. No `fetch`, `XMLHttpRequest`, `WebSocket`, form or cookie. |
| Browser storage | The light/dark preference only. |

Two consequences worth keeping in mind:

**The GitHub repository is public.** The dashboard is not the exposure — the repo
is. Anyone can download `Data/2026-2027 SPH Activity Master List.xlsx` in full,
including the internal notes and the events not yet published on the website.
Make the repository private if that is not intended; the live site does not
depend on it being public.

**Anyone with the URL can read the dashboard.** There is no authentication. The
`noindex` tag and `robots.txt` keep honest crawlers away, but they are requests,
not access control. If the dashboard genuinely must not be seen by outsiders, it
needs a login in front of it.

### Response headers

`public/_headers` sets a strict Content-Security-Policy plus `nosniff`,
`X-Frame-Options`, `Referrer-Policy`, `Permissions-Policy` and HSTS. The policy
denies everything by default and allows only what this page actually uses: its
own inline script and style, and its own `data:` images and font. The page
fetches nothing at all, so every remote origin is denied — `connect-src` and
the rest fall through to `none`.

`frame-ancestors 'self'` stops other sites embedding the dashboard. **To embed it
in beyth.co.uk**, add that origin to the `CSP` value in `scripts/build_dashboard.py`
and rebuild.

Inline script and style are allowed, because the whole page is one inline block.
That is safe here for a specific reason: no value a visitor controls — URL,
search box, or otherwise — is ever written into the page as markup. Spreadsheet
text is HTML-escaped everywhere it is rendered, and the Website URL column is
restricted to `http`/`https`, so a `javascript:` value typed into the workbook
cannot become a clickable link.

## Keeping it out of search results

The dashboard is not meant to be found by the public. Three measures ship with
the build, because each can be ignored independently:

- a `noindex, nofollow, noarchive` meta tag in the page
- `public/robots.txt` disallowing all crawlers
- `public/_headers`, setting an `X-Robots-Tag: noindex` response header

After the first deploy, confirm the header is actually being applied — Sevalla's
`_headers` support was not covered in the settings documentation, so the file's
expected location is unverified:

```bash
curl -sI https://sph-dashboard.chaoscreated.com/ | grep -i x-robots-tag
```

If nothing comes back, the meta tag and robots.txt still stand on their own.

All three are *requests*. Well-behaved crawlers honour them; others ignore them,
and none of them stop anyone who has the URL. Given the repo is public and the
data includes events not yet published on the website, treat access control as a
separate question — if this genuinely must not be seen, it needs authentication
in front of it, not crawler hints.

### If the Action fails to push

It needs write access to the repo. Check *Settings → Actions → General →
Workflow permissions* is set to **Read and write permissions** — that setting is
a ceiling the workflow cannot raise on its own.

## Design

The interface follows **Material Design 3**.

The colour scheme is not hand-picked: it is generated from the Bristol & Beyond
blue `#224c7f` with Material's own HCT engine (`materialyoucolor`, the **Vibrant**
scheme variant), so every role — primary, containers, the five surface container
levels, outline, and both themes — is real M3 output.

Vibrant is chosen over the default Tonal Spot because it carries the hue into
the *surfaces*, not just the accents: the page sits on `#f6f6ff` and cards on
`#eef0ff` rather than a neutral grey, and the dark theme is a deep navy rather
than charcoal. That tint is what stops an M3 layout reading as flat.

The colour roles are then used rather than merely defined. Each measure in the
summary row takes its own tonal container — primary, secondary, tertiary,
neutral and the live green — and the activity-type chip on every card takes the
secondary container. Vibrant's tertiary is very saturated, so the one tile using
it drops to a paler tone of the same hue; at full strength it shouted.

Nothing coloured falls below 5.9:1 against its own text in either theme.

Type, shape, elevation, state layers and motion all use the published M3 token
values: the type scale (body-medium 14/20 at 0.25px tracking, title-small,
label-medium and so on), the shape scale (4/8/12/16/28 and full), elevation
levels, the 8% hover and 10% pressed state layers, and the standard and
emphasised-decelerate easings.

Components map onto their M3 counterparts: a search bar rather than a text
field, outlined and filled buttons at 40px with full corners, 48px list rows
with state layers, filled cards at 12px, chips at 8px, primary tabs with a 3px
rounded indicator, a navigation drawer with a 16px trailing corner, and a side
sheet with a 16px leading corner.

**Status colours keep the spreadsheet's hues.** M3's error/success semantics say
nothing about a publishing pipeline, so each status keeps the hue of its fill
colour in column A and is expressed as an M3 container pair — container at tone
90 with on-container at tone 30 in light, mirrored in dark. Every pair lands at
about 7.2:1.

### The typeface

M3 specifies Roboto, which is not a system font. Rather than reintroduce a
Google Fonts request, the Latin subset is vendored in `src/fonts/` and embedded
in the page as a data URI. It is a variable font, so one 43 KB file covers both
weights the dashboard uses, and the page still loads in a single request with
the Content-Security-Policy denying every remote origin.

Roboto is licensed under the SIL Open Font License 1.1; `src/fonts/OFL.txt` is
redistributed with it as the licence requires.

## What the dashboard does

**Views** — Schedule (grouped by month, date in the gutter) and Table (dense,
sortable by any column). Click any activity for the full record: description,
notes, tickets, every HubSpot tag, and a link to the course page.

**Filters** — reporting period, activity type, format, regional or local,
registration, website status, PD category, brochure category, network, venue,
mailing list, processed for reports, plus a "needs attention" group (needs
ticketing, no website link, venue/time to confirm). Option counts stay live as
you filter. The month bar chart doubles as a month filter.

Because nearly every activity is regional and booked through the website, only
the exceptions are badged on the schedule cards: a **Local** pill, and a pill
naming the booking platform when it is not the website.

**Counting** — the header shows both numbers: individual sessions and distinct
activities. Sessions 2 and 3 of a multi-part CPD programme are one booking, so
they don't count as separate activities. "Count programmes once" at the bottom of
the filter rail hides those rows entirely.

**Sharing a view** — filters are written into the URL, so a filtered view can be
copied out of the address bar and sent to a colleague.

**Other** — light/dark theme toggle, `/` to jump to search, `Esc` to close the
detail panel, and a print stylesheet that drops the controls and prints the list.

## How spreadsheet columns are interpreted

| Dashboard | Source |
|---|---|
| Website status | Fill colour on the Event Name cell — green `92D050` live, purple `7030A0` live but not yet ticketed, blue `00B0F0` built but hidden, yellow `FFFF00` not yet processed, orange `FFA500`/`FFC000` later session |
| Activity type / PD category / Network | Columns J/K/L with the `AT`/`PD`/`NW` prefix and the `RP…-26-27` suffix stripped for display; the full tag is shown in the detail panel |
| Format | Derived from Location — `Online`, `TBC` → "Venue TBC", anything else → face-to-face |
| Day of week | Recomputed from the date, because column C is a formula |
| Session number | Parsed from a `- Session N` title suffix |
| Tickets | A number where the column holds one. Non-numeric text such as `External` is shown as typed, and those sessions are left out of the Ticket capacity total |
| Regional or local | Column S, folded to `Regional` / `Local`; other wording passes through as typed |
| Registration | Column T, folded to `Website` / `External platform`; a named platform passes through as typed |
| Processed for reports | Column U, an Excel checkbox. Read from the evaluated value, since openpyxl reports the cell as the formula `=FALSE()` |

### Data quirks it handles

- `Working With Babies` and `Working with Babies` are merged into one category.
- Mailing List values are split on commas against a known vocabulary, so
  `Communication, Language and Literacy` stays whole while `Leadership, PSED`
  splits into two.
- The venue filter folds every spelling of St Pauls Nursery Community Room into
  that one name — `Community Room`, `St. Paul's Nursery School`, `St Pauls
  Nursery School` and `St Paul's Nursery School & Children's Centre` all refer to
  it. Add any new spelling to `VENUE_ALIASES` in the build script. The detail
  panel always shows the original Location text as typed.
- A blank Network column means "not a network activity", not missing data.

Two things to know when reading the numbers: the venue merge above is a
judgement call baked into `VENUE_ALIASES` in the build script — if `Community
Room` ever means a different building, correct it there. And "Ticket capacity"
sums every session, so a three-part programme counts its capacity three times.
