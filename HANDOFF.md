# Construction Spending Monitor: move to GitHub

Brief for Claude Code. Read this first, then `index.html`.

## Goal

Host this dashboard on GitHub Pages and refresh it with a GitHub Action running a plain Python script. No AI in the refresh: the script fetches the Census construction spending release (C30), rebuilds the data blocks in `index.html`, runs the checks below, and commits only if every check passes. This follows the same pattern as the owner's Public Debt Monitor repo.

## Rules

- The repo is **public**. Put no personal information in any file, commit, comment or workflow: no names, no email addresses, no links to claude.ai. Any alert email goes in a GitHub Actions secret, never in the code.
- Commits use the owner's GitHub no-reply address (`<id>+<username>@users.noreply.github.com`). Before the first commit, confirm it with `git config user.email`. In the workflow, read the name and address from repository variables `COMMIT_NAME` and `COMMIT_EMAIL`. Never write the address into a file.
- Do not change the page's design, CSS, markup or explanatory text. The script changes only the three data lines (see below). A check enforces this.
- Keep the "For education only. Not financial advice." note directly under the header.

## The page

`index.html` is one self-contained page. All numbers come from three one-line JavaScript objects between the markers `// ===== DATA BLOCKS` and `// ===== END DATA BLOCKS =====` in the only `<script>`:

1. `const DATA = {...};` holds the monthly series and the latest-release details.
2. `const NOTES = [...];` holds the release log, one entry per data month.
3. `const HIST = {...};` holds the past-cycle data for the "Then vs. now" chart. It is static, so **never change it**.

Everything else (tiles, four charts, sector table, release log) is drawn from these objects when the page loads. The script replaces each object's line with compact JSON (`json.dumps(obj, ensure_ascii=False, separators=(',',':'))`).

### DATA shape (all dollar values in millions, seasonally adjusted annual rate, integers)

- `months`: `[["YYYY-MM", total, res, nonres, pub, mfg], ...]`, sorted, no gaps, from 2022-01. `res` is private residential, `nonres` is private nonresidential, `pub` is total public, and `mfg` is private manufacturing. `res + nonres + pub` equals `total` within $5M (rounding).
- `latest`:
  - `month`: "YYYY-MM", the newest data month.
  - `status`: "preliminary".
  - `released`: release date as text, e.g. "Oct 1, 2026".
  - `next`: e.g. "Nov 2, 2026 (Sep data)".
  - `moe`: `{total, res, nonres, pub}`. Each is true if that month-over-month change is statistically significant (no asterisk in the release text) and false if it has an asterisk.
  - `ytd`: `{cur, prior, label}`, year-to-date sums in **billions** (one decimal), with a label like "Jan–Aug 2026 vs Jan–Aug 2025" (en dash).
  - `sectorNote`: one sentence. Template: "Selected categories from the {Month} release. Census counts data centers under Office." Replace the current text with this template on the first run; the current text claims a cause no source supports.
  - `sectors`: 8 rows `{name, group?, cur, prev, yago}` in millions, in this order: Private nonresidential (group: true), Manufacturing, Power, Office (incl. data centers), Commercial, Public (group: true), Highway and street, Educational.

### NOTES shape

`[{id: "YYYY-MM", released, headline, points: [..], auto}]`. The two existing entries (2026-07, 2026-08) are hand-written; keep them unchanged. The script adds one entry per new data month with `auto: true` (the page labels it "Automatic summary"). It is a fixed template with numbers only and no causes, every number computed in Python:

- `headline`: "Total ${X}B, {±m.m}% m/m{ (not significant)}, {±y.y}% y/y".
- Point 1: total m/m (flag "not significant" when the asterisk is present), y/y, YTD cur vs prior. End with "Nominal dollars."
- Point 2: residential, nonresidential and public m/m, each flagged when not significant.
- Point 3: y/y for Office (incl. data centers), Power and Manufacturing.
- Point 4: revisions to the two prior months in $B, and the count of other months revised.
- Point 5: the then-vs-now reading, computed exactly as the page does (current peak = max total in `months`; index = latest / peak × 100; months since peak; the 2006–10 value at the same month from `HIST.bust2006`).
- Last point: "Checks: …" saying which sources agreed and what was single source.

## Sources and verified formats

Checked on Oct 19, 2026 against the Aug 2026 release (released Oct 1, 2026).

### 1. Release table (primary for the new month and recent revisions)

- Excel: `https://www.census.gov/construction/c30/xlsx/release.xlsx`. PDF fallback: `https://www.census.gov/construction/c30/pdf/release.pdf` (use `pdfplumber`).
- **The Excel layout was not verified from here** (the tool could not read binary files). The PDF's Table 1 was read and is below; the Excel file is published as the same release in Excel. Before writing the parser, download `release.xlsx`, print the sheet names and the first 70 rows of the Table 1 sheet, and save that printout as a test fixture in `tests/`. Then map rows by exact label.
- Table 1 title: "Table 1. Value of Construction Put in Place in the United States, Seasonally Adjusted Annual Rate", "(Millions of dollars. Details may not add to totals due to rounding.)"
- Columns: `Type of Construction | Aug 2026p | Jul 2026r | Jun 2026r | May 2026 | Apr 2026 | Aug 2025 | % from Jul 2026 | % from Aug 2025`. Suffix `p` = preliminary and `r` = revised. Five recent months plus the year-ago month. Numbers have thousands commas.
- Row labels (exact, in order). Indented rows are shown with two leading spaces here. Footnote digits are glued to some labels:

```
Total Construction
Residential
Nonresidential
  Lodging, Office, Commercial, Health care, Educational, Religious, Public safety,
  Amusement and recreation, Transportation, Communication, Power, Highway and street,
  Sewage and waste disposal, Water supply, Conservation and development, Manufacturing
Total Private Construction1
Residential2
  New single family
  New multifamily
Nonresidential
  Lodging, Office, Commercial, Health care, Educational, Religious,
  Amusement and recreation, Transportation, Communication, Power, Manufacturing
Total Public Construction3
Residential
Nonresidential
  Office, Commercial, Health care, Educational, Public safety, Amusement and recreation,
  Transportation, Power, Highway and street, Sewage and waste disposal, Water supply,
  Conservation and development
```

- **Labels repeat across the three blocks** ("Residential", "Nonresidential", "Office", "Power", "Manufacturing", "Educational", "Highway and street" and others). First split the table into blocks at the exact headings "Total Construction", "Total Private Construction" and "Total Public Construction" (after removing a trailing footnote digit from those three headings and from "Residential2"). Then match rows by exact label **inside the right block**. Never use "contains" or a fuzzy match.
- Mapping:
  - `total` ← "Total Construction".
  - `res` ← Private block "Residential" (printed "Residential2").
  - `nonres` ← Private block "Nonresidential".
  - `pub` ← "Total Public Construction".
  - `mfg` ← Private block "Manufacturing".
  - Sector rows: Private nonresidential ← Private "Nonresidential"; Manufacturing, Power, Commercial ← Private block; Office (incl. data centers) ← Private "Office"; Public ← "Total Public Construction"; Highway and street, Educational ← Public block.
  - `prev` = the prior-month column and `yago` = the year-ago column.
- Aug 2026 values, to use as a test fixture. Each line gives Aug 2026p, Jul 2026r, Jun 2026r, May 2026, Apr 2026 and Aug 2025:
  - Total Construction: 2,203,129, 2,184,474, 2,187,350, 2,168,396, 2,168,231, 2,242,024.
  - Private Residential2: 882,270, 872,707, 882,120, 874,705, 882,091, 926,942.
  - Private Nonresidential: 773,010, 764,979, 758,852, 749,441, 745,266, 780,447.
  - Total Public Construction3: 547,849, 546,788, 546,378, 544,250, 540,874, 534,635.
  - Private Manufacturing: 168,209, 168,220, 168,924, 171,381, 175,033, 209,798.
  - Private Power: 165,900, 164,453, …, 151,277.
  - Private Office: 134,119, 128,199, …, 103,300.
  - Private Commercial: 116,448, 116,516, …, 123,038.
  - Public Highway and street: 150,610, 150,505, …, 144,279.
  - Public Educational: 113,073, 112,939, …, 111,676.
  - These all match the values already stored in `index.html`.

### 2. Release text (significance flags, year to date, release date)

- `https://www.census.gov/construction/c30/current/index.html`. Page heading: "Monthly Construction Spending, {Month} {Year}". Release line: "FOR IMMEDIATE RELEASE: {Weekday}, {Month} {D}, {Year}".
- Sentence templates as published. The direction word is "above" or "below", so match both. The asterisk sits directly after the margin's closing parenthesis, before "above"/"below":
  - Total: "Construction spending during {MONTH} {YEAR} was estimated at a seasonally adjusted annual rate of ${N} billion, {N} percent (±{N} percent)* above the revised {MONTH} estimate of ${N} billion."
  - YoY: "The {MONTH} figure is {N} percent (±{N} percent) below the {MONTH} {YEAR} estimate of ${N} billion."
  - YTD: "During the first eight months of this year, construction spending amounted to ${N} billion, {N} percent (±{N} percent) below the ${N} billion for the same period in {YEAR}." The month count is spelled out ("first eight months").
  - Private sentences: "Residential construction was at a seasonally adjusted annual rate of ${N} billion in {MONTH}, {N} percent (±{N} percent)* above …" and "Nonresidential construction was at …". These are in the "Private Construction" section; match them only there.
  - Public: "In {MONTH}, the estimated seasonally adjusted annual rate of public construction spending was ${N} billion, {N} percent (±{N} percent)* above …"
  - Section headings, in order: "Total Construction", "Private Construction", "Public Construction".
- For the Aug 2026 release, the asterisk was present on Total, Residential and Public, and absent on Nonresidential. That gives `moe = {total:false, res:false, nonres:true, pub:false}`, which matches the stored value.
- The rendered page did not show a footnote explaining the asterisk. Census's meaning is "the 90 percent confidence interval includes zero". **Inspect the raw HTML once** to see how the asterisk is marked up (plain `*`, `<sup>`, or an entity) before writing the regex. Then parse the visible text (e.g. BeautifulSoup `get_text(" ")`), not the HTML.

### 3. Full monthly history (annual revision, and second read)

- Monthly, seasonally adjusted time series in Excel: `https://www.census.gov/construction/c30/xlsx/totsatime.xlsx` (total), `privsatime.xlsx` (private), `pubsatime.xlsx` (public), same folder. **Layout not verified from here**; inspect them first and save a fixture as above.
- Census also does an annual revision that can change older months. Its methodology page (`https://www.census.gov/construction/c30/methodology.html`) confirms annual revisions but does not name the month; the previous refresh process expected it with the July 1 release (May data). So do not rely on the calendar: read these three files on **every** run that finds a new release, compare every stored month, update each one that changed and list it. If the files can't be read, update only the months in Table 1 and say so in the summary.

### 4. Release schedule

- `https://www.census.gov/construction/c30/release.html`. Table columns: `Reporting Period | Release Date | Time`. Rows look like "September | November 2 | 10:00 a.m.". The year appears only on some cells ("November | January 4, 2027", "January 2027 | March 1"), so carry the year forward row by row. Releases are at 10:00 a.m. ET on the first business day of the month.

### 5. FRED (optional third read only)

- `https://fred.stlouisfed.org/data/<ID>.txt` for TTLCONS, PRRESCONS, PNRESCONS, TLPBLCONS and PRMFGCONS. It is a text header (Title, Series ID, Units "Millions of Dollars", Last Updated…) then `DATE VALUE` rows, e.g. `2026-08-01 2203129`. FRED usually lags Census by a few hours.
- FRED times out from GitHub runners. Use a short timeout (15 s, one try), never block on it, and record "FRED not reached" in the run summary.

## Fetching

- Send a browser User-Agent on every request, use a 60-second timeout, and retry 3 times with backoff (e.g. 5 s, 15 s, 45 s).
- Fallback order for the release table: `release.xlsx`, then `release.pdf`. If both fail, change nothing and fail the run.
- The release text page has no fallback. If it cannot be read, keep the previous `moe` and `ytd`, mark them in the summary as carried over, and do not write an automatic note for that month.

## Checks before publishing (any failure means no commit, and the run fails)

1. The newest month in Table 1 equals the reference month on the release page, and it is newer than `DATA.latest.month` (otherwise exit quietly: nothing new).
2. Totals tie: `res + nonres + pub` = `total` within $5M, for every month written. Private residential + private nonresidential = "Total Private Construction" within $5M.
3. Table 1 agrees with the release text: total, private residential, private nonresidential and public to $0.1B, and the MoM and YoY percents to 0.1.
4. Revisions: list every revised month (old → new, %) in the run summary. No value may move by more than 5% without being listed in the commit message as well, and a move over 5% also opens an issue so the owner can look.
5. Only the three data lines changed: compare `index.html` before and after with those three lines removed. Any other difference fails the run.
6. `HIST` is byte-identical to before.
7. `months` is sorted, has no gaps or duplicates, and every row has 6 numbers.
8. The page still parses: extract the script and run `node --check`.

## Dry run

The workflow has a manual "Run workflow" with an input `dry_run` (default **true**). A dry run fetches, parses and checks, then writes to the job summary every value that would change (old → new) and the automatic note it would add. It commits nothing.

## Schedule

- Cron `37 15 1-5 * *` and `37 18 1-5 * *` (UTC): the 1st–5th of each month, after the 10:00 a.m. ET release. The script exits quietly when there is nothing new.
- Commit only when something changed and every check passed. The commit message says the data month and lists revisions.
- On a failed check: no commit, fail the workflow, and open (or comment on) an issue titled "Data update needs attention".
- Also write `data/summary.json`: `{month, status, released, next, total, mom, yoy, ytd_cur, ytd_prior, res, nonres, pub, peak_month, pct_below_peak}`. Other pages can read this file later.

## Not automated

- Written analysis in the release log beyond the automatic summary. The owner may add a hand-written note occasionally.
- The explanatory text (glossary, intros, "Reading it"). Plan a manual review about once a month; the current audit note says the next one is due Nov 18, 2026.

## Suggested repo layout

```
index.html                     the page (served by GitHub Pages)
scripts/update.py              fetch, parse, check, rebuild
tests/                         fixtures (release table printout, sample release text) and parser tests
data/summary.json              small summary for other pages
requirements.txt               pinned: requests, openpyxl, pdfplumber, beautifulsoup4
.github/workflows/update.yml
README.md                      what it is, sources, how the update works, the education-only note
```
