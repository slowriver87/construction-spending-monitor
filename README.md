# Construction Spending Monitor

**For education only. Not financial advice.** Figures come from public government sources and are checked, but may contain errors. Check the original source before relying on it.

A one-page dashboard of U.S. construction spending from the Census Bureau's monthly *Value of Construction Put in Place* release (C30): housing, business structures and public works, what is driving the change, and how this cycle compares with past downturns. The page is `index.html`, served by GitHub Pages.

## Sources

- Census C30 release tables, Table 1 (seasonally adjusted annual rate): [release.xlsx](https://www.census.gov/construction/c30/xlsx/release.xlsx), with [release.pdf](https://www.census.gov/construction/c30/pdf/release.pdf) as fallback.
- Census C30 [release text](https://www.census.gov/construction/c30/current/index.html): significance flags, year to date, release date.
- Census monthly time series ([total](https://www.census.gov/construction/c30/xlsx/totsatime.xlsx), [private](https://www.census.gov/construction/c30/xlsx/privsatime.xlsx), [public](https://www.census.gov/construction/c30/xlsx/pubsatime.xlsx)): annual revisions and a second read.
- Census [release schedule](https://www.census.gov/construction/c30/release.html).
- FRED (TTLCONS, PRRESCONS, PNRESCONS, TLPBLCONS, PRMFGCONS): optional third read.

## How the update works

All numbers on the page come from three one-line JavaScript objects (`DATA`, `NOTES`, `HIST`) between the `// ===== DATA BLOCKS` markers in `index.html`. `HIST` is static.

`.github/workflows/update.yml` runs `scripts/update.py` on the 1st–5th of each month, after the 10:00 a.m. ET release. The script is plain Python with no AI. It:

1. Reads Table 1. If its newest month is not newer than the page, it exits quietly.
2. Reads the release text, the time series, the schedule and (with a short timeout) FRED.
3. Rebuilds `DATA` and adds one automatic release-log entry with numbers only, then rewrites the `DATA` and `NOTES` lines. Nothing else in the page changes.
4. Runs the checks: release month matches, totals tie within $5M, Table 1 agrees with the release text, every revision is listed (moves over 5% also open an issue), only the data lines changed, `HIST` is unchanged, the monthly series has no gaps, and the page script still parses (`node --check`).
5. Writes `data/summary.json`, a small summary other pages can read.

The workflow commits only when something changed and every check passed. If a check fails, nothing is committed, the run fails, and an issue titled "Data update needs attention" is opened or commented on.

### Manual run

Actions → Update data → Run workflow. `dry_run` defaults to true: the run fetches, parses and checks, then writes every value that would change and the automatic note to the job summary, and commits nothing.

Locally:

```
pip install -r requirements.txt
python -m pytest tests
python scripts/update.py --dry-run
```

### Repository settings

- Repository variables `COMMIT_NAME` and `COMMIT_EMAIL` (Settings → Secrets and variables → Actions → Variables) set the commit author. Use the GitHub no-reply address.
- Pages: Settings → Pages → deploy from the default branch, root folder.

## Not automated

Hand-written analysis in the release log and the explanatory text on the page. The explanations get a manual review about once a month.
