# SEO weekly theme report

A script that takes the 3-sheet SEO export (rank tracker, Google Search Console and keyword category map)
and does every week what I first did by hand in Excel for Task 1: cleans the data, merges the three
sources, flags keywords with no theme and builds a theme-level report.

## How to run it
I ran it from the terminal in VS Code.

```
pip install -r requirements.txt
python seo_weekly_report.py data/supply_chain_seo_export.xlsx
```

It creates `output/seo_theme_report_<date>.xlsx` with 4 tabs:

| Tab | What's in it |
|---|---|
| Theme report | Impressions by theme and month, % change, clicks change, average rank and a Gaining / Stable / Losing label |
| Merged | One row per keyword x month with theme, priority, source (both / GSC only), rank and GSC data |
| Unmapped keywords | Keywords with no theme, their traffic and a suggested theme to approve |
| Data issues log | What the script found in that week's file and how it handled it |

The `sql/` folder has the same logic in SQL (`clean_merge_theme.sql`). I ran it with SQLite from VS Code:
`python sql/run_sql.py data/supply_chain_seo_export.xlsx`. It gives the same 72 rows and the same theme
results as the script. The data file is not in the repo; put the export in `data/`.

## How I built it
I did Task 1 first in Excel, step by step, and logged every issue and decision. I'm not strong in Python yet,
so **I coded this with an AI assistant**: I used it to translate the logic I had already worked out in Excel
(and in SQL) into a script, and to explain each part so I could follow it. Each step in the code says which
Excel step it replaces. I ran it and checked the results against my Excel version.

Some things I learned in Excel that shaped the script:
- **Keywords didn't match** because of extra spaces and upper/lower case (51 spellings for 18 keywords),
  so everything is normalized (lower + trim) before any join.
- **Blank CTRs weren't getting flagged** at first. The difference was negative for those rows, so it never
  passed the threshold. Using the absolute difference fixed it, and the script does the same.
- **My first merge left 94 keys instead of 72** because one key pointed to the wrong column. Since then I
  check that total clicks after the merge equal GSC clicks. The script runs that check and **stops without
  writing the report** if it fails, so a wrong number never goes out.
- **A pivot turned my dates into month names** ("May" as text), which broke the key. The script keeps
  months as real dates.
- **When I ran the SQL on my PC it failed on the dates.** The export I downloaded had dates in more than one
  format, and pandas couldn't read them with a single format. I changed the date reading to
  `pd.to_datetime(..., format="mixed")` so each value is read on its own, re-ran it and got the same numbers as
  my Excel: 72 rows, clicks matching GSC (13,754), 54 rows in both sources, 18 GSC only and 27 unmapped.
  I applied the same fix in the main script so both read dates the same way.

## Rules and assumptions
- GSC is the source of truth for clicks and impressions. Rank tracker clicks and impressions are estimates, so
  only rank is taken from it.
- Duplicates = same week + keyword + device. A rank tracker takes one reading per week, keyword and device.
- Blank rank = the site didn't show up that week. Kept blank and left out of the average, never 0.
- CTR is recalculated as clicks / impressions. If the original differs by more than 0.001 (or is blank), it's
  replaced; smaller gaps are rounding.
- Months with fewer than 4 weeks in the rank tracker (July in this file) are flagged as low confidence.
- Visibility = GSC impressions, first vs last month. More than +10% = Gaining, less than -10% = Losing.
- Unmapped keywords are reported as their own group. The suggested theme is **not applied**: mapping is a
  business decision, so someone has to approve it in the category map.
- Thresholds are at the top of the script so they can be changed.

## Bonus: running it unattended every week
I'd replace the manual export with the Search Console API (or its BigQuery export) and the rank tracker's API,
and schedule the script to run every Monday (for example with GitHub Actions or a cloud scheduler), writing to a
BigQuery table that feeds the dashboard. The validation step would send an alert instead of just stopping, along
with a short weekly email listing new issues and any new unmapped keywords, so a person approves new mappings
before they reach the report.
