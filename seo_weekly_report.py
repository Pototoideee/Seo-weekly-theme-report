"""
SEO weekly theme report
=======================
How I built this: I worked out every step by hand in Excel first (Task 1) and wrote each decision
in my data issues log. Then I moved that same logic here with help from an AI assistant, because
I'm not strong in Python. Each step below says which Excel step it replaces, so you can check it
against the workbook. I ran it and compared the results with my Excel version: same 72 rows, same
clicks total and same theme results.

What it does: clean -> merge -> flag unmapped keywords -> theme report -> data issues log.

Run (from the project folder):
    python seo_weekly_report.py data/supply_chain_seo_export.xlsx
    (on my PC "python" wasn't on PATH, so I used: py seo_weekly_report.py ...)

Output:
    output/seo_theme_report_<date>.xlsx with 4 tabs:
      Theme report | Merged (keyword x month) | Unmapped keywords | Data issues log
"""
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font

# ------------------------------------------------------------------ settings
# Same rules I used in Excel. Kept here so they can be changed without touching the logic.
CTR_THRESHOLD = 0.001          # CTR difference bigger than this = replace with my own calculation
VISIBILITY_THRESHOLD = 0.10    # +/-10% change in GSC impressions = Gaining / Losing, otherwise Stable
FULL_MONTH_WEEKS = 4           # months with fewer weeks in the rank tracker are flagged as low confidence

# How to find each sheet (the rank tracker tab came misspelled as "Eport", so I match by words, not exact names)
SHEETS = {"rank": ["rank"], "gsc": ["gsc"], "map": ["category", "map"]}

issues = []  # the data issues log, filled as the script goes


def log(sheet, problem, rows, solution):
    issues.append({"Sheet": sheet, "Action / problem": problem,
                   "Rows affected": rows, "Solution": solution})


def norm(text):
    """Excel: =LOWER(TRIM(B2)). TRIM also turns double spaces into one, so I do the same."""
    return text.astype(str).str.strip().str.lower().apply(lambda s: re.sub(r"\s+", " ", s))


def find_sheet(book, words):
    for name, df in book.items():
        low = name.lower()
        if any(w in low for w in words) and "clean" not in low:
            return df.copy()
    sys.exit(f"Could not find a sheet with {words}. Sheets in the file: {list(book)}")


# ------------------------------------------------------------------ 1. load
def load(path):
    book = pd.read_excel(path, sheet_name=None)
    rt = find_sheet(book, SHEETS["rank"])
    gsc = find_sheet(book, SHEETS["gsc"])
    kmap = find_sheet(book, SHEETS["map"])
    rt.columns = ["Week", "Keyword", "URL", "Device", "Rank", "Est Search Volume", "Clicks", "Impressions"]
    gsc.columns = ["Month", "Query", "Clicks", "Impressions", "CTR", "Avg Position"]
    kmap.columns = ["Keyword", "Theme", "Priority"]
    # format="mixed": the export had dates in more than one format and failed on my PC without it
    rt["Week"] = pd.to_datetime(rt["Week"], format="mixed")
    gsc["Month"] = pd.to_datetime(gsc["Month"], format="mixed")
    return rt, gsc, kmap


# ------------------------------------------------------------------ 2. rank tracker (RT_clean)
def clean_rank_tracker(rt):
    # 2a. Normalize keywords (Excel: Keyword norm column + pivot before/after)
    rt["Keyword norm"] = norm(rt["Keyword"])
    before, after = rt["Keyword"].nunique(), rt["Keyword norm"].nunique()
    if before != after:
        log("RT_clean", f"Same keyword written in different ways (spaces, upper/lower case): "
            f"{before} variants for {after} keywords", int((rt["Keyword"] != rt["Keyword norm"]).sum()),
            "Normalized with lower + trim before grouping and joining")

    # 2b. Duplicates: one reading per week + keyword + device (Excel: key + COUNTIF + Remove Duplicates)
    key = ["Week", "Keyword norm", "Device"]
    dups = rt[rt.duplicated(key, keep=False)]
    if len(dups):
        cols = ["Week", "Keyword norm", "Device", "URL", "Rank", "Est Search Volume", "Clicks", "Impressions"]
        identical = dups[cols].drop_duplicates().shape[0] == dups[key].drop_duplicates().shape[0]
        n = int(rt.duplicated(key).sum())
        log("RT_clean", "Same week + keyword + device more than once (should be unique)", n,
            "Removed, kept the first one" + (" (the repeated rows were identical in every field)"
                                              if identical else " (WARNING: repeated rows have different values, review)"))
        rt = rt.drop_duplicates(key, keep="first").copy()

    # 2c. Blank ranks: site not found in the tracked results that week. Kept blank, never 0.
    blanks = int(rt["Rank"].isna().sum())
    if blanks:
        log("RT_clean", "Blank rank (site didn't show up in the tracked results that week)", blanks,
            "Kept blank, not replaced with 0 (would look like position 1). Left out of the average")

    # 2d. Range checks (Excel: sorting each column)
    bad = int(((rt["Rank"] < 1) | (rt["Rank"] > 100)).sum() + (rt["Clicks"] > rt["Impressions"]).sum()
              + (rt[["Clicks", "Impressions"]] < 0).sum().sum())
    log("RT_clean", "Range check: rank 1-100, no negatives, clicks <= impressions", bad,
        "No changes needed" if bad == 0 else "Out of range values found, review")

    # 2e. Month column (Excel: =DATE(YEAR(A2),MONTH(A2),1)), kept as a real date, not a month name
    rt["Month"] = rt["Week"].dt.to_period("M").dt.to_timestamp()
    weeks = rt.groupby("Month")["Week"].nunique()
    short = weeks[weeks < FULL_MONTH_WEEKS]
    detail = ", ".join(f"{m:%b} {n}" for m, n in weeks.items())
    log("RT_clean", f"Weeks per month are uneven ({detail}). Rank tracker clicks/impressions are estimates",
        int(rt["Month"].isin(short.index).sum()),
        "Average rank instead of sums; months with fewer than "
        f"{FULL_MONTH_WEEKS} weeks flagged as low confidence. Clicks and impressions taken from GSC")

    # 2f. Roll up to keyword x month (Excel: pivot -> RT_month)
    rt_month = (rt.groupby(["Keyword norm", "Month"])
                  .agg(**{"Avg Rank": ("Rank", "mean"), "Weeks ranked": ("Rank", "count"),
                          "URL": ("URL", "first")})
                  .reset_index())
    rt_month["Rank confidence"] = rt_month["Month"].map(
        lambda m: "Low (only %d week)" % weeks[m] if weeks[m] < FULL_MONTH_WEEKS else "OK")
    return rt_month


# ------------------------------------------------------------------ 3. GSC (GSC_clean)
def clean_gsc(gsc):
    gsc["Keyword norm"] = norm(gsc["Query"])

    per_query = gsc.groupby("Keyword norm")["Month"].nunique()
    months = gsc["Month"].nunique()
    gaps = int((per_query != months).sum())
    log("GSC_clean", f"Months per query check ({len(per_query)} queries)", gaps,
        "All queries have every month" if gaps == 0 else "Some queries are missing months, review")

    # CTR rule (Excel: CTR manual -> Calc diff with ABS -> Threshold -> Flag -> CTR final)
    gsc["CTR manual"] = gsc["Clicks"] / gsc["Impressions"]
    gsc["Calc diff"] = (gsc["CTR"].fillna(0) - gsc["CTR manual"]).abs()   # blank counts as 0, like in Excel
    gsc["CTR flag"] = gsc["Calc diff"].gt(CTR_THRESHOLD).map({True: "Replaced", False: "Kept"})
    gsc["CTR final"] = gsc["CTR"].where(gsc["CTR flag"] == "Kept", gsc["CTR manual"])
    blank = int(gsc["CTR"].isna().sum())
    replaced = int((gsc["CTR flag"] == "Replaced").sum())
    log("GSC_clean", "CTR blank or not matching clicks / impressions", replaced,
        f"Recalculated CTR; replaced when the difference is above {CTR_THRESHOLD} "
        f"({blank} blank, {replaced - blank} wrong). Smaller gaps are rounding, kept")

    dup = int(gsc.duplicated(["Keyword norm", "Month"]).sum())
    if dup:
        log("GSC_clean", "Same query + month more than once", dup, "Summed")
    gsc = (gsc.groupby(["Keyword norm", "Month"])
              .agg(Clicks=("Clicks", "sum"), Impressions=("Impressions", "sum"),
                   **{"CTR": ("CTR final", "mean"), "Avg Position": ("Avg Position", "mean")})
              .reset_index())
    return gsc


# ------------------------------------------------------------------ 4. keyword map (MAP_clean)
def clean_map(kmap):
    kmap["Keyword norm"] = norm(kmap["Keyword"])
    for col in ["Theme", "Priority"]:
        kmap[col] = kmap[col].astype(str).str.strip()
    dup = int(kmap.duplicated("Keyword norm").sum())
    log("MAP_clean", f"Map check: {len(kmap)} keywords, repeated keywords", dup,
        "No repeats" if dup == 0 else "Kept the first mapping, review the rest")
    return kmap.drop_duplicates("Keyword norm")[["Keyword norm", "Theme", "Priority"]]


# ------------------------------------------------------------------ 5. merge (Merged)
def merge(rt_month, gsc, kmap):
    # Key = keyword + month (Excel: =A2&"|"&TEXT(B2,"yyyy-mm")). Here I join on both columns directly.
    # Stack both key lists and remove duplicates so a keyword in only one source is not lost.
    keys = pd.concat([rt_month[["Keyword norm", "Month"]], gsc[["Keyword norm", "Month"]]]).drop_duplicates()
    merged = (keys.merge(rt_month, on=["Keyword norm", "Month"], how="left")
                  .merge(gsc, on=["Keyword norm", "Month"], how="left")
                  .merge(kmap, on="Keyword norm", how="left"))
    merged["Theme"] = merged["Theme"].fillna("UNMAPPED")
    merged["Priority"] = merged["Priority"].fillna("Unassigned")
    # In RT / In GSC (Excel: COUNTIF of the key in each sheet)
    pair = list(zip(merged["Keyword norm"], merged["Month"]))
    rt_keys = set(zip(rt_month["Keyword norm"], rt_month["Month"]))
    gsc_keys = set(zip(gsc["Keyword norm"], gsc["Month"]))
    merged["In RT"] = [int(p in rt_keys) for p in pair]
    merged["In GSC"] = [int(p in gsc_keys) for p in pair]
    merged["Source"] = [("Both" if r and g else "Rank tracker only" if r else "GSC only")
                        for r, g in zip(merged["In RT"], merged["In GSC"])]

    # 6. Validation. In Excel my first merge left 94 keys instead of 72 because one key pointed to the
    # wrong column. So the script checks the totals and stops instead of writing a wrong report.
    problems = []
    if merged.duplicated(["Keyword norm", "Month"]).any():
        problems.append("keyword-month rows are not unique")
    if merged["Clicks"].sum() != gsc["Clicks"].sum():
        problems.append(f"clicks after merge ({merged['Clicks'].sum()}) != GSC clicks ({gsc['Clicks'].sum()})")
    if merged["In RT"].sum() != len(rt_month):
        problems.append(f"{len(rt_month) - merged['In RT'].sum()} rank tracker rows didn't match")
    if problems:
        sys.exit("Validation failed, report not written: " + "; ".join(problems))

    only_gsc = sorted(merged.loc[merged["Source"] == "GSC only", "Keyword norm"].unique())
    log("Merged", f"Merged {len(merged)} keyword-month rows ({(merged['Source'] == 'Both').sum()} in both, "
        f"{(merged['Source'] == 'GSC only').sum()} GSC only). Total clicks match GSC",
        len(merged), "Validation passed")
    if only_gsc:
        log("Merged", "Queries in GSC that are not tracked in the rank tracker (variants)", len(only_gsc),
            "Kept as 'GSC only', not merged into a tracked keyword by hand: " + "; ".join(only_gsc))
    return merged


# ------------------------------------------------------------------ 7. unmapped + theme report
# Words that appear everywhere in this client's keywords, so they don't say anything about the product
GENERIC_WORDS = {"software", "platform", "system", "systems", "tool", "tools", "best",
                 "supply", "chain", "management"}


def suggest(keyword, kmap):
    """Suggestion only, never applied: theme of the mapped keyword that shares the most specific words."""
    words = set(keyword.split()) - GENERIC_WORDS
    best, score = "Needs review (no clear theme)", 0
    for kw, theme in zip(kmap["Keyword norm"], kmap["Theme"]):
        s = len(words & (set(kw.split()) - GENERIC_WORDS))
        if s > score:
            best, score = theme, s
    return best


def unmapped_list(merged, kmap):
    u = merged[merged["Theme"] == "UNMAPPED"]
    out = (u.groupby("Keyword norm")
             .agg(Source=("Source", "first"), URL=("URL", "first"),
                  Impressions=("Impressions", "sum"), Clicks=("Clicks", "sum"))
             .reset_index().sort_values("Impressions", ascending=False))
    out["Suggested theme (to approve)"] = out["Keyword norm"].apply(lambda k: suggest(k, kmap))
    share = u["Impressions"].sum() / merged["Impressions"].sum()
    log("Unmapped", f"Keywords with no theme: {len(out)} ({share:.0%} of all impressions)", len(out),
        "Reported as their own group (UNMAPPED). Suggested theme listed but NOT applied: mapping is a business decision")
    return out


def theme_report(merged):
    months = sorted(merged["Month"].unique())
    first, last = months[0], months[-1]
    impr = merged.pivot_table(index="Theme", columns="Month", values="Impressions", aggfunc="sum")
    clicks = merged.pivot_table(index="Theme", columns="Month", values="Clicks", aggfunc="sum")
    rank = merged.pivot_table(index="Theme", columns="Month", values="Avg Rank", aggfunc="mean")

    rep = pd.DataFrame(index=impr.index)
    for m in months:
        rep[f"Impressions {pd.Timestamp(m):%b}"] = impr[m]
    rep["Impressions change"] = impr[last] / impr[first] - 1
    rep["Clicks change"] = clicks[last] / clicks[first] - 1
    rep[f"Avg rank {pd.Timestamp(first):%b}"] = rank[first]
    rep[f"Avg rank {pd.Timestamp(last):%b}"] = rank[last]
    rep["Share of impressions (last month)"] = impr[last] / impr[last].sum()
    rep["Trend"] = rep["Impressions change"].apply(
        lambda c: "Gaining" if c > VISIBILITY_THRESHOLD else "Losing" if c < -VISIBILITY_THRESHOLD else "Stable")
    rep["Note"] = f"Visibility = GSC impressions, {pd.Timestamp(first):%b} vs {pd.Timestamp(last):%b}. " \
                  f"Rank for {pd.Timestamp(last):%b} may be low confidence (see Merged)"
    return rep.reset_index().sort_values("Impressions change", ascending=False)


# ------------------------------------------------------------------ 8. write output
def write(path_out, rep, merged, unm):
    merged = merged.copy()
    merged["Month"] = merged["Month"].dt.strftime("%Y-%m")
    merged = merged[["Keyword norm", "Month", "Theme", "Priority", "Source", "Avg Rank", "Weeks ranked",
                     "Rank confidence", "URL", "Clicks", "Impressions", "CTR", "Avg Position"]] \
        .sort_values(["Theme", "Keyword norm", "Month"])
    with pd.ExcelWriter(path_out, engine="openpyxl") as w:
        for name, df in [("Theme report", rep), ("Merged", merged), ("Unmapped keywords", unm),
                         ("Data issues log", pd.DataFrame(issues))]:
            df.round(4).to_excel(w, sheet_name=name, index=False)
            ws = w.sheets[name]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 60)
            for cell in ws[1]:
                cell.font = Font(bold=True)


def main():
    if len(sys.argv) < 2:
        sys.exit("Usage: python seo_weekly_report.py <path to export .xlsx>")
    rt, gsc, kmap = load(Path(sys.argv[1]))
    rt_month = clean_rank_tracker(rt)
    gsc = clean_gsc(gsc)
    kmap = clean_map(kmap)
    merged = merge(rt_month, gsc, kmap)
    unm = unmapped_list(merged, kmap)
    rep = theme_report(merged)

    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)
    out = out_dir / f"seo_theme_report_{date.today():%Y-%m-%d}.xlsx"
    write(out, rep, merged, unm)

    print(rep[["Theme", "Impressions change", "Trend"]].to_string(index=False))
    print(f"\n{len(unm)} unmapped keywords | {len(issues)} log entries | saved to {out}")


if __name__ == "__main__":
    main()
