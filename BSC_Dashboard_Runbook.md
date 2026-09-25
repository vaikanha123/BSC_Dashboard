# BSC Dashboard Update Runbook

Paste this entire document as your first message in a new chat, along with whatever new files you're sending (sales CSV, refund/CN Excel, etc.), and Claude should be able to pick up the update process without any other context.

---

## 1. What this is

Two live dashboards for Bombay Shirt Company retail performance, hosted on GitHub Pages:

- **Main dashboard**: `https://vaikanha123.github.io/BSC_Dashboard/` (file: `index.html`)
  Daily/MTD sales, targets, refunds & CN, category mix, region-wise, stylist-wise, and a Training Cohort tab.
- **Stylist tracker**: `https://vaikanha123.github.io/BSC_Dashboard/stylist-weekly-tracker.html`
  Detailed AOV/ASP/UPT tracking for two training batches, with a Batch 1 / Batch 2 selector at the top.

Both are single self-contained HTML files (data baked in as JS constants — no server, no database). Updating them means: parse new source files → recompute the relevant JS constants → push to GitHub.

## 2. GitHub access

```
Repo: https://github.com/vaikanha123/BSC_Dashboard.git
Token: NOT stored in this document or in git history on purpose (GitHub's own push
protection blocks commits containing a live token, correctly). Vaibhav will supply the
current token directly in the new chat when needed -- ask for it if it hasn't been given
yet, rather than reusing an old one from a previous conversation (it may have been rotated).
```

**Important:** another Claude session may be working on this same repo concurrently. Always `git pull` immediately before editing, and again immediately before pushing. If a push is rejected, fetch, merge (not overwrite), and retry — never force-push.

```bash
mkdir -p /home/claude/deploy && cd /home/claude/deploy
export GH_TOKEN="<paste the current token here, supplied fresh each session -- never hardcode it in a file that gets committed>"
if [ ! -d repo ]; then
  git clone "https://x-access-token:${GH_TOKEN}@github.com/vaikanha123/BSC_Dashboard.git" repo
fi
cd repo
git config user.email "dashboard-bot@bsc.local"
git config user.name "BSC Dashboard Bot"
git pull "https://x-access-token:${GH_TOKEN}@github.com/vaikanha123/BSC_Dashboard.git" main
# ... make edits ...
git add index.html stylist-weekly-tracker.html
git commit -m "Describe what changed"
git push "https://x-access-token:${GH_TOKEN}@github.com/vaikanha123/BSC_Dashboard.git" main
```

Rotate this token periodically for security; if this document is ever shared outside this immediate context, redact the token first.

## 3. Source files you'll receive, and their formats

**Sales CSV** — filename pattern `Sales_Data_Report_with_Stylist_Name_-_YYYY-MM-DD_-_YYYY-MM-DD.csv` (or without "with_Stylist_Name" for older exports). Key columns: `Day`, `Order name`, `Order or return`, `Line type`, `New or returning customer`, `POS location name`, `Product type`, `Product variant SKU`, `Assisting staff member name`, `Quantity ordered`, `Taxes`, `Gross sales`. Revenue convention used throughout = `Gross sales + Taxes` (this is the established basis for both dashboards — do not switch to "Total sales" or plain "Gross sales" without explicit instruction, as this has been a recurring source of bugs before).

**Refund/CN Excel** — two sheets, one refund one CN (names/order vary: `refund`/`CN`, `refund`/`cn`, etc. — detect by header content, not sheet name). Columns (0-indexed): `Channel`(0), `Customer Name`(1), `Order ID`(2), `Order Date`(3), `POS Location / Store`(4), `Line Item`(5)... amount is column index 7, reason is index 9 for refunds / index 8 for CN, processed-date is index 10 for refunds / index 9 for CN, category is index 6. "Channel" = "Online" identifies online vs offline.

**Targets Excel** — `Daywise_Targets_<Mon>-26.xlsx`, single sheet, columns: `POS location name`, `Date`, `New Revenue`, `Repeat Revenue`, `New orders`, `Repeat orders`. Target per store per day = New Revenue + Repeat Revenue.

**NPS Excel** — sheet named like `Final Pivto` or similar; the SECOND table within the sheet (starting a few rows down) has store name, NPS as a decimal fraction (multiply by 100), respondent counts. Look for a "Grand Total" row for the overall figure.

## 4. Region map (used throughout both dashboards)

```python
REGION_MAP = {
    'Ambience Vasant Kunj': 'North', 'DLF Midtown - Moti Nagar': 'North', 'Gurugram': 'North',
    'Khan Market': 'North', 'Mall of India, Noida': 'North', 'Select City': 'North', 'South Ex.': 'North',
    'Vegas Dwarka': 'North', 'Jaipur Store': 'North', 'Phoenix Palassio': 'North', 'Shakespearesarani': 'North',
    'Express Avenue': 'South', 'Indiranagar': 'South', 'Inorbit mall Hyderabad': 'South', 'Jayanagar': 'South',
    'Jubilee Hills': 'South', 'KNK Chennai': 'South', 'Kochi Store': 'South', 'LakeShore Mall': 'South',
    'Lavelle Road': 'South', 'Phoenix Marketcity, Whitefield': 'South', 'R.K. Salai': 'South',
    'Sarath City-Hyderabad': 'South',
    'Andheri': 'West', 'Inorbit Mall Malad West': 'West', 'Juhu Store': 'West',
    'Kalaghoda, Fort': 'West', 'Kemps Corner': 'West', 'Koregaon Park': 'West', 'Oberoi Mall Store': 'West',
    'Oberoi Sky City': 'West', 'Pali Hill, Bandra': 'West', 'Phoenix Marketcity Kurla': 'West',
    'Sindhu Bhavan Marg': 'West', 'Viviana Mall': 'West', 'PMC Viman Nagar Pune': 'West',
}
# Any store NOT in this dict (e.g. "Online", "Corporate Orders", "Virtual Stylist") is treated as
# non-offline / excluded from region rollups. Never add a broad "Online": "Online" style entry to
# this dict for ad-hoc analysis scripts — that was a real bug earlier (a truthy-check on
# REGION_MAP.get(store) wrongly included Online revenue in "offline" totals).
```

Category logic (MTM vs RTW shirt split):
```python
def categorize(product_type, sku):
    pt = str(product_type).strip() if pd.notna(product_type) else ''
    if pt.lower() == 'shirt':
        sku = str(sku) if pd.notna(sku) else ''
        return 'RTW Shirt' if '-' in sku else 'MTM Shirt'
    return pt if pt else 'Uncategorized'
```

## 5. Main dashboard data pipeline (`SEED_DAYS`, `DAILY_TARGETS`, `SEED_REFUNDS`, `SEED_CN`)

`SEED_DAYS` is rebuilt **fresh each month** (do not accumulate prior months' days in it — that's what makes it a month-transition, not just a daily update). Each date key holds: `stores`, `bills`, `categories`, `categoryUnits`, `newRev`, `repRev`, `newBills`, `repBills`, `stylistRev`, `stylistBills`, `storeUnits`, `stylistStoreRev`, `stylistUnits`, `stylistStoreBills`, `stylistStoreUnits` (the last three are per-stylist-per-store breakdowns, needed for correct store/region filtering and for the store-scoping overrides below).

**Month transition checklist** (do this whenever a new month's first file arrives):
1. Replace `SEED_DAYS` entirely with the new month's data only.
2. Replace `DAILY_TARGETS` with the new month's target file.
3. Replace `SEED_REFUNDS`/`SEED_CN` with the new month (current-month key like `Sep-2026`).
4. Rebuild the **`prevMonth*`** baseline constants (`prevMonthStoreAOV`, `prevMonthStoreUPT`, `prevMonthCategory`, `prevMonthCategoryUnits`, `prevMonthDailyRegion`, `prevMonthDailyCategoryUnits`) from the *just-closed* month's complete data — these drive all "vs last month" comparisons (Store-wise, Category, Region-wise daily trend). These are generically named (not `july26...` or `aug26...`) specifically so this step is a pure data swap, no code changes needed.
5. The Training Cohort tab's July baseline (`JULY_COHORT_BASELINE`, `JULY_STYLIST_DETAIL`) is **fixed and does not change** on month transitions — it's the training program's permanent reference point, not a rolling comparison.
6. The "vs last year" baseline (`aug25StoreRevenue` etc.) currently only has August 2025 data — there is no September 2025 (or later) data available. Leave it untouched and flag this gap; do not silently compare against a mismatched prior-year month.

## 6. Training Cohort tab (main dashboard) + Stylist Tracker — established conventions

- Revenue basis is **Gross Sales + Taxes** everywhere, consistently, on both dashboards. (There was a whole incident where the tracker briefly used pure Gross Sales while the main dashboard used Gross+Taxes — always keep these matched.)
- Comparison numbers between the two dashboards are **pooled** (sum revenue ÷ sum bills across the whole cohort), not a simple average of each stylist's individual AOV. A separate "% of stylists individually improved" stat also exists on both — that one *is* a simple per-stylist headcount. Don't conflate the two.
- **Rest of Network** = everyone NOT in the relevant cohort (Batch 1 excludes Batch 1 members; Batch 2 excludes Batch 2 members — these are computed independently, not nested).

### Store-specific overrides (`STORE_SPECIFIC_STYLISTS` on main dashboard, same logic replicated per-stylist in the tracker's RAW data)

**Rahul Yadav** is the one standing override, because he's genuinely worked from different stores in different months and his blended (all-stores) total was misleading:
- July and August: scoped to **Andheri** only (his base at the time)
- September onward: scoped to **Kemps Corner** only (he moved)
- His July baseline in `JULY_STYLIST_DETAIL`/tracker's `jul` field should be the **Andheri-only** figures: AOV ₹11,300, ASP ₹4,022, UPT 2.81 (42 bills, ₹4,74,612 revenue). Recompute Andheri-only for any month where he's still store-scoped; recompute Kemps-Corner-only for September onward.
- If any other stylist is later found to have a similar multi-store misattribution problem, add them to this same override dict — don't build a one-off special case each time.

### Batch 1 roster — 42 stylists, July baseline (fixed)

```json
[
  {
    "name": "ABHISHEK KUMAR",
    "store": "Gurugram",
    "region": "North",
    "priority": "P2"
  },
  {
    "name": "Abhishek Shukla",
    "store": "Oberoi Sky City",
    "region": "West",
    "priority": "P2"
  },
  {
    "name": "adnan zeri",
    "store": "Pali Hill, Bandra",
    "region": "West",
    "priority": "Aug AOV ok"
  },
  {
    "name": "Anas Asif Palwala",
    "store": "Pali Hill, Bandra",
    "region": "West",
    "priority": "Aug AOV ok"
  },
  {
    "name": "Anil Kumar",
    "store": "Phoenix Marketcity, Whitefield",
    "region": "South",
    "priority": "P2"
  },
  {
    "name": "Anjali Rana",
    "store": "Phoenix Marketcity, Whitefield",
    "region": "South",
    "priority": "P2"
  },
  {
    "name": "arif shaikhh",
    "store": "Kalaghoda, Fort",
    "region": "West",
    "priority": "P1"
  },
  {
    "name": "Arslaan Yousufi",
    "store": "Jubilee Hills",
    "region": "South",
    "priority": "P0"
  },
  {
    "name": "Azharuddin Khan",
    "store": "Oberoi Mall Store",
    "region": "West",
    "priority": "P1"
  },
  {
    "name": "Deborshi Dey",
    "store": "Shakespearesarani",
    "region": "North",
    "priority": "P1"
  },
  {
    "name": "Deepak Singh",
    "store": "Mall of India, Noida",
    "region": "North",
    "priority": "P2"
  },
  {
    "name": "Hiten zakhariya",
    "store": "Sindhu Bhavan Marg",
    "region": "West",
    "priority": "P0"
  },
  {
    "name": "inderjit singh",
    "store": "Oberoi Sky City",
    "region": "West",
    "priority": "P2"
  },
  {
    "name": "Irfan Sayed",
    "store": "Pali Hill, Bandra",
    "region": "West",
    "priority": "P2"
  },
  {
    "name": "kaviya Palanisamy",
    "store": "Express Avenue",
    "region": "South",
    "priority": "Aug AOV ok"
  },
  {
    "name": "Krish Ayar",
    "store": "Oberoi Sky City",
    "region": "West",
    "priority": "P2"
  },
  {
    "name": "Manish Sharma",
    "store": "Jaipur Store",
    "region": "North",
    "priority": "P2"
  },
  {
    "name": "manohar pilla",
    "store": "Kemps Corner",
    "region": "West",
    "priority": "Aug AOV ok"
  },
  {
    "name": "Mohammed Baba",
    "store": "Sarath City-Hyderabad",
    "region": "South",
    "priority": "P0"
  },
  {
    "name": "Mohd Sadath",
    "store": "LakeShore Mall",
    "region": "South",
    "priority": "Aug AOV ok"
  },
  {
    "name": "Mohiz Shaikh",
    "store": "Kemps Corner",
    "region": "West",
    "priority": "P1"
  },
  {
    "name": "Naga Sathish",
    "store": "Jubilee Hills",
    "region": "South",
    "priority": "P0"
  },
  {
    "name": "Neha Gaikar",
    "store": "Kalaghoda, Fort",
    "region": "West",
    "priority": "P1"
  },
  {
    "name": "Nibras Pathan",
    "store": "Kalaghoda, Fort",
    "region": "West",
    "priority": "P1"
  },
  {
    "name": "Niranjan N",
    "store": "Phoenix Marketcity, Whitefield",
    "region": "South",
    "priority": "P2"
  },
  {
    "name": "nitish sharma",
    "store": "DLF Midtown - Moti Nagar",
    "region": "North",
    "priority": "P1"
  },
  {
    "name": "omkar tambe",
    "store": "PMC Viman Nagar Pune",
    "region": "West",
    "priority": "P0"
  },
  {
    "name": "panku iglesias",
    "store": "Vegas Dwarka",
    "region": "North",
    "priority": "P0"
  },
  {
    "name": "Pradeep Pradeep",
    "store": "Ambience Vasant Kunj",
    "region": "North",
    "priority": "P0"
  },
  {
    "name": "Rahul Yadav",
    "store": "Andheri",
    "region": "West",
    "priority": "Aug AOV ok"
  },
  {
    "name": "rajesh nair",
    "store": "Inorbit Mall Malad West",
    "region": "West",
    "priority": "P0"
  },
  {
    "name": "Rajkumar Rajkumar",
    "store": "Indiranagar",
    "region": "South",
    "priority": "Aug AOV ok"
  },
  {
    "name": "ritika dubey",
    "store": "Juhu Store",
    "region": "West",
    "priority": "P0"
  },
  {
    "name": "S Nagasesha Reddy",
    "store": "Lavelle Road",
    "region": "South",
    "priority": "One off month"
  },
  {
    "name": "samir ansari",
    "store": "Andheri",
    "region": "West",
    "priority": "P0"
  },
  {
    "name": "Sarika Tyagi",
    "store": "South Ex.",
    "region": "North",
    "priority": "P2"
  },
  {
    "name": "Suchita Singh",
    "store": "Shakespearesarani",
    "region": "North",
    "priority": "P1"
  },
  {
    "name": "Sunil M",
    "store": "Lavelle Road",
    "region": "South",
    "priority": "One off month"
  },
  {
    "name": "suraj sarkar",
    "store": "LakeShore Mall",
    "region": "South",
    "priority": "P0"
  },
  {
    "name": "Thomas Vincent",
    "store": "Express Avenue",
    "region": "South",
    "priority": "P1"
  },
  {
    "name": "Yasin Sayyed",
    "store": "Kalaghoda, Fort",
    "region": "West",
    "priority": "Aug AOV ok"
  },
  {
    "name": "yogesh manawat",
    "store": "Jaipur Store",
    "region": "North",
    "priority": "P2"
  }
]
```

### Batch 2 roster — 14 stylists, August baseline (this batch started training in August, so it has no July data — the tables gracefully show "N/A" for July, this is expected, don't try to backfill it)

```json
[
  {
    "name": "Hiten zakhariya",
    "store": "Sindhu Bhavan Marg",
    "region": "West"
  },
  {
    "name": "Nisar Mall",
    "store": "Sindhu Bhavan Marg",
    "region": "West"
  },
  {
    "name": "Hari Krishn Sharma",
    "store": "Sindhu Bhavan Marg",
    "region": "West"
  },
  {
    "name": "Akshay Jadhav",
    "store": "Juhu Store",
    "region": "West"
  },
  {
    "name": "Sharik Shaikh",
    "store": "Juhu Store",
    "region": "West"
  },
  {
    "name": "Sushant More",
    "store": "Juhu Store",
    "region": "West"
  },
  {
    "name": "Trisha Roy",
    "store": "Koregaon Park",
    "region": "West"
  },
  {
    "name": "krishna kulkarni",
    "store": "PMC Viman Nagar Pune",
    "region": "West"
  },
  {
    "name": "omkar tambe",
    "store": "PMC Viman Nagar Pune",
    "region": "West"
  },
  {
    "name": "arif shaikhh",
    "store": "Kalaghoda, Fort",
    "region": "West"
  },
  {
    "name": "Neha Gaikar",
    "store": "Kalaghoda, Fort",
    "region": "West"
  },
  {
    "name": "Yasin Sayyed",
    "store": "Kalaghoda, Fort",
    "region": "West"
  },
  {
    "name": "adnan zeri",
    "store": "Pali Hill, Bandra",
    "region": "West"
  },
  {
    "name": "Irfan Sayed",
    "store": "Pali Hill, Bandra",
    "region": "West"
  }
]
```

Batch 2's period selector is locked to "August → September" only (no July option) since there's no July baseline for this batch.

### Tracker's period selector

Three modes exist (Batch 1 only): `jul-aug`, `aug-sep`, `jul-sep`. Each stylist's RAW entry carries `jul`, `mtd` (= frozen August final, kept for backward compatibility), `weeks` (August week-by-week, Mon–Sun boundaries with partial weeks at month start/end), `aug` (same as `mtd`), `sep` (live September MTD — this is the one that keeps changing as new September files arrive).

## 6b. Forecast page (`forecast-1ad8b79c.html`, `scripts/forecast.py`, `data/`)

Offline-store revenue forecast: current-month landing vs target, next 7 days, next month, per region/store, "why this forecast" explanations, plus an accuracy section. Added 2026-09-25 as a tab in index.html, then **moved the same day to its own unlinked page `forecast-1ad8b79c.html`** at the user's request — the main dashboard link is widely shared and they don't want those viewers to see the forecast. Never link to it from index.html or other pages. It is *unlisted, not private*: the repo and GitHub Pages site are public, so anyone browsing the repo can find it. Its month/daily/store targets are copied in by `forecast.py` from index.html's `TARGETS`/`DAILY_TARGETS` (`--dashboard index.html`), using the same exact-daily-else-weekday/weekend rule as `storeTargetForDay()`, and only when those targets are for the current month.

- **Files:** `data/sales_history_daily.csv` (date × POS location revenue/bills/units since 2024-01-01, all locations incl. Online), `data/forecast_log.csv` (every network-level forecast ever made, 1–62 days ahead, `source` = `backtest` or `live`), `data/forecast_backtest.json` (walk-forward results + the champion blend policy), `data/forecast_config.json` (closed stores with their region, temporarily closed stores, festival calendar — **extend the festival dates every year and have HO check them**; 2027 dates are provisional).
- **Daily:** `daily_refresh.py` runs `forecast.py run --sales <MTD csv> --html forecast-1ad8b79c.html --dashboard index.html` as its last step (~30 s). It merges the MTD file into the history (idempotent — replaces those dates), retrains on everything, forecasts, appends to the log and rewrites `const FORECAST` in the forecast page. Non-fatal: on failure `data/` and the forecast page are restored and it keeps the previous day's numbers. **Commit `data/` and `forecast-1ad8b79c.html` together with the HTML files.**
- **Month transition (manual):** the daily task skips the 1st/2nd, so the last day of the closed month never reaches the history through it. Run `forecast.py run --sales "<complete previous month csv>" --sales "<new MTD csv>" --html forecast-1ad8b79c.html --dashboard index.html` after the `--new-month` dashboard update (and after new month targets are loaded, so the page picks them up). If you forget, the next run stops with a "gap" message rather than training on a hole.
- **Method:** three per-store forecasters summed to region/network — A (last 4 same weekdays), B (A × network's same-weeks-last-year movement), G (HistGradientBoosting on target/level ratio). Blended with weights re-learned each run from the last 120 days of scored forecasts (policy `adaptive`); an optional bias-correction policy lost the backtest and is off. 80% ranges = logged error quantiles for the same horizon window.
- **Month landing (added 2026-09-26, revised same day):** landing = MTD actual + a rest-of-month forecast chosen **per phase of the month** by `score_landing()`/`choose_landing()`: days 1-7 `latest`, days 8-21 `avg_start_latest_noB` (average of the forecast made just before the month began and the newest forecast's methods A+G), days 22+ `latest`. Why: the newest forecast over-reacts to a strong/weak start (days 8-14 miss 4.63% -> 3.87% with the blend), but late in the month a start-of-month forecast that has already proven wrong only drags the landing — the first version blended on every day and put Sep-2026 at Rs10.25Cr on day 25 when the 31-Aug forecast had been 14% high all month (per-phase fix: Rs10.15Cr, main dashboard's flat run-rate said Rs10.05Cr). Each run re-scores all candidates per phase on the last 365 days of the log and switches a phase only if another method is >= 5% (relative) better over >= 20 origins. Regions/stores are scaled by the same factor so they add up. The daily chart/next-7-days still use the plain daily forecast. Note the overall differences between these landing methods are small (a few hundredths of a point on average) — the per-phase rule matters most on the individual days where the start-of-month forecast is badly off.
- **Changing the model:** edit `forecast.py`, then `forecast.py backtest --months 18` (~8 min) — it re-scores all policies, reseeds the backtest rows in the log and picks the champion. Only ship a change if its backtest numbers beat the previous `forecast_backtest.json`.
- **Stores:** closures are detected from a ≥14-day zero-sales gap (e.g. R.K. Salai's renovation) and the store rejoins automatically when it sells again, seeded from its pre-closure level. Stores with <28 days of history use a run-rate fallback. Stores that haven't opened aren't forecast. A store permanently closing needs adding to `closed_stores` only if it's also removed from `REGION_MAP`.
- **"Why this forecast" section:** `compute_insights()` in `forecast.py` derives every explanation on the page from the history on each run — the month-to-month bridge (month length, weekend-count calendar effect, learned season/festival/trend), same-store next÷this-month ratios for past years (weekend-adjusted) vs what the forecast assumes, per-region festival-window lift, 13-week same-store YoY growth, weekend multipliers, the three methods' separate totals, and store movers vs the same 4 weeks last year. The JS only turns these numbers into sentences — don't hard-code claims there; add a computed field instead. The collapsible "The science behind it" text is static except the backtest numbers — update it if the method changes.
- **Known limits (2026-09):** only two prior festive seasons, so Oct–Dec month-ahead forecasts are the weakest (backtest Dec 2025 missed by 13%); store-level single days miss by ~50% (use store rows for direction only). West's Oct/Sep same-store uplift is ~1.5× in both past years (Navratri/Dussehra) — a big West October forecast is expected, not a bug.

## 7. Verification discipline — do this before every push

1. **Syntax-check** both files after any edit:
   ```bash
   python3 -c "
   import re
   with open('index.html', encoding='utf-8') as f: content = f.read()
   js = re.findall(r'<script>(.*?)</script>', content, re.DOTALL)[-1]
   with open('/tmp/check.js','w', encoding='utf-8') as f: f.write(js)
   "
   node --check /tmp/check.js && echo OK
   ```
2. **Cross-verify pooled AOV matches** between the main dashboard's Training Cohort tab and the tracker, using the real embedded data (not just eyeballing) — this has caught real bugs multiple times (missed store overrides, stale REST_POOLED figures, partial-day data).
3. Watch for **partial/incomplete "latest day"** data — several times a sales file's most recent day showed a suspiciously low total (an early-morning export before the day's sales came in). Flag this to the user rather than silently treating it as a real low day; a later file with the same date range usually corrects it.

## 8. Style/format conventions established in chat

- ₹ formatted as `₹XX,XXX` (Indian digit grouping) or `₹XX.XXL` / `₹X.XXCr` for large numbers in prose.
- When asked for a shareable image/PDF/report, check `/mnt/skills/public/` for the relevant skill (pdf, docx, pptx) before building.
- Positive framing for anything shared with stakeholders (e.g. "Growth Opportunity List" rather than "Stylists Not Improving").
- The user (Vaibhav) prefers being told directly when a number looks off or a request has a data-quality caveat, rather than being given a clean-looking number that turns out to be wrong later — several past corrections in this project came from him pushing back on a number that "looked too good," and that pushback was right every time it happened. Don't round that instinct off the answer.
