# BSC Dashboard update scripts

See `../BSC_Dashboard_Runbook.md` for full context. Quick reference:

```bash
pip install pandas openpyxl --break-system-packages   # if not already installed

# Daily update, main dashboard:
python3 update_main_dashboard.py --html ../index.html \
    --sales /path/to/Sales_Data_Report_...csv \
    --refund /path/to/Refund_...xlsx --current-month Sep-2026 \
    [--targets /path/to/Daywise_Targets_Sep-26.xlsx] [--nps /path/to/NPS_...xlsx]

# New month transition (also rebuilds the "vs last month" baselines):
python3 update_main_dashboard.py --html ../index.html \
    --sales /path/to/first-file-of-new-month.csv \
    --new-month --prev-month-sales /path/to/complete-previous-month.csv

# Stylist tracker (both training batches), daily update:
python3 update_tracker.py --html ../stylist-weekly-tracker.html \
    --sales /path/to/Sales_Data_Report_...csv --period sep
```

Every script run ends with a JS syntax check and refuses to report success if it fails.
Neither script touches git — commit and push separately:

```bash
cd ..
git add index.html stylist-weekly-tracker.html
git commit -m "Update via script: <describe what changed>"
git push "https://x-access-token:${GH_TOKEN}@github.com/vaikanha123/BSC_Dashboard.git" main
```

Always cross-verify the pooled Training Cohort AOV matches between `index.html` and
`stylist-weekly-tracker.html` before pushing — see runbook section 7 for the exact check.
