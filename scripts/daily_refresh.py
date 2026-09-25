"""daily_refresh.py -- one-command daily refresh of index.html + stylist-weekly-tracker.html
from a Shopify "Sales Data Report with Stylist Name" CSV (month-to-date, 1st -> yesterday).

Usage (from the repo root):
    python scripts/daily_refresh.py --sales "<path to csv>" [--refund "<path to xlsx>"]

What it does, in order (stops with a non-zero exit code and a "STOP:" line on any problem,
leaving the files as they were -- it never touches git):
  1. Sanity-checks the CSV: single month, last date == yesterday, latest day not suspiciously low.
  2. Refuses to run on a new-month transition (needs --new-month + the full previous month's file;
     do that by hand per BSC_Dashboard_Runbook.md).
  3. Runs update_main_dashboard.py and update_tracker.py.
  4. Bumps the tracker's hand-written "as of" labels, which update_tracker.py never touches.
  5. Runs verify_cohort_match.js (pooled Training Cohort 1 AOV must match between the two files).
  6. Runs forecast.py run (separate forecast page FORECAST_PAGE + data/ history/log). Non-fatal: on failure data/ is restored
     and a WARNING is printed, but the sales refresh still counts as OK.
On any failure in steps 3-5 the two HTML files are restored from git (git checkout).
"""
import argparse
import datetime as dt
import glob
import os
import re
import shutil
import subprocess
import sys

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# The forecast lives on its own unlinked page (kept off index.html on purpose -- the main dashboard's
# link is widely shared). forecast.py reads the month targets from index.html.
FORECAST_PAGE = 'forecast-1ad8b79c.html'
MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
               'September', 'October', 'November', 'December']


def stop(msg):
    print('STOP: ' + msg)
    sys.exit(2)


def find_node():
    node = shutil.which('node')
    if node:
        return node
    for p in (r'C:\Program Files\nodejs\node.exe', '/c/Program Files/nodejs/node.exe'):
        if os.path.exists(p):
            return p
    stop('node executable not found')


def run(cmd):
    print('$ ' + ' '.join(cmd))
    res = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True, encoding='utf-8', errors='replace')
    tail = '\n'.join(res.stdout.strip().splitlines()[-8:])
    if tail:
        print(tail)
    if res.returncode != 0:
        print(res.stderr[-1500:])
    return res.returncode


def restore(files):
    subprocess.run(['git', 'checkout', '--'] + files, cwd=ROOT)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--sales', required=True)
    ap.add_argument('--refund')
    ap.add_argument('--today', help='override today (YYYY-MM-DD), for testing')
    a = ap.parse_args()

    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    yesterday = today - dt.timedelta(days=1)

    df = pd.read_csv(a.sales)
    if 'Day' not in df.columns:
        stop('CSV has no "Day" column -- is this the right report?')
    days = pd.to_datetime(df['Day'])
    first, last = days.min().date(), days.max().date()
    print('CSV covers %s -> %s (%d rows)' % (first, last, len(df)))

    if first.day != 1 or (first.year, first.month) != (last.year, last.month):
        stop('CSV is not a single-month 1st->N range (%s -> %s)' % (first, last))
    if last != yesterday:
        stop('latest data day is %s but yesterday is %s -- report stale or incomplete' % (last, yesterday))
    if last.day == 1:
        stop('%s is the 1st of a new month -- new-month transition needed (--new-month + previous '
             "month's full file). Do this by hand per the runbook." % last)

    # Latest day must not look like a partial-day export
    order_rows = df[df['Order or return'].str.lower() == 'order']
    daily = order_rows.groupby('Day')['Gross sales'].sum() if 'Gross sales' in df.columns else None
    if daily is not None and len(daily) >= 4:
        ld = daily.iloc[-1]
        med = daily.iloc[:-1].median()
        print('latest-day gross sales %.0f vs median of earlier days %.0f' % (ld, med))
        if ld < 0.25 * med:
            stop('latest day looks partial (%.0f < 25%% of the median %.0f)' % (ld, med))

    period = MON[last.month - 1].lower()
    cur_month = '%s-%d' % (MON[last.month - 1], last.year)
    py = sys.executable
    html_files = ['index.html', 'stylist-weekly-tracker.html']

    cmd = [py, 'scripts/update_main_dashboard.py', '--html', 'index.html', '--sales', a.sales,
           '--current-month', cur_month]
    if a.refund:
        cmd += ['--refund', a.refund]
    if run(cmd) != 0:
        restore(html_files)
        stop('update_main_dashboard.py failed')
    if run([py, 'scripts/update_tracker.py', '--html', 'stylist-weekly-tracker.html', '--sales', a.sales,
            '--period', period]) != 0:
        restore(html_files)
        stop('update_tracker.py failed')

    # Bump the tracker's hand-written "as of" labels
    p = os.path.join(ROOT, 'stylist-weekly-tracker.html')
    with open(p, encoding='utf-8', newline='') as f:
        t = f.read()
    new_asof = 'Data as of %s %d, %d' % (MON[last.month - 1], last.day, last.year)
    # only the live-month labels ("... — September MTD"); frozen ones like "August fully closed" stay
    # The Batch 2 badge in applyPeriodMode() writes the dash as a literal JS escape (backslash-u2014).
    t, n1 = re.subn(r'Data as of [A-Z][a-z]{2} \d{1,2}, \d{4}(?= (?:—|\\u2014) [A-Z][a-z]+ MTD)', new_asof, t)
    t, n2 = re.subn(r'(%s MTD \()\d+( days\))' % MONTH_NAMES[last.month - 1], r'\g<1>%d\g<2>' % last.day, t)
    t, n3 = re.subn(r'([A-Z][a-z]{2} 1(?:–|\\u2013))\d+( (?:MTD|only))', r'\g<1>%d\g<2>' % last.day, t)
    with open(p, 'w', encoding='utf-8', newline='') as f:
        f.write(t)
    print('tracker labels bumped: asof=%d, days=%d, range=%d' % (n1, n2, n3))
    if n1 < 4 or n2 < 3:
        restore(html_files)
        stop('expected >=4 "as of" and >=3 "days" labels in the tracker, found %d/%d' % (n1, n2))

    if run([find_node(), 'scripts/verify_cohort_match.js', period]) != 0:
        restore(html_files)
        stop('pooled Training Cohort AOV does NOT match between index.html and the tracker')

    # Forecast page: merge this MTD file into data/sales_history_daily.csv, retrain, forecast, log.
    # Non-fatal -- if it fails, the sales refresh still ships and the forecast page keeps yesterday's numbers.
    forecast_ok = run([py, 'scripts/forecast.py', 'run', '--sales', a.sales, '--html', FORECAST_PAGE,
                       '--dashboard', 'index.html']) == 0
    if not forecast_ok:
        restore(['data/', FORECAST_PAGE])
        print('WARNING: forecast.py failed -- forecast page NOT updated (still shows the previous run); data/ and the page restored.')

    print('OK: dashboards refreshed through %s%s. Ready to commit + push (include data/ and %s).'
          % (last, '' if forecast_ok else ' (forecast step failed, see WARNING)', FORECAST_PAGE))


if __name__ == '__main__':
    main()
