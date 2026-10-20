"""Put store footfall into the dashboards for the footfall and conversion views. Run by daily_refresh.py
after footfall.py fetch (and after update_sm_cohort.py).

  python scripts/update_footfall_dashboard.py [--html index.html] [--tracker stylist-weekly-tracker.html]

index.html, const SEED_FOOTFALL (Overview, Region-wise, Store-wise and SM Cohort tabs):
  through:   last date with any footfall recorded,
  ff:        {date: {store: [new, repeat]}} from the 1st of the previous month (store-days with no
             footfall recorded are left out, so the page skips those days for conversion too),
  prevBills: {date: {store: [bills, newBills, repeatBills]}} for the previous month, from
             data/months/<YYYY-MM>.json (the live month's bills are already in SEED_DAYS),
  baseline:  {month: '2026-08', stores: {store: AGG}} -- the SM Cohort's August baseline.

stylist-weekly-tracker.html, const SM_FOOTFALL (the SM Cohort view's "Footfall" lens), for the 7 SM stores:
  months:  {YYYY-MM: {store: AGG}} for August, every closed month since, and the live month to date,
  windows: {w, pw, lm: {label, stores: {store: AGG}}} -- last 7 days, the 7 days before, and the same
           dates last month, all ending on the live month's last footfall day.

AGG = [newFF, repeatFF, bills, newBills, repeatBills, storeDays], over the days the store logged footfall.
The live month is the newest SEED_DAYS month. A file is left as it was if its JS check fails.
"""
import argparse
import csv
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bsc_common import REGION_MAP, extract_const, replace_const, syntax_check_html_js  # noqa: E402
from update_sm_cohort import BASELINE_MONTH, SM_COHORT_STORES  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FF_HIST = os.path.join(ROOT, 'data', 'footfall_history_daily.csv')
MONTHS_DIR = os.path.join(ROOT, 'data', 'months')
MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


def num(v):
    return float(v) if v not in ('', None) else None


def month_add(mk, n):
    y, m = int(mk[:4]), int(mk[5:]) + n
    while m < 1:
        y, m = y - 1, m + 12
    while m > 12:
        y, m = y + 1, m - 12
    return '%d-%02d' % (y, m)


def load_ff(start):
    ff = {}
    with open(FF_HIST, encoding='utf-8') as f:
        for r in csv.DictReader(f):
            if r['date'] < start or r['store'] not in REGION_MAP:
                continue
            new, rep = num(r['new']), num(r['repeat'])
            if (new is None and rep is None) or (new or 0) + (rep or 0) <= 0:
                continue
            ff.setdefault(r['date'], {})[r['store']] = [new or 0, rep or 0]
    return ff


def day_bills(days):
    """{date: {store: [bills, newBills, repBills]}} from SEED_DAYS-shaped day blocks."""
    out = {}
    for dk, d in days.items():
        b, nb, rb = d.get('bills') or {}, d.get('newBills') or {}, d.get('repBills') or {}
        row = {s: [b.get(s, 0), nb.get(s, 0), rb.get(s, 0)] for s in b if s in REGION_MAP}
        if row:
            out[dk] = row
    return out


def archive_bills(mk):
    path = os.path.join(MONTHS_DIR, mk + '.json')
    if not os.path.exists(path):
        print('NOTE: no archive %s -- %s conversion will show as "—"' % (path, mk))
        return {}
    with open(path, encoding='utf-8') as f:
        return day_bills(json.load(f)['days'])


def agg(store, dates, ff, bills):
    a = [0.0, 0.0, 0, 0, 0, 0]
    for dk in dates:
        f = (ff.get(dk) or {}).get(store)
        if not f:
            continue
        b = (bills.get(dk) or {}).get(store, [0, 0, 0])
        a[0] += f[0]; a[1] += f[1]; a[2] += b[0]; a[3] += b[1]; a[4] += b[2]; a[5] += 1
    return [round(a[0], 1), round(a[1], 1)] + a[2:]


def drange(a, b):
    d, e, out = dt.date.fromisoformat(a), dt.date.fromisoformat(b), []
    while d <= e:
        out.append(d.isoformat())
        d += dt.timedelta(days=1)
    return out


def label(a, b):
    x, y = dt.date.fromisoformat(a), dt.date.fromisoformat(b)
    if x.month == y.month:
        return '%d–%d %s' % (x.day, y.day, MON[y.month - 1])
    return '%d %s–%d %s' % (x.day, MON[x.month - 1], y.day, MON[y.month - 1])


def write_const(path, name, blob, anchor, comment):
    with open(path, encoding='utf-8', newline='') as f:
        content = f.read()
    if 'const %s = ' % name in content:
        new = replace_const(content, name, blob)
    else:
        i = content.index(anchor)
        nl = '\r\n' if '\r\n' in content else '\n'
        new = content[:i] + comment.replace('\n', nl) + 'const %s = %s;%s' % (name, blob, nl) + content[i:]
    with open(path, 'w', encoding='utf-8', newline='') as f:
        f.write(new)
    try:
        syntax_check_html_js(path)
    except Exception:
        with open(path, 'w', encoding='utf-8', newline='') as f:
            f.write(content)
        raise


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--html', default=os.path.join(ROOT, 'index.html'))
    ap.add_argument('--tracker', default=os.path.join(ROOT, 'stylist-weekly-tracker.html'))
    a = ap.parse_args()

    with open(a.html, encoding='utf-8', newline='') as f:
        seed_days = extract_const(f.read(), 'SEED_DAYS')
    live = max(seed_days)[:7]
    prev = month_add(live, -1)
    ff = load_ff(BASELINE_MONTH + '-01')
    live_ff = [d for d in ff if d[:7] == live]
    if not live_ff:
        sys.exit('STOP: no footfall for %s in %s' % (live, FF_HIST))
    through = max(live_ff)

    bills = day_bills(seed_days)
    mk = BASELINE_MONTH
    while mk < live:
        bills.update(archive_bills(mk))
        mk = month_add(mk, 1)
    prev_bills = {d: v for d, v in bills.items() if d[:7] == prev}

    def month_dates(m):
        first = dt.date.fromisoformat(m + '-01')
        last = dt.date.fromisoformat(month_add(m, 1) + '-01') - dt.timedelta(days=1)
        return drange(first.isoformat(), min(last.isoformat(), through))

    base_dates = month_dates(BASELINE_MONTH)
    seed = {'through': through,
            'ff': {d: v for d, v in sorted(ff.items()) if d >= prev + '-01'},
            'prevBills': dict(sorted(prev_bills.items())),
            'baseline': {'month': BASELINE_MONTH, 'label': MON[int(BASELINE_MONTH[5:]) - 1],
                         'stores': {s: agg(s, base_dates, ff, bills) for s in REGION_MAP}}}
    write_const(a.html, 'SEED_FOOTFALL', json.dumps(seed, ensure_ascii=False, separators=(',', ':')), 'const SEED_NPS = ',
                "// Store footfall per day from the 1st of last month, last month's bills, and the August baseline\n"
                '// -- written by scripts/update_footfall_dashboard.py, feeds the footfall / conversion views.\n')

    # Tracker SM view: months, plus week / last-month windows ending on the live month's last footfall day.
    end = dt.date.fromisoformat(through)
    w0 = (end - dt.timedelta(days=6)).isoformat()
    pw = ((end - dt.timedelta(days=13)).isoformat(), (end - dt.timedelta(days=7)).isoformat())
    lm_first = dt.date.fromisoformat(prev + '-01')
    lm_len = (dt.date.fromisoformat(live + '-01') - dt.timedelta(days=1)).day
    lm = (lm_first.isoformat(), lm_first.replace(day=min(end.day, lm_len)).isoformat())
    months, mk = {}, BASELINE_MONTH
    while mk <= live:
        months[mk] = {s: agg(s, month_dates(mk), ff, bills) for s in SM_COHORT_STORES}
        mk = month_add(mk, 1)
    win = lambda a_, b_: {'label': label(a_, b_), 'stores': {s: agg(s, drange(a_, b_), ff, bills) for s in SM_COHORT_STORES}}
    sm = {'through': through, 'live': live, 'baseline': BASELINE_MONTH, 'months': months,
          'windows': {'w': win(w0, through), 'pw': win(*pw), 'lm': win(*lm),
                      'mtd': {'label': label(live + '-01', through)}}}
    write_const(a.tracker, 'SM_FOOTFALL', json.dumps(sm, ensure_ascii=False, separators=(',', ':')), 'const SM_COHORT_DATA = ',
                '// SM Cohort footfall and bills on footfall-logged days -- written by scripts/update_footfall_dashboard.py.\n')
    print('OK. SEED_FOOTFALL: footfall %s -> %s, %s bills for %d days, %s baseline; SM_FOOTFALL: %s'
          % (min(seed['ff']), through, prev, len(prev_bills), BASELINE_MONTH, ', '.join(months)))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
