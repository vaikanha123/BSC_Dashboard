"""Put store footfall into index.html (const SEED_FOOTFALL) for the footfall and conversion views on the
Overview, Region-wise and Store-wise tabs. Run by daily_refresh.py after footfall.py fetch.

  python scripts/update_footfall_dashboard.py --html index.html

SEED_FOOTFALL = {
  through:   last date with any footfall recorded,
  ff:        {date: {store: [new, repeat]}} from the 1st of the previous month (store-days with no
             footfall recorded are left out, so the page skips those days for conversion too),
  prevBills: {date: {store: [bills, newBills, repeatBills]}} for the previous month, from
             data/months/<YYYY-MM>.json (the live month's bills are already in SEED_DAYS),
}
The live month is the newest SEED_DAYS month. Leaves index.html as it was if the JS check fails.
"""
import argparse
import csv
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bsc_common import REGION_MAP, extract_const, replace_const, syntax_check_html_js  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FF_HIST = os.path.join(ROOT, 'data', 'footfall_history_daily.csv')
MONTHS_DIR = os.path.join(ROOT, 'data', 'months')


def num(v):
    return float(v) if v not in ('', None) else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--html', default=os.path.join(ROOT, 'index.html'))
    a = ap.parse_args()

    with open(a.html, encoding='utf-8', newline='') as f:
        content = f.read()
    live = max(extract_const(content, 'SEED_DAYS'))[:7]
    y, m = int(live[:4]), int(live[5:])
    prev = '%d-%02d' % ((y, m - 1) if m > 1 else (y - 1, 12))
    start = prev + '-01'

    ff = {}
    with open(FF_HIST, encoding='utf-8') as f:
        for r in csv.DictReader(f):
            if r['date'] < start or r['store'] not in REGION_MAP:
                continue
            new, rep = num(r['new']), num(r['repeat'])
            if new is None and rep is None or (new or 0) + (rep or 0) <= 0:
                continue
            ff.setdefault(r['date'], {})[r['store']] = [new or 0, rep or 0]
    if not ff:
        sys.exit('STOP: no footfall from %s on in %s' % (start, FF_HIST))

    prev_bills = {}
    path = os.path.join(MONTHS_DIR, prev + '.json')
    if os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            days = json.load(f)['days']
        for dk, d in days.items():
            b, nb, rb = d.get('bills') or {}, d.get('newBills') or {}, d.get('repBills') or {}
            row = {s: [b.get(s, 0), nb.get(s, 0), rb.get(s, 0)] for s in b if s in REGION_MAP}
            if row:
                prev_bills[dk] = row
    else:
        print('NOTE: no archive %s -- last-month conversion will show as "—"' % path)

    seed = {'through': max(ff), 'ff': dict(sorted(ff.items())), 'prevBills': dict(sorted(prev_bills.items()))}
    blob = json.dumps(seed, ensure_ascii=False, separators=(',', ':'))
    if 'const SEED_FOOTFALL = ' in content:
        new = replace_const(content, 'SEED_FOOTFALL', blob)
    else:
        anchor = 'const SEED_NPS = '
        i = content.index(anchor)
        new = (content[:i] + '// Store footfall (new, repeat) per day from the 1st of last month, and last month\'s bills per store per day\n'
               '// -- written by scripts/update_footfall_dashboard.py, feeds the footfall / conversion views.\n'
               'const SEED_FOOTFALL = ' + blob + ';\n' + content[i:])
    with open(a.html, 'w', encoding='utf-8', newline='') as f:
        f.write(new)
    try:
        syntax_check_html_js(a.html)
    except Exception:
        with open(a.html, 'w', encoding='utf-8', newline='') as f:
            f.write(content)
        raise
    print('OK. SEED_FOOTFALL: footfall %s -> %s (%d days), %s bills for %d days'
          % (min(ff), seed['through'], len(ff), prev, len(prev_bills)))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
