"""crosscheck_store_pulse.py -- compare this repo's store numbers with the "Ontime Store Pulse" page
(https://claude.ai/artifact/L31BcALVyUPftieBnJ1BBE), which is built by someone else from the same
Shopify export and target files. A cheap independent check on the daily briefing numbers.

Usage (from the repo root):
  python scripts/crosscheck_store_pulse.py --artifact <saved html of the page>

The page cannot be downloaded by a script; read it with the Artifact tool, which saves the raw
HTML locally, and pass that path. The page is a snapshot, usually a day or more behind this repo,
so sales are compared month to date through the latest day BOTH sides have.

Compares, per store: revenue and bills (data/sales_history_daily.csv vs the page's daily facts),
and the month's targets (data/briefing_targets.json vs the page's tg / dt rows). Prints every
store outside tolerance and exits 1 if there is one, 0 if all agree.
"""
import argparse
import csv
import datetime
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EPOCH = datetime.date(2024, 1, 1)
REV_TOL = 0.005      # 0.5% on month-to-date revenue
TARGET_TOL = 0.005
BILLS_TOL = 0.04


def load_page(path):
    with open(path, encoding='utf-8') as f:
        for line in f:
            if line.startswith('const D = '):
                return json.loads(line[len('const D = '):].rstrip().rstrip(';'))
    raise SystemExit(f'No "const D = " data line in {path} -- has the page layout changed?')


def off(a, b, tol):
    return abs(a - b) > tol * max(abs(a), abs(b), 1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--artifact', required=True)
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding='utf-8')

    D = load_page(a.artifact)
    page_last = EPOCH + datetime.timedelta(days=D['meta']['maxDay'])
    with open(os.path.join(ROOT, 'data', 'sales_history_daily.csv'), encoding='utf-8') as f:
        hist = list(csv.DictReader(f))
    ours_last = datetime.date.fromisoformat(max(r['date'] for r in hist))
    last = min(page_last, ours_last)
    first = last.replace(day=1)
    month = last.strftime('%Y-%m')
    mi = (last.year - 2024) * 12 + last.month - 1
    with open(os.path.join(ROOT, 'data', 'briefing_targets.json'), encoding='utf-8') as f:
        targets = json.load(f).get(month, {})
    print(f"Store Pulse built {D['meta']['generated']}, data through {page_last}; repo through {ours_last}. "
          f"Comparing {first} to {last}.")

    d0, d1 = (first - EPOCH).days, (last - EPOCH).days
    page = {}
    for day, sid, _seg, sale, _ret, _units, orders in D['f1']:
        if d0 <= day <= d1:
            p = page.setdefault(D['stores'][sid], [0.0, 0])
            p[0] += sale
            p[1] += orders
    ours = {}
    for r in hist:
        if first.isoformat() <= r['date'] <= last.isoformat():
            o = ours.setdefault(r['loc'], [0.0, 0])
            o[0] += float(r['rev'])
            o[1] += int(r['bills'])

    problems = []
    for store in sorted(targets):
        o, p = ours.get(store), page.get(store)
        if not o and not p:
            continue                      # no sales on either side this month (store not trading)
        if not o or not p:
            problems.append(f"{store}: sales {'only here' if o else 'only on the page'}")
            continue
        if off(o[0], p[0], REV_TOL):
            problems.append(f'{store}: revenue {o[0]:,.0f} here vs {p[0]:,.0f} on the page ({(o[0] / p[0] - 1) * 100:+.1f}%)')
        # Bills are counted slightly differently (here per day, so an order billed across two days
        # counts twice; the page counts each order once), hence the wider tolerance.
        if abs(o[1] - p[1]) > max(2, BILLS_TOL * p[1]):
            problems.append(f'{store}: bills {o[1]} here vs {p[1]} on the page')

    sid = {n: i for i, n in enumerate(D['stores'])}
    page_tg = {r[1]: r for r in D['tg'] if r[0] == mi}
    page_dt = {r[1]: sum(r[2]) for r in D.get('dt', []) if r[0] == mi}
    if not page_tg and not page_dt:
        print(f'  (the page has no {month} targets loaded -- targets not compared)')
    for store, t in sorted(targets.items()):
        i = sid.get(store)
        mine = sum(t['store'].values())
        if i in page_dt and off(mine, page_dt[i], TARGET_TOL):
            problems.append(f'{store}: month target {mine:,.0f} here vs {page_dt[i]:,.0f} on the page (day-wise)')
        r = page_tg.get(i)
        if r:
            for label, key, idx in (('New AOV target', 'newAov', 4), ('Repeat AOV target', 'repAov', 5),
                                    ('New customer target', 'newCust', 6), ('Repeat customer target', 'repCust', 7)):
                if t.get(key) and r[idx] and off(t[key], r[idx], 0.01) and abs(t[key] - r[idx]) > 1:
                    problems.append(f'{store}: {label} {t[key]:,.0f} here vs {r[idx]:,.0f} on the page')

    if problems:
        print(f'{len(problems)} DIFFERENCE(S):')
        for p in problems:
            print('  ' + p)
        sys.exit(1)
    print(f'MATCH: {len(targets)} stores agree on revenue, bills and targets.')


if __name__ == '__main__':
    main()
