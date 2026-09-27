"""footfall.py -- import store footfall (New / Repeat walk-ins) into data/footfall_history_daily.csv.

Source: HO's "FF Master File" workbook. One sheet per month ("Oct 25 FF" ... "September 26 FF"):
a header row of store names, then a block of daily rows for New FF, then the same dates again for
Repeat FF (later blocks -- TOTAL FF, per-weekday summaries -- are ignored; New + Repeat is
authoritative). Blocks are found by the date resetting, not by their labels, because the labels vary
("New FF"/"NEW FF", "Repeat FF", "Rpeat MTD", "New Total", ...). The "FF Overall Data" sheet is not
used: it stops in June and has text in its number columns.

Usage:  python scripts/footfall.py fetch                            (download the live Google Sheet)
        python scripts/footfall.py import --xlsx "FF Master File.xlsx"   (a file the user uploaded)
Idempotent: dates in the workbook replace the same dates already in the history file.
"""
import argparse
import datetime as dt
import os
import re
import sys

import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bsc_common import REGION_MAP  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FF_HIST = os.path.join(ROOT, 'data', 'footfall_history_daily.csv')
SHEET_ID = '1IGzh7GR2GmsJqPAVu6sCHLshi9CG43sxP57CZKpBE20'  # HO's "FF Master File" Google Sheet

# Footfall-sheet spellings (mostly the Oct-Dec 2025 sheets) -> POS location names used everywhere else.
# The 2026 sheets already use the POS names. Each alias was checked against sales bills per day.
ALIASES = {
    'Juhu': 'Juhu Store', 'PMC Kurla': 'Phoenix Marketcity Kurla', 'Sky City Borivali': 'Oberoi Sky City',
    'Pali Hill, Khar': 'Pali Hill, Bandra', 'Sharath City': 'Sarath City-Hyderabad',
    'Inorbit mall': 'Inorbit mall Hyderabad', 'Select City Mall': 'Select City',
    'Ambience Vasnt Kunj': 'Ambience Vasant Kunj', 'Vegas Mall Dwarka': 'Vegas Dwarka',
    'DLF Midtown': 'DLF Midtown - Moti Nagar', 'Jaipur': 'Jaipur Store', 'Kochi': 'Kochi Store',
    'Lakeshore Mall': 'LakeShore Mall', 'Mall of India': 'Mall of India, Noida', 'Oberoi Mall': 'Oberoi Mall Store',
    'Phoenix Lucknow': 'Phoenix Palassio', 'Phoenix Viman Nagar': 'PMC Viman Nagar Pune',
    'PMC Whitefield': 'Phoenix Marketcity, Whitefield', 'Shakespear Sarani': 'Shakespearesarani',
}
MONTH_SHEET = re.compile(r'^\s*[A-Za-z]+\s+\d{2}\s+FF\s*$')
IGNORE_COLS = {'date', 'day wise total', ''}


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, str) and re.fullmatch(r'\s*\d+(\.\d+)?\s*', v):
        return float(v)
    return None


def parse_sheet(ws, known):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], set()
    header = rows[0]
    cols, unknown = {}, set()
    for j, name in enumerate(header):
        if not isinstance(name, str) or name.strip().lower() in IGNORE_COLS:
            continue
        n = ALIASES.get(name.strip(), name.strip())
        if n in known:
            cols[j] = n
        else:
            unknown.add(name.strip())
    # Blocks = runs of consecutive date rows, separated by any non-date row (labels, totals, blanks).
    blocks, cur = [], []
    for r in rows[1:]:
        dcol = next((k for k in (0, 1) if k < len(r) and isinstance(r[k], dt.datetime)), None)
        if dcol is None:
            if cur:
                blocks.append(cur)
                cur = []
            continue
        cur.append((r[dcol].date(), r))
    if cur:
        blocks.append(cur)
    notes = []
    if len(blocks) < 2:
        return [], unknown, ['only %d date block(s) -- sheet skipped' % len(blocks)]
    # Block 0 = New FF and its dates govern. Block 1 = Repeat FF, lined up by position: August 2026's
    # repeat block carried July's dates (a pasted date column) but August's numbers.
    days = [d for d, _ in blocks[0]]
    rep_days = [d for d, _ in blocks[1]]
    if rep_days != days:
        notes.append('Repeat block dates %s..%s differ from New block %s..%s -- matched by row position'
                     % (rep_days[0], rep_days[-1], days[0], days[-1]))
    if len(rep_days) != len(days):
        return [], unknown, notes + ['New/Repeat blocks have %d vs %d rows -- sheet skipped'
                                     % (len(days), len(rep_days))]
    out = []
    for kind, blk in (('new', blocks[0]), ('repeat', blocks[1])):
        for (d, r), day in zip(blk, days):
            for j, store in cols.items():
                v = _num(r[j]) if j < len(r) else None
                if v is not None:
                    out.append((day.isoformat(), store, kind, v))
    # Cross-check against the sheet's own TOTAL FF block where it has one of the same length.
    tot = next((b for b in blocks[2:] if len(b) == len(days)), None)
    if tot is not None:
        vals = {(k, dd, s): v for dd, s, k, v in out}
        bad = 0
        for (d, r), day in zip(tot, days):
            for j, store in cols.items():
                t = _num(r[j]) if j < len(r) else None
                n, p = vals.get(('new', day.isoformat(), store)), vals.get(('repeat', day.isoformat(), store))
                if t is not None and n is not None and p is not None and abs(n + p - t) > 0.5:
                    bad += 1
        if bad:
            notes.append('%d store-days where New + Repeat != the sheet\'s TOTAL FF' % bad)
    return out, unknown, notes


def load_workbook_ff(path):
    known = set(REGION_MAP) | {'Borivali', 'Seawoods', 'Noida', 'Meharchand Market, New Delhi'}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    recs, unknown = [], {}
    for ws in wb.worksheets:
        if not MONTH_SHEET.match(ws.title):
            continue
        r, u, notes = parse_sheet(ws, known)
        recs += r
        if u:
            unknown[ws.title] = sorted(u)
        for n in notes:
            print('  NOTE %s: %s' % (ws.title, n))
    wb.close()  # read-only mode keeps the file open; Windows can't delete the temp download otherwise
    df = pd.DataFrame(recs, columns=['date', 'store', 'kind', 'ff'])
    # a sheet can hold a stray row for the next month's 1st; keep the last value seen per key
    df = df.drop_duplicates(['date', 'store', 'kind'], keep='last')
    wide = df.pivot_table(index=['date', 'store'], columns='kind', values='ff', aggfunc='sum').reset_index()
    for k in ('new', 'repeat'):
        if k not in wide:
            wide[k] = float('nan')
    wide = wide.dropna(subset=['new', 'repeat'], how='all')
    wide['total'] = wide['new'].fillna(0) + wide['repeat'].fillna(0)
    return wide[['date', 'store', 'new', 'repeat', 'total']].sort_values(['date', 'store']), unknown


def cmd_import(a):
    new, unknown = load_workbook_ff(a.xlsx)
    if new.empty:
        raise SystemExit('STOP: no footfall rows found in %s' % a.xlsx)
    # Future-dated rows are sheet templates pre-filled with 0 -- drop anything after the last day with
    # any non-zero footfall across the network.
    daily = new.groupby('date')['total'].sum()
    last_real = daily[daily > 0].index.max()
    new = new[new['date'] <= last_real]
    # Typo guard: a store-day far above that store's usual level (e.g. Kemps Corner 23-Sep-2026 keyed as
    # 1000 new / 105030 repeat) is blanked, not guessed at, and reported so HO can correct the sheet.
    med = new.groupby('store')['total'].transform(lambda s: s[s > 0].median())
    bad = new['total'] > (10 * med).clip(lower=150)
    for _, r in new[bad].iterrows():
        print('  WARNING typo? %s %s: new=%g repeat=%g (store median %g/day) -- blanked'
              % (r['date'], r['store'], r['new'], r['repeat'], med[_]))
    new.loc[bad, ['new', 'repeat', 'total']] = float('nan')
    if os.path.exists(FF_HIST):
        old = pd.read_csv(FF_HIST)
        old = old[~old['date'].isin(set(new['date']))]
        new = pd.concat([old, new]).sort_values(['date', 'store'])
    new.to_csv(FF_HIST, index=False)
    print('footfall: %d store-days, %s..%s, %d stores -> %s' % (
        len(new), new['date'].min(), new['date'].max(), new['store'].nunique(), os.path.relpath(FF_HIST, ROOT)))
    for sheet, names in unknown.items():
        print('  WARNING %s: columns not matched to a store (ignored): %s' % (sheet, ', '.join(names)))


def cmd_fetch(a):
    """Download the live Google Sheet (link-viewable; the xlsx export works from plain Python) and import it."""
    import tempfile
    import urllib.request
    url = 'https://docs.google.com/spreadsheets/d/%s/export?format=xlsx' % a.sheet_id
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=90) as r:
        body = r.read()
    if not body.startswith(b'PK'):
        raise SystemExit('STOP: footfall sheet download did not return an xlsx (sharing changed?)')
    fd, path = tempfile.mkstemp(suffix='.xlsx')
    try:
        with os.fdopen(fd, 'wb') as f:
            f.write(body)
        a.xlsx = path
        cmd_import(a)
    finally:
        os.remove(path)


def main():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest='cmd', required=True)
    s = sub.add_parser('import')
    s.add_argument('--xlsx', required=True)
    s.set_defaults(fn=cmd_import)
    s = sub.add_parser('fetch')
    s.add_argument('--sheet-id', default=SHEET_ID)
    s.set_defaults(fn=cmd_fetch)
    a = p.parse_args()
    a.fn(a)


if __name__ == '__main__':
    main()
