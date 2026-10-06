#!/usr/bin/env python3
"""
update_sm_cohort.py -- Build the "SM Cohort" tab's data block (`const SM_COHORT_DATA`) in
stylist-weekly-tracker.html: the 7 SM-cohort stores, each with new / repeat / total revenue, bills
and units per stylist, for the August baseline, every closed month since, and the live month, plus
the live month's targets (store + stylist revenue, new / repeat AOV, new / repeat bills).

Usage:
  python scripts/update_sm_cohort.py --html stylist-weekly-tracker.html --sales <csv> [--sales <csv> ...]
      [--daywise-targets "Daywise October target.xlsx"]
      [--store-targets "OND overall target.xlsx" --store-targets-sheet Oct]

State lives in data/sm_cohort.json, so the daily run needs only that day's month-to-date --sales file:
  months  : one block per month ever passed in via --sales (closed months stay as they were last loaded).
  targets : one block per month, written when the two target files are passed. Do that once per month.
The live month is the newest month in `months`; months before 2026-08 (the cohort's baseline) are ignored.

Sources, each used for what it states (same split as the store briefing):
  - Revenue targets, store and stylist, by day: the day-wise file ("Store Total" rows + one row per stylist).
  - New / Repeat AOV targets and New / Repeat customer (bill) targets: the monthly store-targets sheet.
    Its revenue column is NOT used -- it disagrees with the day-wise file for some stores.
"""
import argparse
import datetime
import json
import os

import openpyxl
import pandas as pd

from bsc_common import REGION_MAP, MONTHS, prepare_sales_df, replace_const, syntax_check_html_js

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATE_PATH = os.path.join(ROOT, 'data', 'sm_cohort.json')
CONST_NAME = 'SM_COHORT_DATA'
BASELINE_MONTH = '2026-08'  # the batch started in September on the basis of August performance

# Keep in step with SM_COHORT_STORES in index.html.
SM_COHORT_STORES = ['Juhu Store', 'PMC Viman Nagar Pune', 'Phoenix Marketcity, Whitefield', 'Express Avenue',
                    'Gurugram', 'Vegas Dwarka', 'Mall of India, Noida']
# Store spellings in the monthly store-targets sheet that differ from REGION_MAP.
STORE_TARGET_ALIASES = {'MOI': 'Mall of India, Noida'}
NO_STYLIST = '(no stylist tagged)'


def load_sales(path):
    """A Shopify export, or one re-saved through Excel (DD-MM-YYYY days) -- both end up with ISO days."""
    df = pd.read_csv(path, low_memory=False)
    day = df['Day'].astype(str).str[:10]
    if not day.str.match(r'\d{4}-\d{2}-\d{2}').all():
        day = pd.to_datetime(day, dayfirst=True).dt.strftime('%Y-%m-%d')
    df['Day'] = day
    return prepare_sales_df(df)


def seg_totals(sub):
    return {'rev': round(float(sub['Revenue'].sum()), 2), 'bills': int(sub['Order name'].nunique()),
            'units': round(float(sub['Qty'].sum()), 1)}


def three_way(sub):
    """new / rep / all. 'all' is every order line, so it can exceed new + rep when a line has no segment."""
    return {'new': seg_totals(sub[sub['Segment'] == 'new']),
            'rep': seg_totals(sub[sub['Segment'] == 'returning']),
            'all': seg_totals(sub)}


def month_block(df):
    days = sorted(df['Day_str'].unique())
    stores = {}
    for store in SM_COHORT_STORES:
        sub = df[df['POS location name'] == store]
        entry = three_way(sub)
        entry['daily'] = {d: round(float(v), 2) for d, v in sub.groupby('Day_str')['Revenue'].sum().items()}
        stylists = []
        # One row per stylist, matched case-insensitively (the same person is sometimes typed in a
        # different case); the blank-stylist remainder gets its own row so the rows add up to the store.
        for norm, g in sub.groupby('StylistNorm'):
            row = three_way(g)
            row['name'] = g['Stylist'].mode().iloc[0] if norm else NO_STYLIST
            if row['all']['rev'] or row['all']['bills']:
                stylists.append(row)
        stylists.sort(key=lambda r: -r['all']['rev'])
        entry['stylists'] = stylists
        stores[store] = entry
    return {'firstDay': days[0], 'lastDay': days[-1], 'days': len(days), 'stores': stores}


def load_daywise_targets(path):
    """{month: {store: {'store': {date: target}, 'stylists': {name: {date: target}}}}} for the cohort stores."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    date_cols = [(i, c) for i, c in enumerate(rows[0]) if isinstance(c, (datetime.datetime, datetime.date))]
    if not date_cols:
        raise ValueError(f'No date columns in the header row of {path}')
    month = date_cols[0][1].strftime('%Y-%m')
    out = {}
    for r in rows[1:]:
        store, staff = r[1], r[2]
        if store not in SM_COHORT_STORES or not staff:
            continue
        series = {d.strftime('%Y-%m-%d'): round(float(r[i] or 0), 2) for i, d in date_cols}
        s = out.setdefault(store, {'store': None, 'stylists': {}})
        if str(staff).strip().lower() == 'store total':
            s['store'] = series
        else:
            s['stylists'][str(staff).strip()] = series
    missing = [s for s in SM_COHORT_STORES if not out.get(s) or out[s]['store'] is None]
    if missing:
        raise ValueError(f"No 'Store Total' row in {path} for: {missing}")
    return month, out


def load_store_targets(path, sheet):
    """{store: {newAov, repAov, newBills, repBills}} from the monthly store-targets sheet."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(c).strip().lower() if c is not None else '' for c in rows[0]]
    col = {k: hdr.index(k) for k in ('store name', 'new aov', 'repeat aov', 'new cust', 'rpt cust')}
    out = {}
    for r in rows[1:]:
        name = r[col['store name']]
        if name is None:
            continue
        name = STORE_TARGET_ALIASES.get(str(name).strip(), str(name).strip())
        if name in SM_COHORT_STORES:
            out[name] = {'newAov': float(r[col['new aov']]), 'repAov': float(r[col['repeat aov']]),
                         'newBills': round(float(r[col['new cust']]), 1), 'repBills': round(float(r[col['rpt cust']]), 1)}
    missing = [s for s in SM_COHORT_STORES if s not in out]
    if missing:
        raise ValueError(f'Stores missing from sheet {sheet!r} of {path}: {missing}')
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--html', required=True)
    ap.add_argument('--sales', action='append', default=[])
    ap.add_argument('--daywise-targets')
    ap.add_argument('--store-targets')
    ap.add_argument('--store-targets-sheet')
    a = ap.parse_args()

    state = {'months': {}, 'targets': {}}
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, encoding='utf-8') as f:
            state = json.load(f)

    for path in a.sales:
        df = load_sales(path)
        for month, mdf in df.groupby(df['Day_str'].str[:7]):
            if month < BASELINE_MONTH:
                continue
            state['months'][month] = month_block(mdf)
            print(f"  {month}: {state['months'][month]['firstDay']} -> {state['months'][month]['lastDay']}")

    if a.daywise_targets:
        month, daywise = load_daywise_targets(a.daywise_targets)
        t = state['targets'].setdefault(month, {})
        for store, v in daywise.items():
            t.setdefault(store, {}).update(v)
        print(f'  day-wise targets loaded for {month}')
        if a.store_targets:
            sheet = a.store_targets_sheet or MONTHS[int(month[5:]) - 1]
            for store, v in load_store_targets(a.store_targets, sheet).items():
                t[store].update(v)
            print(f'  AOV / bill targets loaded for {month} (sheet {sheet})')
    elif a.store_targets:
        raise SystemExit('--store-targets needs --daywise-targets (it says which month the targets are for)')

    if BASELINE_MONTH not in state['months']:
        raise SystemExit(f'No {BASELINE_MONTH} baseline in {STATE_PATH} -- pass the full August file via --sales once')

    with open(STATE_PATH, 'w', encoding='utf-8') as f:
        json.dump(state, f, separators=(',', ':'), sort_keys=True)

    months = sorted(state['months'])
    live = months[-1]
    payload = {
        'stores': [{'name': s, 'region': REGION_MAP[s]} for s in SM_COHORT_STORES],
        'baseline': BASELINE_MONTH, 'live': live, 'order': months,
        'months': {m: state['months'][m] for m in months},
        'targets': state['targets'].get(live),  # None until that month's target files are loaded
    }
    with open(a.html, encoding='utf-8', newline='') as f:
        content = f.read()
    content = replace_const(content, CONST_NAME, json.dumps(payload, separators=(',', ':')))
    with open(a.html, 'w', encoding='utf-8', newline='') as f:
        f.write(content)
    syntax_check_html_js(a.html)
    print(f"OK. {CONST_NAME} updated: months {months}, live {live} through {payload['months'][live]['lastDay']}, "
          f"targets: {'yes' if payload['targets'] else 'NOT LOADED for ' + live}.")


if __name__ == '__main__':
    main()
