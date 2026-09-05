#!/usr/bin/env python3
"""
update_tracker.py — Update stylist-weekly-tracker.html with new September (or later) data.

This script updates the LIVE period only (currently 'sep' -- the period that keeps changing
as new files arrive). It does NOT rebuild jul/aug, which are frozen historical baselines by
design (see runbook section 6). If a new month closes and becomes the new frozen baseline
(e.g. September closes and October becomes live), that's a bigger change -- extend this
script's PERIOD_KEY logic rather than hand-editing the tracker file directly.

Usage:
  python3 update_tracker.py --html stylist-weekly-tracker.html --sales sales.csv
      [--period sep] [--batch1-only] [--batch2-only]

  --period : which period key this sales file updates (default: sep). Change this once a
             new month becomes the "live" period being tracked.
"""
import argparse
import json

from bsc_common import (
    load_sales_csv, extract_const, replace_const, syntax_check_html_js,
    STORE_SCOPING_OVERRIDES,
)


def calc_metrics(sub):
    rev = float(sub['Revenue'].sum())
    bills = int(sub['Order name'].nunique())
    units = float(sub['Qty'].sum())
    return {
        'aov': round(rev / bills, 0) if bills else 0,
        'asp': round(rev / units, 0) if units else 0,
        'upt': round(units / bills, 2) if bills else 0,
        'bills': bills, 'revenue': round(rev, 0), 'units': round(units, 0),
    }


def calc_mix(sub):
    new_df = sub[sub['Segment'] == 'new']
    rep_df = sub[sub['Segment'] == 'returning']
    new_bills = int(new_df['Order name'].nunique())
    rep_bills = int(rep_df['Order name'].nunique())
    total = new_bills + rep_bills
    return {
        'new_bills': new_bills, 'rep_bills': rep_bills, 'total_bills': total,
        'new_pct': round(new_bills / total * 100, 1) if total else 0,
        'rep_pct': round(rep_bills / total * 100, 1) if total else 0,
    }


def update_batch(content, raw_const_name, rest_const_name, df, period_key, is_array=True):
    raw = extract_const(content, raw_const_name, is_array=is_array)
    cohort_names = set(r['name'].strip().lower() for r in raw)

    updated = 0
    for r in raw:
        norm = r['name'].strip().lower()
        sub = df[df['StylistNorm'] == norm]
        override = STORE_SCOPING_OVERRIDES.get(norm, {})
        target_store = override.get(period_key)
        if target_store:
            sub = sub[sub['POS location name'] == target_store]
        r[period_key] = calc_metrics(sub)
        if 'mix' not in r:
            r['mix'] = {}
        r['mix'][period_key] = calc_mix(sub)
        updated += 1
    content = replace_const(content, raw_const_name, json.dumps(raw))
    print(f"  {raw_const_name}: updated {updated} stylists for period '{period_key}'")

    # Rest of network = everyone in this sales file NOT in this batch's cohort
    rest = df[~df['StylistNorm'].isin(cohort_names) & (df['Stylist'] != '')]
    rest_metrics = calc_metrics(rest)
    rest_pooled = extract_const(content, rest_const_name, is_array=False)
    rest_pooled[period_key] = rest_metrics
    content = replace_const(content, rest_const_name, json.dumps(rest_pooled))
    print(f"  {rest_const_name}.{period_key}: Rs{rest_metrics['revenue']:,.0f} ({rest_metrics['bills']} bills)")

    return content


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--html', required=True, help='Path to stylist-weekly-tracker.html to update in place')
    ap.add_argument('--sales', required=True, help='Sales CSV covering the live period (e.g. Sep 1-15)')
    ap.add_argument('--period', default='sep', help="Period key being updated (default: 'sep')")
    ap.add_argument('--batch1-only', action='store_true')
    ap.add_argument('--batch2-only', action='store_true')
    args = ap.parse_args()

    with open(args.html, encoding='utf-8') as f:
        content = f.read()

    print(f"Loading sales file: {args.sales}")
    df = load_sales_csv(args.sales)

    if not args.batch2_only:
        print("Updating Batch 1 (BATCH1_RAW / BATCH1_REST_POOLED)...")
        content = update_batch(content, 'BATCH1_RAW', 'BATCH1_REST_POOLED', df, args.period)

    if not args.batch1_only:
        print("Updating Batch 2 (BATCH2_RAW / BATCH2_REST_POOLED)...")
        content = update_batch(content, 'BATCH2_RAW', 'BATCH2_REST_POOLED', df, args.period)

    with open(args.html, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Running JS syntax check...")
    syntax_check_html_js(args.html)
    print(f"OK. {args.html} updated successfully.")
    print("Next: git add / commit / push, and cross-verify pooled AOV against the main "
          "dashboard's Training Cohort tab before pushing (runbook section 7).")


if __name__ == '__main__':
    main()
