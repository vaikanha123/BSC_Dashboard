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
    STORE_SCOPING_OVERRIDES, MALL_STORES, HS_STORES, COHORT3_HS, COHORT3_MALL,
    drop_uniform_orders,
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


def update_batch(content, raw_const_name, rest_const_name, df, period_key, is_array=True,
                 segment=None, rest_stores=None, rest_exclude=None):
    """segment: only count this customer segment's bills ('returning' for Cohort 3's mall track).
    rest_stores: limit "rest of network" to these stores (Cohort 3 compares within a store format).
    rest_exclude: extra stylist names to keep out of "rest" (the other half of Cohort 3)."""
    raw = extract_const(content, raw_const_name, is_array=is_array)
    cohort_names = set(r['name'].strip().lower() for r in raw)
    seg_df = df[df['Segment'] == segment] if segment else df

    updated = 0
    for r in raw:
        norm = r['name'].strip().lower()
        sub = seg_df[seg_df['StylistNorm'] == norm]
        full = df[df['StylistNorm'] == norm]
        override = STORE_SCOPING_OVERRIDES.get(norm, {})
        target_store = override.get(period_key)
        if target_store:
            sub = sub[sub['POS location name'] == target_store]
            full = full[full['POS location name'] == target_store]
        r[period_key] = calc_metrics(sub)
        if 'mix' not in r:
            r['mix'] = {}
        r['mix'][period_key] = calc_mix(full)
        updated += 1
    content = replace_const(content, raw_const_name, json.dumps(raw))
    print(f"  {raw_const_name}: updated {updated} stylists for period '{period_key}'")

    # Rest of network = everyone in this sales file NOT in this batch's cohort
    rest = seg_df[~seg_df['StylistNorm'].isin(cohort_names | set(rest_exclude or [])) & (seg_df['Stylist'] != '')]
    if rest_stores is not None:
        rest = rest[rest['POS location name'].isin(rest_stores)]
    rest_metrics = calc_metrics(rest)
    rest_pooled = extract_const(content, rest_const_name, is_array=False)
    rest_pooled[period_key] = rest_metrics
    content = replace_const(content, rest_const_name, json.dumps(rest_pooled))
    print(f"  {rest_const_name}.{period_key}: Rs{rest_metrics['revenue']:,.0f} ({rest_metrics['bills']} bills)")

    return content


def update_cohort3(content, df, period_key):
    """Cohort 3: high street on overall AOV vs the rest of high-street stores; mall on repeat-customer
    AOV vs the rest of mall stores' repeat bills."""
    content = update_batch(content, 'BATCH3HS_RAW', 'BATCH3HS_REST_POOLED', df, period_key,
                           rest_stores=HS_STORES, rest_exclude=COHORT3_MALL)
    content = update_batch(content, 'BATCH3MALL_RAW', 'BATCH3MALL_REST_POOLED', df, period_key,
                           segment='returning', rest_stores=MALL_STORES, rest_exclude=COHORT3_HS)
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
    n_before = df['Order name'].nunique()
    df = drop_uniform_orders(df)
    print(f"Left out {n_before - df['Order name'].nunique()} staff-uniform order(s).")

    if not args.batch2_only:
        print("Updating Batch 1 (BATCH1_RAW / BATCH1_REST_POOLED)...")
        content = update_batch(content, 'BATCH1_RAW', 'BATCH1_REST_POOLED', df, args.period)

    if not args.batch1_only:
        print("Updating Batch 2 (BATCH2_RAW / BATCH2_REST_POOLED)...")
        content = update_batch(content, 'BATCH2_RAW', 'BATCH2_REST_POOLED', df, args.period)

    if not args.batch1_only and not args.batch2_only and 'const BATCH3HS_RAW' in content:
        print("Updating Cohort 3 (BATCH3HS_* / BATCH3MALL_*)...")
        content = update_cohort3(content, df, args.period)

    with open(args.html, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Running JS syntax check...")
    syntax_check_html_js(args.html)
    print(f"OK. {args.html} updated successfully.")
    print("Next: git add / commit / push, and cross-verify pooled AOV against the main "
          "dashboard's Training Cohort tab before pushing (runbook section 7).")


if __name__ == '__main__':
    main()
