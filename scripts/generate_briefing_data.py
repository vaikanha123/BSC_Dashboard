"""
generate_briefing_data.py -- Computes per-store daily briefing numbers (required New/Repeat AOV,
repeat customers needed, conversion needed for the rest of the month) and writes them into
store-briefing.html as `const BRIEFING_DATA = {...};`, the same replace-a-JS-constant pattern used
by update_main_dashboard.py / update_tracker.py.

Usage:
  python3 generate_briefing_data.py --html store-briefing.html --sales sales.csv \
      --targets "Daywise Targets Sep-26.xlsx" --footfall footfall.csv [--aov-targets "Targets for Sep.xlsx"]

  --sales       : cumulative current-month sales CSV.
  --targets     : Daywise Targets Excel for the current month (POS location name, Date, New
                   Revenue, Repeat Revenue, New orders, Repeat orders, ...). This is the
                   authoritative source for the full-month revenue/order targets used in the
                   "required AOV" / "orders needed" math -- see BSC_Dashboard_Runbook.md.
  --footfall    : CSV export of the "<Month> FF" tab of the footfall tracker Google Sheet.
  --aov-targets : optional. The separate "Targets for <Mon>.xlsx" file (Store, Target, New Rev
                   Tar, Rep Rev Tar, New AOV Target, Repeat AOV Target, New Bills, Repeat Bill).
                   As of 2026-09, the Rev/Bills columns in this file don't reconcile against the
                   Daywise Targets file or against each other (AOV x Bills != Rev Tar) -- period
                   unclear, unresolved with Vaibhav. Only the two AOV Target columns are pulled
                   from this file, as a labeled reference figure ("HO Target AOV") shown alongside
                   the calculated required AOV -- NOT used in any calculation. Store names in this
                   file are normalized via AOV_TARGET_STORE_MAP in bsc_common.py.

Same calculation logic as scripts/generate_store_briefing.py (that script remains the
plain-text/WhatsApp-friendly version of the same output); see its docstring for the formulas.
"""
import argparse
import calendar
import csv
import datetime
import json

import openpyxl

from bsc_common import REGION_MAP, AOV_TARGET_STORE_MAP, load_sales_csv, replace_const, syntax_check_html_js


def load_targets(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb[wb.sheetnames[0]]
    totals = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        store, date, new_rev, rep_rev, new_orders, rep_orders = row[0], row[1], row[2], row[3], row[4], row[5]
        if store is None or date is None:
            continue
        t = totals.setdefault(store, {'new_rev': 0.0, 'rep_rev': 0.0, 'new_orders': 0, 'rep_orders': 0})
        t['new_rev'] += new_rev or 0
        t['rep_rev'] += rep_rev or 0
        t['new_orders'] += new_orders or 0
        t['rep_orders'] += rep_orders or 0
    return totals


def load_aov_targets(path):
    """Store -> {new_aov, rep_aov, new_bills, rep_bills}, keyed by canonical REGION_MAP name via
    AOV_TARGET_STORE_MAP. This file's own New Bills/Repeat Bill columns, combined with its own AOV
    Target columns, turn out to be the internally-consistent basis for a segment's revenue target
    when the Daywise Targets file's bill-count column doesn't reconcile with the AOV target (see
    segment_block) -- e.g. for Ambience Vasant Kunj, New AOV Target (11,514) x New Bills (250) =
    28,78,500, which lines up with the store's overall Rs29,00,000 target; Daywise's own New orders
    column (270) does not. Confirmed with Vaibhav on 2026-09-09."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb[wb.sheetnames[0]]
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        store = row[0]
        if store is None:
            continue
        canonical = AOV_TARGET_STORE_MAP.get(store)
        if canonical is None:
            print(f"[warn] '{store}' in AOV targets file has no entry in AOV_TARGET_STORE_MAP -- skipped")
            continue
        out[canonical] = {'new_aov': row[4], 'rep_aov': row[5], 'new_bills': row[6], 'rep_bills': row[7]}
    return out


def load_footfall(path, cutoff_date):
    with open(path, encoding='utf-8') as f:
        rows = list(csv.reader(f))

    header = rows[0]
    stores = header[2:-1]

    def parse_date(s):
        s = s.strip()
        for fmt in ('%d-%b-%Y', '%Y-%m-%d'):
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except ValueError:
                continue
        return None

    def num(v):
        v = (v or '').strip().replace(',', '')
        try:
            return float(v)
        except ValueError:
            return 0.0

    blocks = {'New FF': {}, 'Repeat FF': {}}
    current_block = None
    days_seen = {'New FF': set(), 'Repeat FF': set()}
    for row in rows[1:]:
        if len(row) < 2:
            continue
        label = row[1].strip()
        if label in ('New FF', 'Repeat FF'):
            current_block = label
            continue
        if label == 'TOTAL FF':
            current_block = None
            continue
        if current_block is None:
            continue
        d = parse_date(label)
        if d is None or d > cutoff_date:
            continue
        days_seen[current_block].add(d)
        for store, cell in zip(stores, row[2:2 + len(stores)]):
            blocks[current_block].setdefault(store, 0.0)
            blocks[current_block][store] += num(cell)

    return blocks['New FF'], blocks['Repeat FF'], len(days_seen['New FF']), len(days_seen['Repeat FF'])


def segment_block(rev_t, orders_t, rev_a, orders_a, ff_achieved, ff_days, days_remaining, ho_aov, ho_bills=None):
    # The Daywise Targets file's bill-count column (orders_t) and this store's separately-stated
    # AOV target (ho_aov) are supposed to multiply out to the same revenue target. When they
    # don't, Daywise's bill-count column is the unreliable one -- and the "Targets for <Mon>.xlsx"
    # file's own New Bills/Repeat Bill column (ho_bills), paired with its own AOV Target, IS
    # reliable: for Ambience Vasant Kunj, New AOV Target (Rs11,514) x New Bills (250) = Rs28,78,500,
    # matching the store's overall Rs29,00,000 target almost exactly, whereas Daywise's New orders
    # column (270) implies an AOV of just Rs5,195. Confirmed with Vaibhav on 2026-09-09 -- when
    # this mismatch shows up, override both the revenue and bill-count target with this file's own
    # internally-consistent pair rather than trusting Daywise's bill-count column.
    implied_target_aov = (rev_t / orders_t) if orders_t else None
    target_orders_unreliable = bool(
        implied_target_aov and ho_aov and (implied_target_aov / ho_aov < 0.65 or implied_target_aov / ho_aov > 1.5))
    if target_orders_unreliable and ho_aov and ho_bills:
        orders_t = ho_bills
        rev_t = ho_bills * ho_aov

    rev_rem = rev_t - rev_a
    orders_rem = orders_t - orders_a
    cur_aov = (rev_a / orders_a) if orders_a else None

    block = {
        'targetRev': round(rev_t, 2), 'targetOrders': orders_t, 'achievedRev': round(rev_a, 2), 'achievedBills': orders_a,
        'currentAOV': round(cur_aov, 2) if cur_aov is not None else None,
        'hoTargetAOV': ho_aov,
        'remainingRev': round(rev_rem, 2), 'remainingOrders': orders_rem,
        'noTarget': (rev_t == 0 and rev_a == 0),
        'effectiveTargetOrders': orders_t,
    }

    if rev_rem <= 0 and not block['noTarget']:
        block['status'] = 'achieved'
        block['surplus'] = round(-rev_rem, 2)
    elif orders_rem > 0 and rev_rem > 0:
        block['requiredAOV'] = round(rev_rem / orders_rem, 2)
        block['status'] = 'on-track'
    else:
        block['status'] = 'orders-hit-revenue-short'

    block['effectiveOrdersRem'] = orders_rem

    if ff_days > 0 and ff_achieved > 0:
        avg_daily_ff = ff_achieved / ff_days
        proj_ff_rem = avg_daily_ff * days_remaining
        cur_conv = (orders_a / ff_achieved) if ff_achieved else None
        block['ffAchieved'] = round(ff_achieved, 1)
        block['ffDays'] = ff_days
        block['avgDailyFF'] = round(avg_daily_ff, 2)

        if cur_conv is not None and cur_conv > 1.0:
            # More bills than recorded footfall -- footfall for this store/segment isn't being
            # logged reliably (can't actually convert more visitors than walked in). Surface that
            # instead of a >100% "conversion" or a downstream negative "extra footfall needed".
            block['ffUnreliable'] = True
        else:
            block['currentConversion'] = round(cur_conv, 4) if cur_conv is not None else None
            block['projectedFFRemaining'] = round(proj_ff_rem, 1)
            if orders_rem > 0 and proj_ff_rem > 0:
                req_conv = orders_rem / proj_ff_rem
                if req_conv <= 1.0:
                    block['requiredConversion'] = round(req_conv, 4)
                    block['conversionImpossible'] = False
                else:
                    block['conversionImpossible'] = True
                    if cur_conv:
                        block['extraFFNeeded'] = round(orders_rem / cur_conv - proj_ff_rem, 1)
    else:
        block['ffAvailable'] = False

    return block


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--html', required=True)
    ap.add_argument('--sales', required=True)
    ap.add_argument('--targets', required=True)
    ap.add_argument('--footfall', required=True)
    ap.add_argument('--aov-targets', default=None)
    args = ap.parse_args()

    df = load_sales_csv(args.sales)
    df = df[df['POS location name'].isin(REGION_MAP.keys())]

    days = sorted(df['Day_str'].unique())
    last_day = datetime.datetime.strptime(days[-1], '%Y-%m-%d').date()
    days_elapsed = len(days)
    days_in_month = calendar.monthrange(last_day.year, last_day.month)[1]
    days_remaining = days_in_month - days_elapsed

    targets = load_targets(args.targets)
    new_ff, rep_ff, new_ff_days, rep_ff_days = load_footfall(args.footfall, last_day)
    aov_targets = load_aov_targets(args.aov_targets) if args.aov_targets else {}

    stores_out = {}
    for store in sorted(REGION_MAP.keys()):
        if store not in targets:
            continue
        sub = df[df['POS location name'] == store]
        new_sub = sub[sub['Segment'] == 'new']
        rep_sub = sub[sub['Segment'] == 'returning']
        t = targets[store]
        ho = aov_targets.get(store, {})

        new_block = segment_block(
            t['new_rev'], t['new_orders'], float(new_sub['Revenue'].sum()), int(new_sub['Order name'].nunique()),
            new_ff.get(store, 0.0), new_ff_days, days_remaining, ho.get('new_aov'), ho.get('new_bills'))
        rep_block = segment_block(
            t['rep_rev'], t['rep_orders'], float(rep_sub['Revenue'].sum()), int(rep_sub['Order name'].nunique()),
            rep_ff.get(store, 0.0), rep_ff_days, days_remaining, ho.get('rep_aov'), ho.get('rep_bills'))

        stores_out[store] = {
            'region': REGION_MAP[store],
            # Sum the two segments' own (possibly-overridden) targets rather than re-reading
            # Daywise's raw new_rev+rep_rev, so this always agrees with what the segment cards show.
            'monthTarget': round(new_block['targetRev'] + rep_block['targetRev'], 2),
            'achievedTotal': round(float(sub['Revenue'].sum()), 2),
            'new': new_block, 'rep': rep_block,
        }

    payload = {
        'generatedAt': datetime.datetime.now().isoformat(timespec='seconds'),
        'lastSalesDay': last_day.isoformat(),
        'daysElapsed': days_elapsed,
        'daysRemaining': days_remaining,
        'daysInMonth': days_in_month,
        'stores': stores_out,
    }

    with open(args.html, encoding='utf-8') as f:
        content = f.read()
    content = replace_const(content, 'BRIEFING_DATA', json.dumps(payload))
    with open(args.html, 'w', encoding='utf-8') as f:
        f.write(content)

    syntax_check_html_js(args.html)
    print(f"OK. {args.html} updated with {len(stores_out)} store(s), "
          f"through {last_day.isoformat()} ({days_remaining} days remaining).")


if __name__ == '__main__':
    main()
