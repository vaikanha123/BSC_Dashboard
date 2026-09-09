"""
generate_store_briefing.py -- Per-store daily briefing: what New/Repeat AOV, how many repeat
customers, and what conversion each store needs for the REST of the month to still hit its
monthly target, given what it has already achieved.

Usage:
  python3 generate_store_briefing.py --sales sales.csv --targets "Daywise Targets Sep-26.xlsx" \
      --footfall footfall.csv [--store "Kemps Corner"] [--out briefings.txt]

Inputs:
  --sales    : cumulative current-month sales CSV (same file used for update_main_dashboard.py).
               The latest date present in this file is treated as "yesterday" / last complete day.
  --targets  : Daywise Targets Excel for the current month. Must have columns (in order):
               POS location name, Date, New Revenue, Repeat Revenue, New orders, Repeat orders,
               ... (extra columns are ignored).
  --footfall : CSV export of the "<Month> FF" tab of the footfall tracker Google Sheet (File >
               Download > CSV, or via the sheet's export?format=csv&gid=... URL once it's
               link-viewable). Expected layout: a "New FF" block, then a "Repeat FF" block, then
               a "TOTAL FF" block, each with one row per date and one column per store, matching
               REGION_MAP store names exactly.
  --store    : optional, limit output to one store (must match REGION_MAP key exactly).
  --out      : optional, write output to this file instead of stdout.

Calculation logic (see BSC_Dashboard_Runbook.md if this is ever revised):
  Required AOV (remaining days)   = (month revenue target - achieved) / (month orders target - achieved orders)
  Repeat customers needed         = month repeat orders target - repeat bills achieved
  Required conversion (remaining) = orders still needed / footfall projected for remaining days
                                     (footfall projected from this store's own average daily
                                     footfall so far this month -- there is no footfall target file)
  Computed separately for New and Repeat throughout, since that's how the sales/targets/footfall
  data is all split.
"""
import argparse
import calendar
import csv
import datetime

import openpyxl

from bsc_common import REGION_MAP, load_sales_csv


def load_targets(path):
    """Sum New/Repeat Revenue and Orders targets per store across every dated row in the file
    (the whole month's targets), keyed by store name as it appears in the file."""
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


def load_footfall(path, cutoff_date):
    """Parse the New FF / Repeat FF / TOTAL FF stacked blocks. Only dates <= cutoff_date are
    read (later rows in this sheet are unfilled placeholders for days that haven't happened yet,
    and sometimes contain stray non-numeric formula artifacts)."""
    with open(path, encoding='utf-8') as f:
        rows = list(csv.reader(f))

    header = rows[0]
    stores = header[2:-1]  # column 0 blank, column 1 = Date, last column = "Day Wise Total"

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
            current_block = None  # we compute totals ourselves; New+Repeat is authoritative
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

    new_ff_days = len(days_seen['New FF'])
    rep_ff_days = len(days_seen['Repeat FF'])
    return blocks['New FF'], blocks['Repeat FF'], new_ff_days, rep_ff_days


def fmt_inr(n):
    return f"Rs{n:,.0f}"


def fmt_pct(n):
    return f"{n*100:.1f}%"


def build_briefing(store, days_elapsed, days_remaining, achieved, target, ff):
    lines = []
    lines.append(f"=== {store} - Daily Briefing (as of day {days_elapsed} of the month, {days_remaining} days remaining) ===")

    total_target = target['new_rev'] + target['rep_rev']
    total_achieved = achieved['new_rev'] + achieved['rep_rev']
    pct = (total_achieved / total_target * 100) if total_target else 0
    lines.append(f"Month target: {fmt_inr(total_target)}  |  Achieved MTD: {fmt_inr(total_achieved)} ({pct:.0f}%)")
    lines.append("")

    for label, key in (('NEW', 'new'), ('REPEAT', 'rep')):
        rev_t, orders_t = target[f'{key}_rev'], target[f'{key}_orders']
        rev_a, orders_a = achieved[f'{key}_rev'], achieved[f'{key}_bills']
        rev_rem, orders_rem = rev_t - rev_a, orders_t - orders_a
        cur_aov = (rev_a / orders_a) if orders_a else None

        lines.append(f"{label} CUSTOMERS")
        lines.append(f"  AOV (achieved/target): {fmt_inr(cur_aov) if cur_aov else 'n/a'} / "
                      f"{fmt_inr(rev_t/orders_t) if orders_t else 'n/a'}")
        lines.append(f"  Bills (achieved/target): {orders_a:.0f} / {orders_t:.0f}")
        lines.append(f"  Revenue: {fmt_inr(rev_a)} achieved of {fmt_inr(rev_t)} target")

        if rev_t == 0 and rev_a == 0:
            lines.append("  No target set for this segment this month")
        elif orders_rem > 0 and rev_rem > 0:
            req_aov = rev_rem / orders_rem
            lines.append(f"  Needed for remaining {days_remaining} days: {orders_rem:.0f} more orders "
                          f"averaging {fmt_inr(req_aov)} AOV to close {fmt_inr(rev_rem)}")
        elif rev_rem <= 0:
            lines.append(f"  Revenue target already achieved for this segment - keep the pace, "
                          f"{fmt_inr(-rev_rem)} ahead")
        else:
            # revenue still short but order-count target already hit -- needs bonus orders at a
            # bump in AOV, since we can't divide by a non-positive remaining order count.
            lines.append(f"  Order-count target already hit, but still {fmt_inr(rev_rem)} short on revenue "
                          f"- needs extra orders beyond target, or a higher AOV on the ones already counted")

        # Footfall/conversion is supporting context, not the headline -- kept to one terse line.
        ff_key = 'new' if key == 'new' else 'rep'
        ff_days = ff[f'{ff_key}_ff_days']
        ff_achieved = ff[f'{ff_key}_ff']
        if ff_days > 0 and ff_achieved > 0:
            avg_daily_ff = ff_achieved / ff_days
            proj_ff_rem = avg_daily_ff * days_remaining
            cur_conv = orders_a / ff_achieved if ff_achieved else None
            if cur_conv is not None and cur_conv > 1.0:
                lines.append(f"  (Footfall MTD {ff_achieved:.0f} - more orders than recorded footfall, so this "
                              f"store's footfall isn't being logged reliably this month; conversion skipped)")
                lines.append("")
                continue
            note = f"  (Footfall MTD {ff_achieved:.0f} ({avg_daily_ff:.1f}/day) - conversion so far " \
                   f"{fmt_pct(cur_conv) if cur_conv is not None else 'n/a'}"
            if orders_rem > 0 and proj_ff_rem > 0:
                req_conv = orders_rem / proj_ff_rem
                if req_conv <= 1.0:
                    note += f" - conversion needed: {fmt_pct(req_conv)}"
                else:
                    extra_ff_needed = (orders_rem / cur_conv - proj_ff_rem) if cur_conv else None
                    note += f" - even at 100% conversion, projected footfall only covers ~{proj_ff_rem:.0f} of the " \
                            f"{orders_rem:.0f} bills still needed"
                    if extra_ff_needed is not None:
                        note += f" (needs ~{extra_ff_needed:.0f} more footfall)"
            note += ")"
            lines.append(note)
        else:
            lines.append("  (Footfall data not available for this store this month)")

        lines.append("")

    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--sales', required=True)
    ap.add_argument('--targets', required=True)
    ap.add_argument('--footfall', required=True)
    ap.add_argument('--store', default=None)
    ap.add_argument('--out', default=None)
    args = ap.parse_args()

    df = load_sales_csv(args.sales)
    df = df[df['POS location name'].isin(REGION_MAP.keys())]

    days = sorted(df['Day_str'].unique())
    last_day = datetime.datetime.strptime(days[-1], '%Y-%m-%d').date()
    days_elapsed = len(days)
    days_in_month = calendar.monthrange(last_day.year, last_day.month)[1]
    days_remaining = days_in_month - days_elapsed

    print(f"Sales file covers {days_elapsed} day(s), through {last_day.isoformat()}. "
          f"{days_remaining} day(s) remaining in the month.\n")

    targets = load_targets(args.targets)
    new_ff, rep_ff, new_ff_days, rep_ff_days = load_footfall(args.footfall, last_day)

    stores = [args.store] if args.store else sorted(REGION_MAP.keys())
    output_blocks = []
    for store in stores:
        if store not in targets:
            print(f"[skip] No target row found for '{store}' in {args.targets}")
            continue
        sub = df[df['POS location name'] == store]
        new_sub = sub[sub['Segment'] == 'new']
        rep_sub = sub[sub['Segment'] == 'returning']
        achieved = {
            'new_rev': float(new_sub['Revenue'].sum()), 'new_bills': int(new_sub['Order name'].nunique()),
            'rep_rev': float(rep_sub['Revenue'].sum()), 'rep_bills': int(rep_sub['Order name'].nunique()),
        }
        ff = {
            'new_ff': new_ff.get(store, 0.0), 'new_ff_days': new_ff_days,
            'rep_ff': rep_ff.get(store, 0.0), 'rep_ff_days': rep_ff_days,
        }
        output_blocks.append(build_briefing(store, days_elapsed, days_remaining, achieved, targets[store], ff))

    output = "\n\n".join(output_blocks)
    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            f.write(output)
        print(f"Wrote {len(output_blocks)} store briefing(s) to {args.out}")
    else:
        print(output)


if __name__ == '__main__':
    main()
