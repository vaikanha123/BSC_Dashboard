"""
generate_briefing_data.py -- Computes per-store daily briefing numbers (required New/Repeat AOV,
repeat customers needed, conversion needed for the rest of the month, ASP/UPT diagnosis, per-store
NPS/audit score, and a stylist-wise breakdown) and writes them into store-briefing.html as
`const BRIEFING_DATA = {...};`, the same replace-a-JS-constant pattern used by
update_main_dashboard.py / update_tracker.py.

Usage:
  python3 generate_briefing_data.py --html store-briefing.html --sales sales.csv \
      --targets "Daywise Targets Sep-26.xlsx" --footfall footfall.csv [--aov-targets "Targets for Sep.xlsx"] \
      [--compliance "Compliance Score Sheet - July to Sep MTD 26.xlsx"] [--index index.html]

  --sales       : cumulative current-month sales CSV.
  --targets     : Daywise Targets Excel for the current month (POS location name, Date, New
                   Revenue, Repeat Revenue, New orders, Repeat orders, ...). This is the
                   authoritative source for the full-month revenue/order targets used in the
                   "required AOV" / "orders needed" math -- see BSC_Dashboard_Runbook.md.
  --footfall    : CSV export of the "<Month> FF" tab of the footfall tracker Google Sheet -- see
                   the bsc-footfall-tracker-sheet reference for the URL and how to pull it without
                   File > Download access (its raw export needs reconstructing into New FF/Repeat
                   FF/TOTAL FF blocks; this script expects the already-reconstructed form).
  --aov-targets : optional. The separate "Targets for <Mon>.xlsx" file (Store, Target, New Rev
                   Tar, Rep Rev Tar, New AOV Target, Repeat AOV Target, New Bills, Repeat Bill).
                   As of 2026-09, the Rev/Bills columns in this file don't reconcile against the
                   Daywise Targets file or against each other (AOV x Bills != Rev Tar) -- period
                   unclear, unresolved with Vaibhav. Only the two AOV Target columns are pulled
                   from this file, as a labeled reference figure ("HO Target AOV") shown alongside
                   the calculated required AOV -- NOT used in any calculation. Store names in this
                   file are normalized via AOV_TARGET_STORE_MAP in bsc_common.py.
  --compliance  : optional. The "Compliance Score Sheet - <range>.xlsx" file (Store Name, then one
                   column per period, e.g. "July MTD", "Aug W1"..."Aug MTD", "Sep W1", ...). Uses
                   the "Aug MTD" column as the closed-month baseline and whichever column is
                   furthest right as "current" -- so a later week's column slots in with no code
                   change. Store names normalized via COMPLIANCE_STORE_MAP below.
  --index       : optional, default 'index.html' next to --html. Pulls SEED_NPS.byStore (current
                   NPS per store) and BASELINE.prevMonthStoreUPT/prevMonthStoreAOV (last month's
                   ASP/UPT, for diagnosing whether a shortfall in AOV is an ASP problem or a UPT
                   problem) straight from the main dashboard rather than re-entering them -- it's
                   already the single source of truth for both.

Same calculation logic as scripts/generate_store_briefing.py (that script remains the
plain-text/WhatsApp-friendly version of the same output); see its docstring for the formulas.
"""
import argparse
import calendar
import csv
import datetime
import json
import os

import openpyxl

from bsc_common import REGION_MAP, AOV_TARGET_STORE_MAP, load_sales_csv, replace_const, syntax_check_html_js


# Compliance Score Sheet uses its own store-name spellings (shortened/renamed vs REGION_MAP) --
# built the same way AOV_TARGET_STORE_MAP was, by matching that sheet's own Store Name column.
# Add new entries here if a similar sheet is used later, don't assume the two alias maps match.
COMPLIANCE_STORE_MAP = {
    'Ambience Vasnt Kunj': 'Ambience Vasant Kunj', 'Andheri': 'Andheri', 'DLF Midtown': 'DLF Midtown - Moti Nagar',
    'Express Avenue': 'Express Avenue', 'Gurugram': 'Gurugram', 'Indiranagar': 'Indiranagar',
    'Inorbit mall': 'Inorbit mall Hyderabad', 'Inorbit Mall Malad West': 'Inorbit Mall Malad West',
    'Jaipur': 'Jaipur Store', 'Jayanagar': 'Jayanagar', 'Jubilee Hills': 'Jubilee Hills', 'Juhu': 'Juhu Store',
    'Kalaghoda, Fort': 'Kalaghoda, Fort', 'Kemps Corner': 'Kemps Corner', 'Khan Market': 'Khan Market',
    'KNK Chennai': 'KNK Chennai', 'Kochi': 'Kochi Store', 'Koregaon Park': 'Koregaon Park',
    'Lakeshore mall, Kukatpally': 'LakeShore Mall', 'Lavelle Road': 'Lavelle Road',
    'Mall of India': 'Mall of India, Noida', 'Oberoi Mall': 'Oberoi Mall Store', 'Pali Hill, Khar': 'Pali Hill, Bandra',
    'Phoenix Lucknow': 'Phoenix Palassio', 'Phoenix Viman Nagar': 'PMC Viman Nagar Pune',
    'PMC Kurla': 'Phoenix Marketcity Kurla', 'PMC Whitefield': 'Phoenix Marketcity, Whitefield',
    'R.K. Salai': 'R.K. Salai', 'Select City Mall': 'Select City', 'Shakespear Sarani': 'Shakespearesarani',
    'Sharath City': 'Sarath City-Hyderabad', 'Sindhu Bhavan Marg': 'Sindhu Bhavan Marg',
    'Sky City Borivali': 'Oberoi Sky City', 'South Ex.': 'South Ex.', 'Vegas Mall Dwarka': 'Vegas Dwarka',
    'Viviana Mall': 'Viviana Mall',
}


def load_targets(path, cutoff_date):
    """Returns (totals, todate): totals is the full-month per-store target (used for the "rest of
    month" required-AOV/bills math, unchanged); todate is the target summed only through
    cutoff_date -- the correct denominator for "Target vs Achievement" on the briefing, which must
    compare like-for-like (target-to-date vs achieved-to-date), not a mid-month achieved figure
    against the full month's target."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb[wb.sheetnames[0]]
    totals = {}
    todate = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        store, date, new_rev, rep_rev, new_orders, rep_orders = row[0], row[1], row[2], row[3], row[4], row[5]
        if store is None or date is None:
            continue
        t = totals.setdefault(store, {'new_rev': 0.0, 'rep_rev': 0.0, 'new_orders': 0, 'rep_orders': 0})
        t['new_rev'] += new_rev or 0
        t['rep_rev'] += rep_rev or 0
        t['new_orders'] += new_orders or 0
        t['rep_orders'] += rep_orders or 0
        d = date.date() if hasattr(date, 'date') else date
        if d <= cutoff_date:
            todate[store] = todate.get(store, 0.0) + (new_rev or 0) + (rep_rev or 0)
    return totals, todate


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


def load_compliance(path):
    """Store -> {aug: pct 0-100 or None, current: pct 0-100 or None, currentLabel: str}."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    aug_idx = None
    for i, h in enumerate(header):
        if h and 'Aug MTD' in str(h):
            aug_idx = i
    current_idx = len(header) - 1
    current_label = header[current_idx]

    def pct(v):
        return round(v * 100, 1) if isinstance(v, (int, float)) else None

    out = {}
    for row in rows[1:]:
        store = row[0]
        if store is None:
            continue
        canonical = COMPLIANCE_STORE_MAP.get(str(store).strip())
        if canonical is None:
            print(f"[warn] '{store}' in compliance sheet has no entry in COMPLIANCE_STORE_MAP -- skipped")
            continue
        out[canonical] = {
            'aug': pct(row[aug_idx]) if aug_idx is not None else None,
            'current': pct(row[current_idx]) if current_idx < len(row) else None,
            'currentLabel': str(current_label) if current_label else None,
        }
    return out


def _extract_field_object(content, field_name):
    """Find `<field_name>: {...}` anywhere in the text and brace-balance out the object literal --
    unlike bsc_common.extract_const, this works on a field nested inside a larger object (like
    BASELINE.prevMonthStoreUPT) whose surrounding object has // comments breaking naive JSON
    parsing of the whole thing. The extracted substring itself is pure data, no comments."""
    marker = f'{field_name}:'
    start = content.index(marker)
    val_start = content.index('{', start)
    depth = 0
    i = val_start
    while i < len(content):
        if content[i] == '{':
            depth += 1
        elif content[i] == '}':
            depth -= 1
            if depth == 0:
                break
        i += 1
    return json.loads(content[val_start:i + 1])


def load_index_context(path):
    """Pull SEED_NPS.byStore and BASELINE.prevMonthStoreUPT/prevMonthStoreAOV straight out of the
    main dashboard rather than re-entering them -- it's already the single source of truth for
    both, and prevMonth*'s "last closed month" ASP/UPT is exactly what's needed to tell a manager
    whether a store's AOV shortfall this month is an ASP problem or a UPT problem."""
    with open(path, encoding='utf-8') as f:
        html = f.read()
    marker = 'const SEED_NPS = '
    start = html.index(marker) + len(marker)
    depth = 0
    i = start
    while i < len(html):
        if html[i] == '{':
            depth += 1
        elif html[i] == '}':
            depth -= 1
            if depth == 0:
                break
        i += 1
    nps = json.loads(html[start:i + 1])
    return {
        'npsByStore': nps.get('byStore', {}),
        'prevMonthStoreUPT': _extract_field_object(html, 'prevMonthStoreUPT'),
        'prevMonthStoreAOV': _extract_field_object(html, 'prevMonthStoreAOV'),
    }


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


def segment_block(rev_t, orders_t, rev_a, orders_a, units_a, ff_achieved, ff_days, days_remaining, ho_aov, ho_bills=None):
    # Two independently-confirmed sources, each trusted for what it actually states -- don't try
    # to make them reconcile with each other:
    #  - Revenue target (rev_t) = Daywise Targets file. Confirmed correct by Vaibhav's own
    #    store-by-store revenue table on 2026-09-09 (matched exactly).
    #  - Bill-count target (ho_bills) = "Targets for <Mon>.xlsx"'s own New Bills/Repeat Bill
    #    column. Confirmed correct for Ambience Vasant Kunj (250, not Daywise's own 270 orders
    #    column) on 2026-09-09 -- prefer it over Daywise's bill-count column when available.
    #  - AOV Target (ho_aov), from the same file, is a separate reference figure shown alongside
    #    the calculated required AOV. It is NOT assumed to multiply out with rev_t or ho_bills --
    #    an earlier attempt to "reconcile" it by overriding rev_t when it didn't match was wrong
    #    and got corrected; a required AOV far from ho_aov is a real signal (this store's revenue
    #    target implies a different average ticket than its stated AOV target), not a data bug.
    if ho_bills:
        orders_t = ho_bills

    rev_rem = rev_t - rev_a
    orders_rem = orders_t - orders_a
    cur_aov = (rev_a / orders_a) if orders_a else None
    cur_asp = (rev_a / units_a) if units_a else None
    cur_upt = (units_a / orders_a) if orders_a else None

    block = {
        'targetRev': round(rev_t, 2), 'targetOrders': orders_t, 'achievedRev': round(rev_a, 2), 'achievedBills': orders_a,
        'achievedUnits': round(units_a, 1),
        'currentAOV': round(cur_aov, 2) if cur_aov is not None else None,
        'currentASP': round(cur_asp, 2) if cur_asp is not None else None,
        'currentUPT': round(cur_upt, 3) if cur_upt is not None else None,
        'hoTargetAOV': ho_aov,
        'remainingRev': round(rev_rem, 2), 'remainingOrders': orders_rem,
        'noTarget': (rev_t == 0 and rev_a == 0),
    }

    if rev_rem <= 0 and not block['noTarget']:
        block['status'] = 'achieved'
        block['surplus'] = round(-rev_rem, 2)
    elif orders_rem > 0 and rev_rem > 0:
        block['requiredAOV'] = round(rev_rem / orders_rem, 2)
        block['status'] = 'on-track'
    else:
        block['status'] = 'orders-hit-revenue-short'

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


def build_stylist_breakdown(sub_df):
    """Per-stylist new/repeat rev, bills, units, aov, asp, upt at this one store, for whoever
    worked here this month. No stylist-level target exists anywhere in the pipeline (only
    store-level), so this is plain performance, not a vs-target comparison -- the store manager
    reads it against the store-level targets/diagnosis shown above it on the page."""
    styl_df = sub_df[sub_df['Stylist'] != '']
    out = []
    for name, g in styl_df.groupby('Stylist'):
        entry = {'name': name}
        for label, key in (('new', 'new'), ('rep', 'returning')):
            seg = g[g['Segment'] == key]
            rev = float(seg['Revenue'].sum())
            bills = int(seg['Order name'].nunique())
            units = float(seg['Qty'].sum())
            entry[label] = {
                'rev': round(rev, 2), 'bills': bills, 'units': round(units, 1),
                'aov': round(rev / bills, 2) if bills else None,
                'asp': round(rev / units, 2) if units else None,
                'upt': round(units / bills, 3) if bills else None,
            }
        entry['totalRev'] = round(entry['new']['rev'] + entry['rep']['rev'], 2)
        entry['totalBills'] = entry['new']['bills'] + entry['rep']['bills']
        out.append(entry)
    out.sort(key=lambda e: e['totalRev'], reverse=True)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--html', required=True)
    ap.add_argument('--sales', required=True)
    ap.add_argument('--targets', required=True)
    ap.add_argument('--footfall', required=True)
    ap.add_argument('--aov-targets', default=None)
    ap.add_argument('--compliance', default=None)
    ap.add_argument('--index', default=None)
    args = ap.parse_args()

    df = load_sales_csv(args.sales)
    df = df[df['POS location name'].isin(REGION_MAP.keys())]

    days = sorted(df['Day_str'].unique())
    last_day = datetime.datetime.strptime(days[-1], '%Y-%m-%d').date()
    days_elapsed = len(days)
    days_in_month = calendar.monthrange(last_day.year, last_day.month)[1]
    days_remaining = days_in_month - days_elapsed

    targets, targets_todate = load_targets(args.targets, last_day)
    new_ff, rep_ff, new_ff_days, rep_ff_days = load_footfall(args.footfall, last_day)
    aov_targets = load_aov_targets(args.aov_targets) if args.aov_targets else {}
    compliance = load_compliance(args.compliance) if args.compliance else {}

    index_path = args.index or os.path.join(os.path.dirname(os.path.abspath(args.html)) or '.', 'index.html')
    index_ctx = load_index_context(index_path) if os.path.exists(index_path) else {'npsByStore': {}, 'prevMonthStoreUPT': {}, 'prevMonthStoreAOV': {}}

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
            float(new_sub['Qty'].sum()),
            new_ff.get(store, 0.0), new_ff_days, days_remaining, ho.get('new_aov'), ho.get('new_bills'))
        rep_block = segment_block(
            t['rep_rev'], t['rep_orders'], float(rep_sub['Revenue'].sum()), int(rep_sub['Order name'].nunique()),
            float(rep_sub['Qty'].sum()),
            rep_ff.get(store, 0.0), rep_ff_days, days_remaining, ho.get('rep_aov'), ho.get('rep_bills'))

        prev_upt = index_ctx['prevMonthStoreUPT'].get(store)
        prev_aov = index_ctx['prevMonthStoreAOV'].get(store)
        prev_month = None
        if prev_upt and prev_aov:
            prev_month = {
                'newUpt': prev_upt.get('newUpt'), 'repUpt': prev_upt.get('repUpt'),
                'newAsp': round(prev_aov['newAov'] / prev_upt['newUpt'], 2) if prev_upt.get('newUpt') else None,
                'repAsp': round(prev_aov['repAov'] / prev_upt['repUpt'], 2) if prev_upt.get('repUpt') else None,
                'newAov': prev_aov.get('newAov'), 'repAov': prev_aov.get('repAov'),
            }

        stores_out[store] = {
            'region': REGION_MAP[store],
            # Sum the two segments' own (possibly-overridden) targets rather than re-reading
            # Daywise's raw new_rev+rep_rev, so this always agrees with what the segment cards show.
            'monthTarget': round(new_block['targetRev'] + rep_block['targetRev'], 2),
            'monthTargetToDate': round(targets_todate.get(store, 0.0), 2),
            'achievedTotal': round(float(sub['Revenue'].sum()), 2),
            'new': new_block, 'rep': rep_block,
            'prevMonth': prev_month,
            'nps': index_ctx['npsByStore'].get(store),
            'audit': compliance.get(store),
            'stylists': build_stylist_breakdown(sub),
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
          f"through {last_day.isoformat()} ({days_remaining} days remaining). "
          f"Compliance data: {'yes' if compliance else 'no'}. NPS from index.html: {'yes' if index_ctx['npsByStore'] else 'no'}.")


if __name__ == '__main__':
    main()
