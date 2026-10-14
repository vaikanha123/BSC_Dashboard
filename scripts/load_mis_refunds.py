#!/usr/bin/env python3
"""
load_mis_refunds.py -- Write MIS_REFUND_CN into index.html from the finance MIS workbook, for the
Refunds & CN tab's long-run trend (finance-booked refunds and credit notes as a % of revenue).

Reads the per-store blocks of the "Revenue Info (FY26)" and "Revenue Info (FY27)" sheets: row 4
names each block (store, "Online", "Corporate Order", ..., "GRAND TOTAL"), row 5 holds the month
dates, and these rows hold the figures:
  FY27: 14 Merchandise Gross Revenue, 11 GV (Credit Note Issued), 12 Refunds and Cancellations
  FY26: 17 Gross Revenue,             13 GV (Credit Note Issued), 14 Refunds and Cancellations
Online = the Online, Myntra Online, international online and Corporate Order blocks (the same
"not a store" definition the dashboard uses); offline = every other block. Each month's
online + offline must equal the GRAND TOTAL block, or the script stops.

Usage:
  python scripts/load_mis_refunds.py --html index.html --mis "MIS - Sept 2026.xlsx"
"""
import argparse
import datetime as dt
import json
import os
import sys
from collections import defaultdict

import openpyxl

from bsc_common import replace_const, syntax_check_html_js

ONLINE_BLOCKS = {'Online', 'Myntra Online', 'Intl - New York, USA - Online', 'Corporate Order'}
SKIP_BLOCKS = {'Particulars', 'For SSG', 'Offline', 'STORES - GRAND TOTAL'}
SHEETS = {
    'Revenue Info (FY26)': {'rev': 17, 'cn': 13, 'ref': 14},
    'Revenue Info (FY27)': {'rev': 14, 'cn': 11, 'ref': 12},
}
ROW_LABELS = {'rev': ('gross revenue', 'merchandise gross revenue'), 'cn': ('gv (credit note issued)',),
              'ref': ('refunds and cancellations',)}


def read_sheet(ws, rows):
    R = {i: r for i, r in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), start=1)}
    for key, r in rows.items():
        label = str(R[r][0] or '').strip().lower()
        if label not in ROW_LABELS[key]:
            sys.exit(f'STOP: {ws.title} row {r} is {R[r][0]!r}, expected {ROW_LABELS[key]} -- the MIS layout changed')
    r4, r5 = R[4], R[5]
    grand_seen = False
    out = defaultdict(lambda: defaultdict(float))  # month -> {on_rev, off_rev, gt_rev, ...}
    for col, name in enumerate(r4):
        if not name or name in SKIP_BLOCKS:
            continue
        name = str(name).strip()
        if grand_seen:  # summary blocks to the right of GRAND TOTAL (e.g. a second "Online") are not stores
            continue
        if name == 'GRAND TOTAL':
            grand_seen = True
        ch = 'gt' if name == 'GRAND TOTAL' else ('on' if name in ONLINE_BLOCKS else 'off')
        # The block's name sits mid-block; its month columns are the dated cells around it.
        seen = set()
        for j in range(max(1, col - 7), min(col + 8, len(r5))):
            d = r5[j]
            if not isinstance(d, dt.datetime):
                continue
            m = d.strftime('%Y-%m')
            if m in seen:
                continue
            seen.add(m)
            for key, r in rows.items():
                v = R[r][j]
                out[m][f'{ch}_{key}'] += v if isinstance(v, (int, float)) else 0
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--html', required=True)
    ap.add_argument('--mis', required=True)
    a = ap.parse_args()

    wb = openpyxl.load_workbook(a.mis, read_only=True, data_only=True)
    months = {}
    for sn, rows in SHEETS.items():
        for m, v in read_sheet(wb[sn], rows).items():
            if v.get('gt_rev'):
                months[m] = v
    out = []
    for m in sorted(months):
        v = months[m]
        for k in ('rev', 'cn', 'ref'):
            if abs(v[f'on_{k}'] + v[f'off_{k}'] - v[f'gt_{k}']) > 1:
                sys.exit(f"STOP: {m} {k}: online + offline {v[f'on_{k}'] + v[f'off_{k}']:,.0f} != GRAND TOTAL {v[f'gt_{k}']:,.0f}")
        out.append({'month': m, **{f'{ch}{k.capitalize()}': round(v[f'{ch}_{k}']) for ch in ('on', 'off')
                                   for k in ('rev', 'ref', 'cn')}})
    payload = {'source': os.path.basename(a.mis), 'months': out}

    with open(a.html, encoding='utf-8') as f:
        content = f.read()
    if 'const MIS_REFUND_CN = ' in content:
        content = replace_const(content, 'MIS_REFUND_CN', json.dumps(payload))
    else:
        anchor = 'const PREV_MONTH_REFUND_CN = '
        i = content.index(anchor)
        content = content[:i] + 'const MIS_REFUND_CN = ' + json.dumps(payload) + ';\n' + content[i:]
    with open(a.html, 'w', encoding='utf-8') as f:
        f.write(content)
    syntax_check_html_js(a.html)

    print(f"OK. MIS_REFUND_CN: {len(out)} months, {out[0]['month']} -> {out[-1]['month']} from {payload['source']}")
    for r in out:
        on, off = r['onRev'], r['offRev']
        tot = on + off
        print(f"  {r['month']}  online refund {r['onRef']/on*100:4.1f}% CN {r['onCn']/on*100:4.1f}%  |  offline refund "
              f"{r['offRef']/off*100:4.1f}% CN {r['offCn']/off*100:4.1f}%  |  total refund {(r['onRef']+r['offRef'])/tot*100:4.1f}% "
              f"CN {(r['onCn']+r['offCn'])/tot*100:4.1f}%")


if __name__ == '__main__':
    main()
