#!/usr/bin/env python3
"""
nps_mistag.py -- Collect the survey replies the CX team marked "tagged unlikely by mistake" into
data/nps_mistag.json, which update_nps.py reads to count those replies as promoters.

Source: the CX "Master File Escalation" sheet (1Q6maamKWjsnb3L1nV6ZaW1H7qb-9A-vTRMwwuHU33os), tabs
"Master Escalation 2026" (to Oct 2026) and "Nov 2026 onwards". Pull them with the Sheets connector
(get_values on A:T) and pass the saved JSON results. Columns by position, as in the sheet's header:
A Source (only "NPS" rows: the WhatsApp survey), D Date (D-M-YYYY text), H Order ID, T Type of Escalation. A range that starts below the
header works too.

Only the order ID and the escalation date are kept (the repo is public; no names or phones).
Existing entries are kept, so older tabs need not be pulled again.

Usage:
  python scripts/nps_mistag.py --sheet-json master_oct.json nov_tab.json
"""
import argparse
import datetime as dt
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'data', 'nps_mistag.json')
TAG = 'tagged unlikely by mistake'


def parse_date(s):
    m = re.match(r'\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})', s or '')
    if not m:
        return None
    d, mo, y = map(int, m.groups())
    try:
        return dt.date(y, mo, d).isoformat()
    except ValueError:
        return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--sheet-json', nargs='+', required=True)
    a = ap.parse_args()

    old = {'orders': []}
    if os.path.exists(OUT):
        with open(OUT, encoding='utf-8') as f:
            old = json.load(f)
    old['orders'] = [e for e in old['orders'] if e.get('source', 'NPS') == 'NPS']
    seen = {(e['order'], e['date']) for e in old['orders']}
    added = 0
    for path in a.sheet_json:
        with open(path, encoding='utf-8') as f:
            values = json.load(f)['values']
        if values and values[0][:1] == ['Source'] and (values[0] + [''] * 20)[19].strip() != 'Type of Escalation':
            sys.exit('STOP: %s: column T is not "Type of Escalation" -- the sheet layout changed' % path)
        for r in values:
            r = r + [''] * (20 - len(r))
            if r[19].strip().lower() != TAG or r[0].strip() != 'NPS':   # survey replies only, not JudgeMe/Google
                continue
            day = parse_date(r[3])
            for order in re.findall(r'IND-\d+', r[7]):
                if day and (order, day) not in seen:
                    seen.add((order, day))
                    old['orders'].append({'order': order, 'date': day, 'source': r[0].strip()})
                    added += 1
    old['orders'].sort(key=lambda e: (e['date'], e['order']))
    old['note'] = ('Replies the CX team tagged "tagged unlikely by mistake" in the Master File Escalation sheet; '
                   'update_nps.py counts a matching negative reply as a promoter / very satisfied.')
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(old, f, ensure_ascii=False, indent=1)
    print('OK. %d new, %d total mis-tagged orders -> %s' % (added, len(old['orders']), OUT))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
