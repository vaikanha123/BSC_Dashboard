#!/usr/bin/env python3
"""
nps_reasons.py -- Why each NPS detractor / fit "not satisfied" reply on store-audit.html was negative,
from the CX team's call log, the "Master File Escalation" sheet (1Q6maamKWjsnb3L1nV6ZaW1H7qb-9A-vTRMwwuHU33os,
tabs "Master Escalation 2026" (to Oct 2026) and "Nov 2026 onwards"). Pull each tab with the Sheets
connector (get_values on A1:U<rows>) and pass the saved JSON results:

  python scripts/nps_reasons.py --sheet-json master_2026.json nov_onwards.json

Columns by position, as in the sheet's header: A Source (only "NPS" rows, i.e. the WhatsApp survey),
B Ticket ID ("TECH" = NPS survey, "FIT ASSES" = fit survey), C Customer Name, D Date (D-M-YYYY),
G Phone, H Order ID, M Escalation, P Final Update, T Type of Escalation (the reason).

A negative reply in data/store_audit_nps.json / store_audit_fit.json "detail" is matched to the CX row
for the same phone or order, logged from 3 days before to 30 days after the reply (nearest first,
same survey first). Writes data/store_audit_reasons.json and its copy, the "reasons" key of
data/store_audit_data.js: {"nps"|"fit": {"<reply date>|<order>": [type, what the customer told CX, how CX closed it]}}.

The page is public: the notes drop the customer's name, phone-like numbers and emails. Run it after
update_nps.py (a new day's replies need the matching CX rows; CX usually calls within a day or two).
"""
import argparse
import datetime as dt
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'scripts'))
from build_store_audit import stamp_page  # noqa: E402

OUT = os.path.join(ROOT, 'data', 'store_audit_reasons.json')
AUDIT_JS = os.path.join(ROOT, 'data', 'store_audit_data.js')
SURVEYS = {'nps': os.path.join(ROOT, 'data', 'store_audit_nps.json'),
           'fit': os.path.join(ROOT, 'data', 'store_audit_fit.json')}
BOTTOM = 2           # detail slot of a detractor / not satisfied reply
NOTE_MAX = 260


def parse_date(s):
    m = re.match(r'\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4})', s or '')
    if not m:
        return None
    d, mo, y = map(int, m.groups())
    try:
        return dt.date(y, mo, d)
    except ValueError:
        return None


def phone10(s):
    return re.sub(r'\D', '', s or '')[-10:]


def clean_note(text, name):
    text = re.sub(r'^\s*final\s+(clos\w*|update)\s*:?\s*', '', text or '', flags=re.I).strip()
    text = re.sub(r'\S+@\S+', '[email]', text)
    text = re.sub(r'\+?\d[\d\s-]{8,}\d', '[number]', text)
    for w in sorted({w for w in re.split(r'\s+', name or '') if len(w) >= 3}, key=len, reverse=True):
        text = re.sub(r'\b%s\b' % re.escape(w), 'the customer', text, flags=re.I)
    text = re.sub(r'\s+', ' ', text).strip()
    return text if len(text) <= NOTE_MAX else text[:NOTE_MAX - 1].rsplit(' ', 1)[0] + '…'


def load_cx(paths):
    rows = []
    for path in paths:
        with open(path, encoding='utf-8') as f:
            values = json.load(f)['values']
        head = values[0] + [''] * 21 if values else []
        if head[0] == 'Source' and (head[19].strip() != 'Type of Escalation' or head[7].strip() != 'Order ID'):
            sys.exit('STOP: %s: column T is not "Type of Escalation" / H not "Order ID" -- the sheet layout changed' % path)
        for r in values:
            r = r + [''] * (21 - len(r))
            day = parse_date(r[3])
            if r[0].strip() != 'NPS' or not day:
                continue
            rows.append({'kind': 'fit' if 'FIT' in r[1].upper() else 'nps', 'day': day, 'phone': phone10(r[6]),
                         'orders': set(re.findall(r'IND-\d+', r[7])), 'type': r[19].strip() or 'Not tagged',
                         'said': clean_note(r[12], r[2]), 'outcome': clean_note(r[15], r[2])})
    return rows


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--sheet-json', nargs='+', required=True)
    a = ap.parse_args()
    cx = load_cx(a.sheet_json)
    by_phone, by_order = {}, {}
    for e in cx:
        if e['phone']:
            by_phone.setdefault(e['phone'], []).append(e)
        for o in e['orders']:
            by_order.setdefault(o, []).append(e)

    out = {}
    for kind, path in SURVEYS.items():
        with open(path, encoding='utf-8') as f:
            detail = json.load(f).get('detail', {})
        got, total = {}, {}
        for store, months in detail.items():
            for mk, replies in months.items():
                for slot, day, order, phone, _ in replies:
                    if slot != BOTTOM:
                        continue
                    total[mk] = total.get(mk, 0) + 1
                    d = dt.date.fromisoformat(day)
                    cand = {id(e): e for e in by_phone.get(phone10(phone), []) + by_order.get(order, [])}.values()
                    cand = [e for e in cand if -3 <= (e['day'] - d).days <= 30]
                    if not cand:
                        continue
                    e = min(cand, key=lambda e: (e['kind'] != kind, abs((e['day'] - d).days)))
                    out.setdefault(kind, {})['%s|%s' % (day, order)] = [e['type'], e['said'], e['outcome']]
                    got[mk] = got.get(mk, 0) + 1
        for mk in sorted(total):
            print('%s %s: reasons for %d of %d negative replies' % (kind, mk, got.get(mk, 0), total[mk]))

    data = {'source': 'CX Master File Escalation sheet (NPS rows), matched on phone or order', 'nps': out.get('nps', {}),
            'fit': out.get('fit', {})}
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=0)
    with open(AUDIT_JS, encoding='utf-8') as f:
        js = f.read()
    m = re.search(r'^window\.AUDIT_DATA = (.*);\s*$', js, re.MULTILINE | re.DOTALL)
    if not m:
        sys.exit('STOP: window.AUDIT_DATA not found in %s' % AUDIT_JS)
    audit = json.loads(m.group(1))
    audit['reasons'] = data
    with open(AUDIT_JS, 'w', encoding='utf-8', newline='') as f:
        f.write(js[:m.start(1)] + json.dumps(audit, ensure_ascii=False, separators=(',', ':')) + ';\n')
    stamp_page()
    print('OK ->', OUT)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
