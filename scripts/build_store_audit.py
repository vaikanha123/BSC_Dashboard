"""Build data/store_audit_data.js for store-audit.html from the Store Audit Google Sheet.

The sheet (1On62S5C0mXyl6XDX0_Lsfj7_ZuGoBfwbGN-N6P5XZbA) is not link-shared, so the four
source tabs are first saved as JSON into a raw folder (kept out of the repo, the remarks
carry customer phone numbers and emails):

  tracker.json   'Audit Tracker'                    -> list of rows
  ba.json        'BA - Response Sheet'              -> {"rows": [...]}
  sales.json     'Sales Response Sheet - Stores'    -> {"rows": [...]}
  inbound.json   'Inbound Response Sheet - Stores'  -> {"rows": [...]}

Usage:
  python scripts/build_store_audit.py --raw-dir <folder> [--pulled 2026-10-12] [--dump-notes 2026-10] [--verbatim]

The auditor's remarks are left out of the data file unless --verbatim is passed: they can name
customers, and the page is served publicly. --dump-notes prints a month's remarks for writing
the feedback summaries.

Store-wise NPS comes from the WhatsApp survey workbook (its "Data" tab, one row per response):
  python scripts/build_store_audit.py --raw-dir <folder> --nps "NPS_Fit_Assessment_....xlsx"
That writes data/store_audit_nps.json (counts only, no customer details), which later builds reuse.

Weeks follow the tracker's own Period column ("Oct W1"); call and behaviour audits are
placed in the tracker period that their interaction date falls in.
Feedback summaries written by hand live in data/store_audit_feedback.json and are merged in.
"""
import argparse
import collections
import datetime as dt
import json
import os
import re
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(REPO, 'data', 'store_audit_data.js')
FEEDBACK = os.path.join(REPO, 'data', 'store_audit_feedback.json')
NPS = os.path.join(REPO, 'data', 'store_audit_nps.json')
PAGE = os.path.join(REPO, 'store-audit.html')

# The survey file starts on 26 Aug 2026, so August is only six days; it is left off the page.
NPS_SKIP_MONTHS = {'2026-08'}

# Store names in the NPS workbook (Shopify POS locations) -> names used in the audit tracker.
NPS_STORES = {
    'Ambience Vasant Kunj': 'Ambience Vasnt Kunj', 'DLF Midtown - Moti Nagar': 'DLF Midtown',
    'Inorbit mall Hyderabad': 'Inorbit mall', 'Jaipur Store': 'Jaipur', 'Juhu Store': 'Juhu',
    'Kochi Store': 'Kochi', 'LakeShore Mall': 'Lakeshore mall, Kukatpally', 'Mall of India, Noida': 'Mall of India',
    'Oberoi Mall Store': 'Oberoi Mall', 'Oberoi Sky City': 'Sky City Borivali', 'PMC Viman Nagar Pune': 'Phoenix Viman Nagar',
    'Pali Hill, Bandra': 'Pali Hill, Khar', 'Phoenix Marketcity Kurla': 'PMC Kurla',
    'Phoenix Marketcity, Whitefield': 'PMC Whitefield', 'Phoenix Palassio': 'Phoenix Lucknow',
    'Sarath City-Hyderabad': 'Sharath City', 'Select City': 'Select City Mall', 'Shakespearesarani': 'Shakespear Sarani',
    'Vegas Dwarka': 'Vegas Mall Dwarka',
}

MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
               'September', 'October', 'November', 'December']
COMP_PARAMS = ['Footfall', 'NPS Calling', 'At Store Calling', 'Delay Calling', 'Pre Confirmed',
               'Alteration', 'Staff Grooming']
COMP_KEYS = [p.lower() for p in COMP_PARAMS]
WEEKS = ['W1', 'W2', 'W3', 'W4', 'W5']


def load(raw_dir, name):
    with open(os.path.join(raw_dir, name + '.json'), encoding='utf-8') as f:
        j = json.load(f)
    return j['rows'] if isinstance(j, dict) else j


def tracker_date(s):
    m = re.match(r'\s*(\d{1,2})[-\s]+([A-Za-z]+)[-\s]+(\d{2,4})', s or '')
    if not m or m.group(2)[:3].lower() not in MONTHS:
        return None
    y = int(m.group(3))
    return dt.date(y + 2000 if y < 100 else y, MONTHS.index(m.group(2)[:3].lower()) + 1, int(m.group(1)))


def audit_date(s):
    m = re.match(r'\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$', s or '')
    if not m:
        return None
    try:
        return dt.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    except ValueError:
        return None


def period(label, year):
    """'Oct W1' / 'June W4' -> ('2026-10', 'W1')."""
    m = re.match(r'\s*([A-Za-z]{3})[a-z]*\s*W(\d)', label or '')
    if not m or m.group(1).lower() not in MONTHS:
        return None
    return '%d-%02d' % (year, MONTHS.index(m.group(1).lower()) + 1), 'W' + m.group(2)


def num(s):
    try:
        return float(str(s).replace('%', '').strip())
    except ValueError:
        return None


def pct(s):
    s = str(s or '').strip()
    return num(s) if s.endswith('%') else None


def clean(text):
    """Auditor remark with customer phone numbers and emails removed."""
    t = re.sub(r'[\w.+-]+@[\w-]+\.[\w.-]+', '', text or '')
    t = re.sub(r'(?<!\d)(?:\+?91[\s-]?)?[6-9]\d{9}(?!\d)', '', t)
    t = re.sub(r'^[\s\-:,/]+', '', t)
    return re.sub(r'[ \t]+', ' ', t).strip()


def short(d):
    return '%d %s' % (d.day, MONTH_NAMES[d.month - 1][:3])


def build_nps(xlsx):
    """Promoter / neutral / detractor counts per store and calendar month, store responses only."""
    import openpyxl
    rows = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)['Data'].iter_rows(values_only=True)
    head = [str(h or '').strip() for h in next(rows)]
    col = {h: i for i, h in enumerate(head)}
    slot = {'Promoter': 0, 'Neutral': 1, 'Detractor': 2}
    stores, dates = {}, collections.defaultdict(list)
    for r in rows:
        if r[col['Campaign']] != 'NPS' or r[col['Counted in scores']] != 'Yes' or r[col['Channel']] != 'Offline (stores)':
            continue
        d = r[col['Response date']]
        mk = '%d-%02d' % (d.year, d.month)
        if mk in NPS_SKIP_MONTHS:
            continue
        name = str(r[col['Store']]).strip()
        counts = stores.setdefault(NPS_STORES.get(name, name), {}).setdefault(mk, [0, 0, 0])
        counts[slot[r[col['Class']]]] += 1
        dates[mk].append(d)
    months = []
    for mk in sorted(dates):
        lo, hi = min(dates[mk]), max(dates[mk])
        months.append({'key': mk, 'label': MONTH_NAMES[lo.month - 1][:3], 'range': '%d–%s' % (lo.day, short(hi))})
    return {'source': os.path.basename(xlsx), 'months': months, 'stores': stores}


def build(raw_dir, pulled):
    tracker = load(raw_dir, 'tracker')[1:]
    report = collections.Counter()

    # --- Compliance, and the date -> period map the other three blocks reuse ---
    comp = {}
    notes = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(list)))
    by_date = collections.defaultdict(collections.Counter)
    stores = set()
    for r in tracker:
        r = [str(x) for x in r] + [''] * (13 - len(r))
        d = tracker_date(r[0])
        store = r[4].strip()
        if not d or not store:
            report['tracker rows without date/store'] += 1
            continue
        per = period(r[3], d.year)
        if not per:
            report['tracker rows without period'] += 1
            continue
        stores.add(store)
        by_date[d][per] += 1
        key = re.sub(r'\s+', ' ', r[5].strip().lower().replace('-', ' '))
        if key not in COMP_KEYS:
            report['tracker parameter not on the card: ' + r[5].strip()] += 1
            continue
        pi = COMP_KEYS.index(key)
        s, m = num(r[6]), num(r[9])
        if s is not None and m is not None:
            cell = comp.setdefault(per[0], {}).setdefault(store, {}).setdefault(per[1], [None] * len(COMP_PARAMS))
            cell[pi] = [(cell[pi] or [0, 0])[0] + s, (cell[pi] or [0, 0])[1] + m]
        text = '\n'.join(clean(x) for x in (r[7], r[12]) if x.strip() and not re.fullmatch(r'\s*\d+%?\s*', x))
        if text:
            notes[per[0]][store]['comp'].append({'d': short(d), 'w': per[1], 'p': COMP_PARAMS[pi], 't': text, '_s': d.isoformat()})

    date_map = {d: c.most_common(1)[0][0] for d, c in by_date.items()}
    week_votes = collections.defaultdict(collections.Counter)
    for d, per in date_map.items():
        week_votes[d - dt.timedelta(days=d.weekday())][per] += by_date[d][per]
    week_map = {mon: c.most_common(1)[0][0] for mon, c in week_votes.items()}
    last_mon = max(week_map)

    def place(d):
        if d in date_map:
            return date_map[d]
        mon = d - dt.timedelta(days=d.weekday())
        if mon in week_map:
            return week_map[mon]
        if mon > last_mon:  # weeks the tracker has not reached yet: keep counting within the month
            mk, wk = week_map[last_mon]
            n = int(wk[1]) + (mon - last_mon).days // 7
            if n <= 5 and '%d-%02d' % (mon.year, mon.month) == mk:
                return mk, 'W%d' % n
        return None

    # --- Behaviour, inbound and sales audits ---
    audits = {}
    for key, name in (('ba', 'ba'), ('sal', 'sales'), ('inb', 'inbound')):
        rows = load(raw_dir, name)
        head = [str(x).strip() for x in rows[1]]
        score_cols = [i for i, h in enumerate(head) if h == 'Parameter Score']
        remark_col = max(i for i, h in enumerate(head) if h in ('Remarks', 'Additonal Comment', 'Additional Comment'))
        out = audits[key] = {}
        for r in rows[2:]:
            r = [str(x) for x in r] + [''] * (len(head) + 1 - len(r))
            store = r[3].strip()
            if not store:
                continue
            d = audit_date(r[1])
            if not d:
                report['%s rows with unreadable interaction date' % name] += 1
                continue
            if d > pulled:
                report['%s rows dated in the future (dropped)' % name] += 1
                continue
            per = place(d)
            if not per:
                report['%s rows outside the tracker weeks' % name] += 1
                continue
            if store not in stores:
                report['%s store not in tracker: %s' % (name, store)] += 1
                stores.add(store)
            cell = out.setdefault(per[0], {}).setdefault(store, {}).setdefault(per[1], [0, 0, 0] + [0, 0] * len(score_cols))
            cell[0] += 1
            f = pct(r[5])
            if f is not None:
                cell[1] += f
                cell[2] += 1
            for i, c in enumerate(score_cols):
                v = pct(r[c])
                if v is not None:
                    cell[3 + 2 * i] += v
                    cell[4 + 2 * i] += 1
            text = clean(r[remark_col])
            if text:
                notes[per[0]][store][key].append({'d': short(d), 'w': per[1], 'who': r[4].strip(), 'sc': r[5].strip(), 't': text, '_s': d.isoformat()})

    for mk in notes:
        for st in notes[mk]:
            for k in notes[mk][st]:
                notes[mk][st][k].sort(key=lambda n: n.pop('_s'))

    # --- Months and the date range each week covers ---
    ranges = collections.defaultdict(list)
    for d, per in date_map.items():
        ranges[per].append(d)
    month_keys = sorted(set(comp) | {m for a in audits.values() for m in a})
    months = []
    for mk in month_keys:
        y, m = map(int, mk.split('-'))
        wk = {}
        for w in WEEKS:
            ds = ranges.get((mk, w))
            if ds:
                lo = min(ds) - dt.timedelta(days=min(ds).weekday())
                hi = lo + dt.timedelta(days=6)
                if mk == '2026-05':  # May's periods did not follow calendar weeks; show the audited dates
                    lo, hi = min(ds), max(ds)
                wk[w] = short(lo) + '–' + short(hi) if lo.month != hi.month else '%d–%s' % (lo.day, short(hi))
        months.append({'key': mk, 'label': '%s %d' % (MONTH_NAMES[m - 1], y), 'weeks': wk})

    feedback = {}
    if os.path.exists(FEEDBACK):
        with open(FEEDBACK, encoding='utf-8') as f:
            feedback = json.load(f)

    nps = {'months': [], 'stores': {}}
    if os.path.exists(NPS):
        with open(NPS, encoding='utf-8') as f:
            nps = json.load(f)
        for name in sorted(set(nps['stores']) - stores):
            report['NPS store not in tracker: ' + name] += 1

    data = {'pulledAt': pulled.isoformat(), 'months': months, 'stores': sorted(stores, key=str.lower),
            'compParams': COMP_PARAMS, 'comp': comp, 'ba': audits['ba'], 'inb': audits['inb'], 'sal': audits['sal'],
            'notes': notes, 'feedback': feedback, 'nps': nps}
    return data, report


def dump_notes(data, mk):
    labels = {'comp': 'COMPLIANCE', 'ba': 'BEHAVIOUR', 'inb': 'INBOUND', 'sal': 'SALES'}
    for store in data['stores']:
        blocks = data['notes'].get(mk, {}).get(store)
        if not blocks:
            continue
        print('=== ' + store)
        for k in ('comp', 'ba', 'inb', 'sal'):
            for n in blocks.get(k, []):
                tag = n.get('p') or ('%s %s' % (n.get('who', ''), n.get('sc', ''))).strip()
                print('[%s %s %s | %s] %s' % (labels[k], n['d'], n['w'], tag, n['t'].replace('\n', ' / ')))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--raw-dir', required=True)
    ap.add_argument('--pulled', default=dt.date.today().isoformat())
    ap.add_argument('--nps', metavar='XLSX', help='NPS survey workbook; refreshes data/store_audit_nps.json first')
    ap.add_argument('--verbatim', action='store_true', help='include the auditor remarks word for word')
    ap.add_argument('--dump-notes', metavar='YYYY-MM', help='print that month\'s auditor comments instead of writing the data file')
    a = ap.parse_args()
    if a.nps:
        with open(NPS, 'w', encoding='utf-8') as f:
            json.dump(build_nps(a.nps), f, ensure_ascii=False, indent=1)
    data, report = build(a.raw_dir, dt.date.fromisoformat(a.pulled))
    if a.dump_notes:
        dump_notes(data, a.dump_notes)
        return
    if not a.verbatim:
        data['notes'] = {}
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('// Generated by scripts/build_store_audit.py from the Store Audit Google Sheet. Do not edit by hand.\n')
        f.write('window.AUDIT_DATA = ' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';\n')
    print('wrote', OUT, os.path.getsize(OUT), 'bytes')
    # Version the data file's URL so browsers holding last week's copy fetch the new one.
    with open(PAGE, encoding='utf-8') as f:
        page = f.read()
    page, n = re.subn(r'src="data/store_audit_data\.js(\?v=[^"]*)?"', 'src="data/store_audit_data.js?v=%s"' % a.pulled, page)
    if n != 1:
        sys.exit('STOP: data script tag not found in store-audit.html')
    with open(PAGE, 'w', encoding='utf-8', newline='') as f:
        f.write(page)
    print('months:', ', '.join('%s (%s)' % (m['key'], ' '.join(m['weeks'])) for m in data['months']))
    for k, v in sorted(report.items()):
        print('  %4d  %s' % (v, k))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
