"""Store-wise NPS by month from the WhatsApp survey Google Sheet ("bombay_shirt_company_whatsapp_data",
1U0-PZ33Tq2IICmv0_QCGUDHjrq5bLJG1qUMXVKC8cZU, tab Sheet1). The sheet is refreshed every morning by
HO with one row per button reply; it is not link-shared, so it is pulled with the Google Sheets
connector (get_values 'Sheet1!A1:M5000'), whose oversized result is saved as a JSON file:

  python scripts/update_nps.py --sheet-json <saved get_values result> --sales <MTD sales csv>

The sheet has no store or order. A response is put against the POS location of the customer's
order, matched by email, using the same rule as the NPS_Fit_Assessment workbook (Notes tab): the
customer's latest order placed at least 7 days before the response, else their latest order on or
before the response date. On the September responses whose true order is known, that rule picks
the right store about 93% of the time. When a customer answers more than once for the same order,
only the latest answer counts.

Order history is a git-ignored email -> orders file (CUSTOMER_ORDERS below, it holds customer
emails) kept in the main checkout. Each run adds the --sales file to it; --seed adds older sales
reports (the 2025 report only has a Month column, so those orders are dated to the 1st).

Writes:
  data/store_audit_nps.json   promoter / neutral / detractor counts per store and month
                              (audit-tracker store names); read by store-audit.html (through
                              data/store_audit_data.js, whose "nps" key is rewritten here too)
                              and by update_store_briefing.py
  data/store_audit_fit.json   the same for the fit assessment campaign (~7sf91: Very satisfied /
                              Neutral / Not satisfied), the "fit" key of store_audit_data.js
  index.html SEED_NPS         the latest month's NPS per POS location, "Online", per-store counts, overall
                              (store responses pooled, matching store-audit.html) and online apart
  --raw-out <json>            (optional) one row per NPS reply in the dashboard's live month, with
                              the matched order's store, stylist and products, for the "BSC NPS Raw
                              Data" Google Sheet. It holds customer emails and phones: write it outside
                              the repo or under the git-ignored data/shopify_orders/, never commit it.
                              Products and stylists come from a second git-ignored cache,
                              ORDER_DETAILS, fed by --sales / --seed like the order history.

September 2026 is kept as it is in store_audit_nps.json: it was built from the September survey
files, which carry the true order of most responses, so it is better than an email match.
"""
import argparse
import collections
import datetime as dt
import json
import os
import re
import subprocess
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bsc_common import REGION_MAP, extract_const, replace_const, syntax_check_html_js  # noqa: E402
from build_store_audit import NPS_STORES, MONTH_NAMES, PAGE  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NPS_JSON = os.path.join(ROOT, 'data', 'store_audit_nps.json')
FIT_JSON = os.path.join(ROOT, 'data', 'store_audit_fit.json')
AUDIT_JS = os.path.join(ROOT, 'data', 'store_audit_data.js')
NPS_CAMPAIGN = '~3fp3nn'
SLOT = {'Very likely': 0, 'Neutral': 1, 'Not likely': 2}
FIT_CAMPAIGN = '~7sf91'           # fit assessment, same sheet; same store match and repeat rule as NPS
FIT_SLOT = {'Very satisfied': 0, 'Neutral': 1, 'Not satisfied': 2}
FIRST_MONTH = '2026-09'          # August starts on the 26th; it is left off, as before
FROZEN_MONTHS = {'2026-09'}      # see the docstring
MIN_GAP_DAYS = 7
SHEET_HEAD = ['user_id', 'campaign_id', 'event_date', 'event_name', 'button_label', 'button_type',
              'response_time', 'email', 'first_name', 'last_name', 'phone']


def main_checkout():
    """The order history sits in the main checkout's git-ignored data/, also when run from a worktree."""
    try:
        common = subprocess.run(['git', 'rev-parse', '--path-format=absolute', '--git-common-dir'], cwd=ROOT,
                                capture_output=True, text=True, check=True).stdout.strip()
        return os.path.dirname(common)
    except (OSError, subprocess.CalledProcessError):
        return ROOT


CUSTOMER_ORDERS = os.path.join(main_checkout(), 'data', 'shopify_orders', 'customer_orders.csv.gz')
ORDER_DETAILS = os.path.join(main_checkout(), 'data', 'shopify_orders', 'order_details.csv.gz')
NPS_GROUP = ['Promoter', 'Passive', 'Detractor']
RAW_HEAD = ['Response date', 'Response time (IST)', 'Customer name', 'Email', 'Phone', 'Answer', 'NPS group',
            'Counted in NPS', 'Matched order', 'Order date', 'Days order to reply', 'Store', 'Region',
            'Sales channel', 'Stylist', 'Products', 'Product types', 'Units', 'Order value (total sales)']


def read_orders(path):
    """Distinct orders (email, order, day, pos) from a Shopify "Sales Data Report" CSV."""
    head = pd.read_csv(path, nrows=0).columns
    date_col = 'Day' if 'Day' in head else 'Month'
    df = pd.read_csv(path, usecols=[date_col, 'Order name', 'Order or return', 'Customer email', 'POS location name'],
                     dtype=str)
    df = df[(df['Order or return'] == 'order') & df['Customer email'].notna()]
    out = pd.DataFrame({'email': df['Customer email'].str.strip().str.lower(), 'order': df['Order name'],
                        'day': df[date_col].str[:10], 'pos': df['POS location name'].fillna('')})
    return out.drop_duplicates('order')


def load_order_history(sales, seeds):
    frames = []
    if os.path.exists(CUSTOMER_ORDERS):
        frames.append(pd.read_csv(CUSTOMER_ORDERS, dtype=str, keep_default_na=False))
    for p in list(seeds) + ([sales] if sales else []):
        frames.append(read_orders(p))
    if not frames:
        sys.exit('STOP: no order history at %s -- pass --seed <sales report csv> once' % CUSTOMER_ORDERS)
    hist = pd.concat(frames, ignore_index=True)
    # A later file wins for the same order (e.g. a Day-dated row over a Month-dated one).
    hist = hist.drop_duplicates('order', keep='last').sort_values(['email', 'day', 'order'])
    os.makedirs(os.path.dirname(CUSTOMER_ORDERS), exist_ok=True)
    hist.to_csv(CUSTOMER_ORDERS, index=False, compression='gzip')
    by_email = collections.defaultdict(list)
    for e, o, d, p in hist[['email', 'order', 'day', 'pos']].itertuples(index=False):
        by_email[e].append((d, o, p))
    print('order history: %d orders, %d customers, %s -> %s' % (len(hist), len(by_email), hist['day'].min(),
                                                              hist['day'].max()))
    return by_email


def read_details(path):
    """Per order: channel, stylist(s), product lines, units and total sales, from a "with Stylist Name" report."""
    head = pd.read_csv(path, nrows=0).columns
    if 'Assisting staff member name' not in head:
        return None
    df = pd.read_csv(path, dtype=str, usecols=['Order name', 'Order or return', 'Line type', 'Sales channel',
                                               'Product title at time of sale', 'Product variant title at time of sale',
                                               'Product type', 'Assisting staff member name', 'Quantity ordered',
                                               'Total sales'])
    df = df[df['Order or return'] == 'order']
    df = df.assign(qty=pd.to_numeric(df['Quantity ordered'], errors='coerce').fillna(0),
                   total=pd.to_numeric(df['Total sales'], errors='coerce').fillna(0))
    out = []
    for order, g in df.groupby('Order name', sort=False):
        lines = g[(g['Line type'] == 'product') & g['Product title at time of sale'].notna()]
        prods = {}
        for t, v, q in lines[['Product title at time of sale', 'Product variant title at time of sale', 'qty']].itertuples(index=False):
            label = t + (' (%s)' % v if isinstance(v, str) and v else '')
            prods[label] = prods.get(label, 0) + int(q)
        uniq = lambda col: ', '.join(dict.fromkeys(x for x in g[col].dropna() if x))
        out.append({'order': order, 'channel': uniq('Sales channel'), 'stylist': uniq('Assisting staff member name'),
                    'products': '; '.join('%s x%d' % (k, q) if q > 1 else k for k, q in prods.items()),
                    'types': uniq('Product type'), 'units': int(lines['qty'].sum()), 'value': round(g['total'].sum(), 2)})
    return pd.DataFrame(out)


def load_order_details(sales, seeds):
    frames = []
    if os.path.exists(ORDER_DETAILS):
        frames.append(pd.read_csv(ORDER_DETAILS, dtype=str, keep_default_na=False))
    for p in list(seeds) + ([sales] if sales else []):
        d = read_details(p)
        if d is not None:
            frames.append(d.astype(str))
    if not frames:
        return {}
    det = pd.concat(frames, ignore_index=True).drop_duplicates('order', keep='last')
    det.to_csv(ORDER_DETAILS, index=False, compression='gzip')
    return {r['order']: r for r in det.to_dict('records')}


def match(orders, resp_day):
    """The customer's latest order placed >= 7 days before the response, else latest on/before it."""
    cutoff = (dt.date.fromisoformat(resp_day) - dt.timedelta(days=MIN_GAP_DAYS)).isoformat()
    early = [x for x in orders if x[0] <= cutoff]
    if early:
        return early[-1]
    same = [x for x in orders if x[0] <= resp_day]
    return same[-1] if same else None


def load_sheet(path):
    with open(path, encoding='utf-8') as f:
        values = json.load(f)['values']
    head = [h.strip() for h in values[0][:len(SHEET_HEAD)]]
    if head != SHEET_HEAD:
        sys.exit('STOP: survey sheet columns changed: %s' % head)
    stamp = next((c for c in values[0][len(SHEET_HEAD):] if 'refreshed' in c.lower()), '')
    rows = [r + [''] * (len(SHEET_HEAD) - len(r)) for r in values[1:] if r and len(r) > 6]
    return rows, stamp


def build(rows, by_email, campaign=NPS_CAMPAIGN, slots=SLOT):
    """Counts [promoters, neutral, detractors] per (month, POS location or channel), latest answer per order."""
    latest = {}
    report = collections.Counter()
    for r in rows:
        if r[1] != campaign or r[4] not in slots:
            continue
        day, email = r[2][:10], r[7].strip().lower()
        if day[:7] < FIRST_MONTH:
            continue
        m = match(by_email.get(email, []), day)
        pos = m[2] if m else None
        key = (email, m[1] if m else None)
        if key in latest:
            report['repeat answers for the same order (older ones dropped)'] += 1
            if latest[key][0] >= r[6]:
                continue
        latest[key] = (r[6], day, pos, slots[r[4]])
    counts = collections.defaultdict(lambda: collections.defaultdict(lambda: [0, 0, 0]))
    dates = collections.defaultdict(list)
    for _, day, pos, slot in latest.values():
        mk = day[:7]
        if pos is None:
            where = 'Not matched'
        elif pos in REGION_MAP:
            where = pos
        elif pos == '':
            where = 'Online'
        else:
            where = pos              # Virtual Stylist, Corporate Orders, Head Office
        counts[mk][where][slot] += 1
        dates[mk].append(day)
        report['responses ' + ('at a store' if pos in REGION_MAP else 'not at a store: ' + where)] += 1
    return counts, dates, report


def raw_rows(rows, by_email, details, mk, campaign=NPS_CAMPAIGN, slots=SLOT):
    """Every NPS reply of month mk with its matched order, store, stylist and products. "Counted in NPS"
    is No for an older answer to the same order (build() keeps only the latest)."""
    picked, latest = [], {}
    for i, r in enumerate(rows):
        if r[1] != campaign or r[4] not in slots or r[2][:7] != mk:
            continue
        day, email = r[2][:10], r[7].strip().lower()
        m = match(by_email.get(email, []), day)
        key = (email, m[1] if m else None)
        if key not in latest or latest[key][0] < r[6]:     # as build(): a same-second repeat keeps the first
            latest[key] = (r[6], i)
        picked.append((i, r, day, email, m, key))
    table = []
    for i, r, day, email, m, key in sorted(picked, key=lambda x: x[1][6]):
        try:
            ist = (dt.datetime.strptime(r[6][:19], '%Y-%m-%d %H:%M:%S')
                   + dt.timedelta(hours=5, minutes=30)).strftime('%Y-%m-%d %H:%M')
        except ValueError:
            ist = r[6]
        d = details.get(m[1], {}) if m else {}
        store = 'Not matched' if m is None else (m[2] or 'Online')
        table.append([day, ist, ('%s %s' % (r[8], r[9])).strip(), email, r[10], r[4], NPS_GROUP[slots[r[4]]],
                      'Yes' if latest[key][1] == i else 'No (older answer, same order)',
                      m[1] if m else '', m[0] if m else '',
                      (dt.date.fromisoformat(day) - dt.date.fromisoformat(m[0])).days if m else '',
                      store, REGION_MAP.get(m[2], '') if m else '', d.get('channel', ''), d.get('stylist', ''),
                      d.get('products', ''), d.get('types', ''), int(d['units']) if d.get('units') else '',
                      float(d['value']) if d.get('value') else ''])
    return table


def nps_detail(table, slots=SLOT):
    """The counted store replies of one month's raw_rows() table, per dashboard store: [slot, date, order,
    phone, products]. store-audit.html lists them when a promoter / passive / detractor count is clicked."""
    out = collections.defaultdict(list)
    for t in table:
        if t[7] == 'Yes' and t[11] in REGION_MAP:
            out[NPS_STORES.get(t[11], t[11])].append([slots[t[5]], t[0], t[8], t[4], t[15]])
    return out


def write_raw(path, rows, by_email, details, html, stamp):
    """The dashboard's live month (as SEED_NPS), so on the 1st the closing month's tab gets its last day."""
    with open(html, encoding='utf-8') as f:
        mk = max(extract_const(f.read(), 'SEED_DAYS'))[:7]
    table = raw_rows(rows, by_email, details, mk)
    data = {'month': mk, 'tab': '%s %s' % (MONTH_NAMES[int(mk[5:]) - 1][:3], mk[:4]), 'stamp': stamp,
            'header': RAW_HEAD, 'rows': table}
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)
    print('raw NPS rows for %s: %d (%d counted in NPS, %d not matched to an order, %d matched without product '
          'detail) -> %s' % (data['tab'], len(table), sum(1 for t in table if t[7] == 'Yes'),
                             sum(1 for t in table if not t[8]), sum(1 for t in table if t[8] and not t[15]), path))


def month_meta(mk, days):
    lo, hi = dt.date.fromisoformat(min(days)), dt.date.fromisoformat(max(days))
    return {'key': mk, 'label': MONTH_NAMES[lo.month - 1][:3],
            'range': '%d–%d %s' % (lo.day, hi.day, MONTH_NAMES[hi.month - 1][:3])}


def score(c):
    n = sum(c)
    return round((c[0] - c[2]) / n * 100, 1) if n else None


def write_audit_nps(counts, dates, source, detail=None, path=NPS_JSON, key='nps'):
    """Also writes the fit assessment (path=FIT_JSON, key='fit'), in the same shape."""
    old = {'months': [], 'stores': {}}
    if os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            old = json.load(f)
    months = [m for m in old['months'] if m['key'] in FROZEN_MONTHS]
    stores = {s: {mk: c for mk, c in v.items() if mk in FROZEN_MONTHS} for s, v in old['stores'].items()}
    # Per-reply detail (order, phone, products): frozen months keep theirs, the rest come from this run.
    keep = {s: {mk: r for mk, r in v.items() if mk in FROZEN_MONTHS} for s, v in old.get('detail', {}).items()}
    for mk, by_store in (detail or {}).items():
        if mk not in FROZEN_MONTHS:
            for s, r in by_store.items():
                keep.setdefault(s, {})[mk] = r
    for mk in sorted(counts):
        if mk in FROZEN_MONTHS:
            continue
        months.append(month_meta(mk, dates[mk]))
        for pos, c in counts[mk].items():
            if pos in REGION_MAP:
                stores.setdefault(NPS_STORES.get(pos, pos), {})[mk] = c
    months.sort(key=lambda m: m['key'])
    stores = {s: v for s, v in stores.items() if v}
    data = {'source': source, 'months': months, 'stores': stores, 'detail': {s: v for s, v in keep.items() if v}}
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    # store-audit.html reads the copy embedded in store_audit_data.js; swap just that key.
    with open(AUDIT_JS, encoding='utf-8') as f:
        js = f.read()
    m = re.search(r'^window\.AUDIT_DATA = (.*);\s*$', js, re.MULTILINE | re.DOTALL)
    if not m:
        sys.exit('STOP: window.AUDIT_DATA not found in %s' % AUDIT_JS)
    audit = json.loads(m.group(1))
    audit[key] = data
    with open(AUDIT_JS, 'w', encoding='utf-8', newline='') as f:
        f.write(js[:m.start(1)] + json.dumps(audit, ensure_ascii=False, separators=(',', ':')) + ';\n')
    # Version the data file's URL (as build_store_audit.py does) so browsers drop yesterday's cached copy.
    with open(PAGE, encoding='utf-8') as f:
        page = f.read()
    page, n = re.subn(r'src="data/store_audit_data\.js(\?v=[^"]*)?"',
                      'src="data/store_audit_data.js?v=%s"' % dt.date.today().isoformat(), page)
    if n != 1:
        sys.exit('STOP: data script tag not found in store-audit.html')
    with open(PAGE, 'w', encoding='utf-8', newline='') as f:
        f.write(page)
    return data


def write_seed_nps(html, counts, dates):
    """SEED_NPS follows the dashboard's live month (newest SEED_DAYS month), not the newest survey month,
    so on the 1st -- before the month transition -- it keeps the closing month, and the archive
    snapshot (month_archive.py) picks up that month's NPS."""
    with open(html, encoding='utf-8', newline='') as f:
        content = f.read()
    mk = max(extract_const(content, 'SEED_DAYS'))[:7]
    if mk not in counts:
        print('NOTE: no survey responses yet for the live month %s -- SEED_NPS left as it is' % mk)
        return None
    c = counts[mk]
    stores = {pos: list(v) for pos, v in sorted(c.items()) if pos in REGION_MAP}
    by_store = {pos: score(v) for pos, v in stores.items()}
    if c.get('Online'):
        by_store['Online'] = score(c['Online'])
    # overall = store responses pooled, as on store-audit.html (online, corporate and virtual stylist kept apart)
    total = [sum(v[i] for v in stores.values()) for i in range(3)]
    other = [sum(v[i] for pos, v in c.items() if pos not in stores) for i in range(3)]
    y, mo = int(mk[:4]), int(mk[5:])
    seed = {'month': '%s-%d' % (MONTH_NAMES[mo - 1][:3], y), 'range': month_meta(mk, dates[mk])['range'],
            'responses': sum(total), 'byStore': by_store, 'counts': stores, 'overall': score(total),
            'online': {'nps': score(other), 'responses': sum(other)}}
    new = replace_const(content, 'SEED_NPS', json.dumps(seed, ensure_ascii=False))
    with open(html, 'w', encoding='utf-8', newline='') as f:
        f.write(new)
    try:
        syntax_check_html_js(html)
    except Exception:
        with open(html, 'w', encoding='utf-8', newline='') as f:
            f.write(content)    # leave the page as it was
        raise
    return seed


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--sheet-json', required=True, help="saved Sheets-connector get_values result for 'Sheet1!A1:M5000'")
    ap.add_argument('--sales', help='month-to-date sales CSV, added to the order history')
    ap.add_argument('--seed', nargs='*', default=[], help='older sales report CSVs to add to the order history')
    ap.add_argument('--html', default=os.path.join(ROOT, 'index.html'))
    ap.add_argument('--raw-out', help="write the live month's per-reply raw rows (has PII) to this JSON file")
    a = ap.parse_args()

    rows, stamp = load_sheet(a.sheet_json)
    print('survey sheet: %d rows. %s' % (len(rows), stamp))
    by_email = load_order_history(a.sales, a.seed)
    counts, dates, report = build(rows, by_email)
    if not counts:
        sys.exit('STOP: no NPS responses from %s on in the sheet' % FIRST_MONTH)
    details = load_order_details(a.sales, a.seed)
    detail = {mk: nps_detail(raw_rows(rows, by_email, details, mk)) for mk in counts if mk not in FROZEN_MONTHS}
    source = 'WhatsApp survey sheet (%s)' % (stamp or 'refresh time not shown')
    data = write_audit_nps(counts, dates, source, detail)
    fcounts, fdates, freport = build(rows, by_email, FIT_CAMPAIGN, FIT_SLOT)
    fdetail = {mk: nps_detail(raw_rows(rows, by_email, details, mk, FIT_CAMPAIGN, FIT_SLOT), FIT_SLOT)
               for mk in fcounts if mk not in FROZEN_MONTHS}
    fit = write_audit_nps(fcounts, fdates, source, fdetail, FIT_JSON, 'fit')
    seed = write_seed_nps(a.html, counts, dates)
    if a.raw_out:
        write_raw(a.raw_out, rows, by_email, details, a.html, stamp)
    for k, v in sorted(report.items()):
        print('  %5d  %s' % (v, k))
    print('store_audit_nps.json months: %s' % ', '.join('%s (%s)' % (m['key'], m['range']) for m in data['months']))
    print('store_audit_fit.json months: %s; %s' % (', '.join('%s (%s)' % (m['key'], m['range']) for m in fit['months']),
          ', '.join('%s %d' % (mk, sum(sum(c) for p, c in v.items() if p in REGION_MAP)) for mk, v in sorted(fcounts.items()))))
    if seed:
        print('OK. SEED_NPS %s (%s): %d responses, overall NPS %s, %d stores'
              % (seed['month'], seed['range'], seed['responses'], seed['overall'],
                 sum(1 for s in seed['byStore'] if s in REGION_MAP)))
    else:
        print('OK. store_audit_nps.json updated; SEED_NPS unchanged')


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
