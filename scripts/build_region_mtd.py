"""build_region_mtd.py -- one region's "MTD <Month> Dashboard" table as a static page (+ optional PNG).

Rebuilds the store-wise MTD sheet the West RM used to keep by hand (LY / LM / target / MTD / projection /
bills / UPT / new + repeat AOV and revenue vs target) from data the main dashboard already has, so the
numbers always agree with index.html:
  - MTD revenue, bills, units, new/repeat split ... index.html SEED_DAYS (current month, 1st -> last day)
  - Month + MTD target ............................ index.html DAILY_TARGETS (day-wise targets)
  - LY (same month last year) ..................... index.html BASELINE.aug25StoreRevenue
  - LM (previous month) ........................... data/months/<YYYY-MM>.json (closed-month archive)
  - New/Repeat AOV targets, new-revenue share ..... data/briefing_targets.json
Final projection = MTD actual x days in month / days elapsed (straight-line, as in the original sheet);
Achv % = projection / month target. New/Repeat revenue targets = MTD target x new-revenue share.

Usage (repo root):
    python scripts/build_region_mtd.py --region West --out west-mtd.html [--png <path>]
--png renders a screenshot with headless Edge/Chrome (for the daily Slack post).
"""
import argparse
import calendar
import datetime as dt
import html
import json
import os
import subprocess
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bsc_common import extract_const  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Row order of the original West sheet; stores not listed here (new openings) go at the end.
STORE_ORDER = {
    'West': ['Andheri', 'Inorbit Mall Malad West', 'Juhu Store', 'Kalaghoda, Fort', 'Kemps Corner',
             'Koregaon Park', 'Oberoi Mall Store', 'Oberoi Sky City', 'Pali Hill, Bandra',
             'Phoenix Marketcity Kurla', 'Sindhu Bhavan Marg', 'Viviana Mall', 'PMC Viman Nagar Pune'],
}
BROWSERS = [r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
            r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
            r'C:\Program Files\Google\Chrome\Application\chrome.exe']


def build_rows(region):
    content = open(os.path.join(ROOT, 'index.html'), encoding='utf-8').read()
    region_map = extract_const(content, 'REGION_MAP')
    seed_days = extract_const(content, 'SEED_DAYS')
    daily_targets = extract_const(content, 'DAILY_TARGETS')
    baseline_start = content.index('const BASELINE = {')
    ly_rev = extract_const(content[baseline_start:].replace('aug25StoreRevenue: ', 'const LY = ', 1), 'LY')

    dates = sorted(seed_days)
    last = dt.date.fromisoformat(dates[-1])
    month_key = last.strftime('%Y-%m')
    days_in_month = calendar.monthrange(last.year, last.month)[1]
    elapsed = last.day
    prev = (last.replace(day=1) - dt.timedelta(days=1)).strftime('%Y-%m')

    lm_rev = {}
    lm_path = os.path.join(ROOT, 'data', 'months', prev + '.json')
    if os.path.exists(lm_path):
        for d in json.load(open(lm_path, encoding='utf-8'))['days'].values():
            for s, v in d['stores'].items():
                lm_rev[s] = lm_rev.get(s, 0) + v
    aov_t = json.load(open(os.path.join(ROOT, 'data', 'briefing_targets.json'), encoding='utf-8')).get(month_key, {})

    stores = [s for s, r in region_map.items() if r == region]
    order = STORE_ORDER.get(region, [])
    stores.sort(key=lambda s: (order.index(s) if s in order else len(order), s))

    rows = []
    for s in stores:
        agg = {k: 0.0 for k in ('rev', 'newRev', 'repRev', 'newBills', 'repBills', 'bills', 'units')}
        for d in dates:
            day = seed_days[d]
            agg['rev'] += day['stores'].get(s, 0)
            agg['newRev'] += day.get('newRev', {}).get(s, 0)
            agg['repRev'] += day.get('repRev', {}).get(s, 0)
            agg['newBills'] += day.get('newBills', {}).get(s, 0)
            agg['repBills'] += day.get('repBills', {}).get(s, 0)
            agg['bills'] += day.get('bills', {}).get(s, 0)
            agg['units'] += day.get('storeUnits', {}).get(s, 0)
        tgt = daily_targets.get(s, {})
        month_tgt = sum(v for k, v in tgt.items() if k.startswith(month_key))
        mtd_tgt = sum(v for k, v in tgt.items() if k.startswith(month_key) and k <= dates[-1])
        t = aov_t.get(s, {})
        share = t.get('newRevShare')
        rows.append({
            'store': s, 'ly': ly_rev.get(s), 'lm': lm_rev.get(s), 'target': month_tgt or None,
            'mtdTarget': mtd_tgt or None, 'mtd': agg['rev'], 'proj': agg['rev'] * days_in_month / elapsed,
            'newBills': agg['newBills'], 'repBills': agg['repBills'],
            'bills': agg['newBills'] + agg['repBills'], 'units': agg['units'],
            'newAovT': t.get('newAov'), 'repAovT': t.get('repAov'),
            'newRev': agg['newRev'], 'repRev': agg['repRev'],
            'newRevT': mtd_tgt * share if share is not None and mtd_tgt else None,
            'repRevT': mtd_tgt * (1 - share) if share is not None and mtd_tgt else None,
        })
    meta = {'last': last, 'elapsed': elapsed, 'days': days_in_month,
            'lyLabel': last.replace(year=last.year - 1).strftime("%b '%y"),
            'lmLabel': dt.date.fromisoformat(prev + '-01').strftime("%b '%y"),
            'month': last.strftime('%B')}
    return rows, meta


def total_row(rows):
    def tot(k):
        vals = [r[k] for r in rows if r[k] is not None]
        return sum(vals) if vals else None
    t = {k: tot(k) for k in ('ly', 'lm', 'target', 'mtdTarget', 'mtd', 'proj', 'newBills', 'repBills', 'bills',
                             'units', 'newRev', 'repRev', 'newRevT', 'repRevT')}
    t['store'] = 'Total'
    # Region AOV target = revenue target / implied bills target (each store's revenue target / its AOV target).
    for seg in ('new', 'rep'):
        pairs = [(r[seg + 'RevT'], r[seg + 'AovT']) for r in rows if r[seg + 'RevT'] and r[seg + 'AovT']]
        t[seg + 'AovT'] = sum(p[0] for p in pairs) / sum(p[0] / p[1] for p in pairs) if pairs else None
    return t


def fmt_int(v):
    if v is None:
        return '–'
    n = int(round(v))
    s = str(abs(n))
    if len(s) > 3:  # Indian grouping: 12,34,567
        head, tail = s[:-3], s[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        s = ','.join(parts) + ',' + tail
    return ('-' if n < 0 else '') + s


def pct_cell(num, den):
    if not num or not den:
        return '<td class="num">–</td>'
    p = num / den * 100
    cls = 'good' if p >= 100 else ('warn' if p >= 90 else 'bad')
    return '<td class="num pct %s">%d%%</td>' % (cls, round(p))


def row_html(r, is_total=False):
    new_aov = r['newRev'] / r['newBills'] if r['newBills'] else None
    rep_aov = r['repRev'] / r['repBills'] if r['repBills'] else None
    upt = r['units'] / r['bills'] if r['bills'] else None
    c = []
    c.append('<td class="store">%s</td>' % html.escape(r['store']))
    for k in ('ly', 'lm', 'target', 'mtdTarget', 'mtd'):
        c.append('<td class="num">%s</td>' % fmt_int(r[k]))
    c.append('<td class="num proj">%s</td>' % fmt_int(r['proj']))
    c.append(pct_cell(r['proj'], r['target']))
    for k in ('newBills', 'repBills', 'bills', 'units'):
        c.append('<td class="num">%s</td>' % fmt_int(r[k]))
    c.append('<td class="num">%s</td>' % ('%.1f' % upt if upt else '–'))
    c.append('<td class="num">%s</td>' % fmt_int(r['newAovT']))
    c.append('<td class="num">%s</td>' % fmt_int(new_aov))
    c.append(pct_cell(new_aov, r['newAovT']))
    c.append('<td class="num">%s</td>' % fmt_int(r['repAovT']))
    c.append('<td class="num">%s</td>' % fmt_int(rep_aov))
    c.append(pct_cell(rep_aov, r['repAovT']))
    c.append('<td class="num">%s</td>' % fmt_int(r['newRevT']))
    c.append('<td class="num">%s</td>' % fmt_int(r['newRev']))
    c.append(pct_cell(r['newRev'], r['newRevT']))
    c.append('<td class="num">%s</td>' % fmt_int(r['repRevT']))
    c.append('<td class="num">%s</td>' % fmt_int(r['repRev']))
    c.append(pct_cell(r['repRev'], r['repRevT']))
    return '<tr%s>%s</tr>' % (' class="total"' if is_total else '', ''.join(c))


def render(region, rows, meta):
    last = meta['last']
    through = '%d %s %d' % (last.day, last.strftime('%b'), last.year)
    head = ('<tr class="h1"><th rowspan="2" class="store">Store</th><th rowspan="2">LY<br>%s</th>'
            '<th rowspan="2">LM<br>%s</th><th rowspan="2">Month<br>Target</th><th rowspan="2">MTD<br>Target</th>'
            '<th rowspan="2">MTD<br>Actual</th><th rowspan="2">Final<br>Proj</th><th rowspan="2">Achv<br>%%</th>'
            '<th colspan="5" class="grp">Bills &amp; units</th><th colspan="3" class="grp">New cust AOV</th>'
            '<th colspan="3" class="grp">Repeat cust AOV</th><th colspan="3" class="grp">New cust revenue</th>'
            '<th colspan="3" class="grp">Repeat cust revenue</th></tr>'
            '<tr class="h2"><th>New</th><th>Rep</th><th>Total</th><th>Qty</th><th>UPT</th>'
            + '<th>Target</th><th>Actual</th><th>Achv</th>' * 4 + '</tr>') % (meta['lyLabel'], meta['lmLabel'])
    body = ''.join(row_html(r) for r in rows) + row_html(total_row(rows), True)
    return '''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>%(region)s MTD Dashboard</title>
<style>
:root{--bg:#f6f7f5;--card:#fff;--ink:#1d2320;--muted:#5f6b66;--line:#d9ded9;--head:#e6efe7;--head2:#f1f5f1;
--good:#d4edda;--good-ink:#1e6b33;--warn:#fff3cd;--warn-ink:#7a5a00;--bad:#f8d7da;--bad-ink:#9b1c2a;--total:#fff6bf}
@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){--bg:#141816;--card:#1c211f;--ink:#e8ece9;
--muted:#9aa6a0;--line:#2f3633;--head:#24302a;--head2:#202824;--good:#1f3d29;--good-ink:#8fdba5;--warn:#3d3416;
--warn-ink:#f0d080;--bad:#45202a;--bad-ink:#f2a3ad;--total:#3a3517}}
:root[data-theme="dark"]{--bg:#141816;--card:#1c211f;--ink:#e8ece9;--muted:#9aa6a0;--line:#2f3633;--head:#24302a;
--head2:#202824;--good:#1f3d29;--good-ink:#8fdba5;--warn:#3d3416;--warn-ink:#f0d080;--bad:#45202a;--bad-ink:#f2a3ad;--total:#3a3517}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);
font:13px/1.35 "Segoe UI",system-ui,-apple-system,sans-serif;padding:16px}
.wrap{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px;width:max-content;max-width:100%%}
h1{font-size:18px;margin:0 0 2px}.sub{color:var(--muted);margin:0 0 10px}
.scroll{overflow-x:auto}table{border-collapse:collapse;font-variant-numeric:tabular-nums;white-space:nowrap}
th,td{border:1px solid var(--line);padding:4px 7px}th{background:var(--head);font-weight:600;font-size:12px;text-align:center}
tr.h2 th{background:var(--head2)}th.grp{border-bottom-color:var(--line)}
td.num{text-align:right}td.store,th.store{text-align:left;position:sticky;left:0;background:var(--card);z-index:1}
th.store{background:var(--head)}td.proj{font-weight:600}
td.pct{font-weight:600}td.good{background:var(--good);color:var(--good-ink)}td.warn{background:var(--warn);color:var(--warn-ink)}
td.bad{background:var(--bad);color:var(--bad-ink)}
tr.total td{background:var(--total);font-weight:700}tr.total td.pct.good,tr.total td.pct.warn,tr.total td.pct.bad{border-top:2px solid var(--ink)}
tr.total td{border-top:2px solid var(--ink)}
.note{color:var(--muted);font-size:11.5px;margin:10px 0 0;max-width:1100px}
</style></head><body><div class="wrap">
<h1>%(region)s &middot; MTD %(month)s Dashboard</h1>
<p class="sub">Sales through %(through)s &middot; day %(elapsed)d of %(days)d &middot; revenue = gross sales + taxes + shipping</p>
<div class="scroll"><table><thead>%(head)s</thead><tbody>%(body)s</tbody></table></div>
<p class="note">Final Proj = MTD actual &times; %(days)d / %(elapsed)d (straight-line). Achv %% = Final Proj vs month target.
AOV columns compare actual AOV with the store&rsquo;s HO AOV target; revenue targets are the MTD target split by the
store&rsquo;s planned new-customer share. Green = at or above 100%%, amber = 90&ndash;99%%, red = below 90%%.
LY and LM are full-month totals. Same data as the main BSC dashboard, refreshed daily around 10 AM.</p>
</div></body></html>
''' % dict(region=html.escape(region), month=meta['month'], through=through, elapsed=meta['elapsed'],
           days=meta['days'], head=head, body=body)


def screenshot(page, png, rows_count):
    browser = next((b for b in BROWSERS if os.path.exists(b)), None)
    if not browser:
        raise SystemExit('STOP: no Edge/Chrome found for --png')
    height = 190 + 27 * (rows_count + 1)
    url = 'file:///' + os.path.abspath(page).replace(os.sep, '/')
    png = os.path.abspath(png)
    if os.path.exists(png):
        os.remove(png)
    # Own throwaway profile, so a running Edge/Chrome window doesn't swallow the headless call. The launcher
    # can return before the headless child has written the file, hence the wait.
    with tempfile.TemporaryDirectory() as prof:
        subprocess.run([browser, '--headless=new', '--disable-gpu', '--hide-scrollbars', '--force-device-scale-factor=2',
                        '--user-data-dir=' + prof, '--window-size=1880,%d' % height, '--screenshot=' + png, url],
                       capture_output=True, timeout=120)
        for _ in range(60):
            if os.path.exists(png) and os.path.getsize(png) > 0:
                time.sleep(2)
                break
            time.sleep(1)
    if not os.path.exists(png):
        raise SystemExit('STOP: screenshot not written')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--region', default='West')
    ap.add_argument('--out', default='west-mtd.html')
    ap.add_argument('--png')
    a = ap.parse_args()
    rows, meta = build_rows(a.region)
    out = os.path.join(ROOT, a.out) if not os.path.isabs(a.out) else a.out
    with open(out, 'w', encoding='utf-8', newline='\n') as f:
        f.write(render(a.region, rows, meta))
    print('OK: %s written (%d stores, through %s)' % (a.out, len(rows), meta['last']))
    if a.png:
        screenshot(out, a.png, len(rows))
        print('OK: screenshot %s' % a.png)


if __name__ == '__main__':
    main()
