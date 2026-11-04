"""Plan the daily write of store-wise NPS to the "NPS & Quality Master sheet" Google Sheet
(19emM9CUzF5wZIPZD_TpSEehrlrrQAGR8DNanj8Ve-ms, tab "NPS master sheet": store names in column A, one
column per month headed "Oct 26", "Nov 26", ...; rows "Online" and "Grand Total" at the bottom).

  python scripts/nps_master_sheet.py <counts.json> <layout.json> --out <plan.json> [--today YYYY-MM-DD]

counts.json is update_nps.py --counts-out. layout.json is the connector's get_values result for
'NPS master sheet'!A1:AZ80 ({"range", "values"}), so rows and columns are found by name, not position.
Writes a JSON plan to --out (read it with the Read tool: the console mangles the en dash): {"writes": [{"month", "range", "values"}], "missing": [...], "extra": [...]}; each
write is one column (header + one cell per row) for update_values.

Which months: the current month, plus on the 1st-3rd the month before (late replies), but only from
November 2026 on -- earlier columns are HO's own numbers and are left as they are.
Cells: NPS = (promoters - detractors) / responses, whole %, written as "71%"; "–" with no responses.
Grand Total pools every row of the sheet that has counts (stores, Online, Virtual Stylist).
Store names are the POS locations, matched without regard to case ("Lakeshore Mall" = "LakeShore Mall").
"""
import argparse
import datetime as dt
import json
import sys

TAB = 'NPS master sheet'
FIRST_OWNED = '2026-11'
MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def label(mk):
    return '%s %s' % (MONTHS[int(mk[5:]) - 1], mk[2:4])


def prev_month(mk):
    y, m = int(mk[:4]), int(mk[5:])
    return '%d-%02d' % (y - (m == 1), 12 if m == 1 else m - 1)


def is_month(head, mk):
    """'Nov 26', 'Nov', 'November', 'Nov 2026' all name 2026-11; 'Nov 25' does not."""
    w = head.strip().split()
    if not w or not w[0].lower().startswith(MONTHS[int(mk[5:]) - 1].lower()):
        return False
    return len(w) == 1 or w[1] in (mk[2:4], mk[:4])


def find_column(header, mk):
    """0-based column for month mk: the one headed with it, else the column right after the month before."""
    for i in range(len(header) - 1, 0, -1):
        if header[i].strip() == label(mk):
            return i
    p = label(prev_month(mk))
    for i in range(len(header) - 1, 0, -1):
        if header[i].strip() == p:
            nxt = header[i + 1] if i + 1 < len(header) else ''
            if not nxt.strip() or is_month(nxt, mk):
                return i + 1
            sys.exit('STOP: the column after %r is headed %r, not %s' % (p, nxt, label(mk)))
    sys.exit('STOP: no %r or %r column in row 1 of %s' % (label(mk), p, TAB))


def pct(c):
    n = sum(c)
    return '%d%%' % round((c[0] - c[2]) / n * 100) if n else '–'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('counts')
    ap.add_argument('layout')
    ap.add_argument('--out', required=True)
    ap.add_argument('--today', default=dt.date.today().isoformat())
    a = ap.parse_args()
    with open(a.counts, encoding='utf-8') as f:
        counts = json.load(f)['months']
    with open(a.layout, encoding='utf-8') as f:
        grid = json.load(f)['values']
    header = grid[0] + [''] * 3
    names = [r[0].strip() if r else '' for r in grid]
    while names and not names[-1]:
        names.pop()

    mk = a.today[:7]
    months = [mk] + ([prev_month(mk)] if int(a.today[8:10]) <= 3 else [])
    months = [m for m in months if m >= FIRST_OWNED]
    writes, seen = [], set()
    for m in months:
        c = {k.lower(): v for k, v in counts.get(m, {}).items()}
        ci = find_column(header, m)
        vals, total = [label(m)], [0, 0, 0]
        for name in names[1:]:
            key = name.lower()
            if key == 'grand total':
                vals.append(None)                    # filled in below
            elif not key:
                vals.append(None)                    # blank row: leave alone
            else:
                v = c.get(key)
                vals.append(pct(v) if v else '–')
                if v:
                    seen.add(key)
                    total = [t + x for t, x in zip(total, v)]
        vals = [pct(total) if n.lower() == 'grand total' else v for n, v in zip(names, vals)]
        L = col(ci + 1)
        writes.append({'month': m, 'range': "'%s'!%s1:%s%d" % (TAB, L, L, len(names)),
                       'values': [[v] for v in vals], 'responses': sum(total), 'grand_total': pct(total)})
    sheet = {n.lower() for n in names[1:] if n}
    extra = sorted({k for m in months for k in counts.get(m, {}) if k.lower() not in sheet})
    missing = sorted(n for n in names[1:] if n and n.lower() not in seen and n.lower() != 'grand total')
    with open(a.out, 'w', encoding='utf-8') as f:
        json.dump({'writes': writes, 'missing': missing, 'extra': extra}, f, ensure_ascii=False)
    for w in writes:
        print('%s %s: %d responses, Grand Total %s' % (w['month'], w['range'], w['responses'], w['grand_total']))
    print('no responses: %s' % (', '.join(missing) or 'none'))
    print('not in the sheet: %s' % (', '.join(extra) or 'none'))


if __name__ == '__main__':
    main()
