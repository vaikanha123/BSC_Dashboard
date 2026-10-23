"""Plan the daily write of update_nps.py's --raw-out file to the "BSC NPS Raw Data" Google Sheet
(one tab per month, e.g. "Oct 2026"; row 1 headers, one row per WhatsApp NPS reply, oldest first).

The sheet is written with the Google Sheets connector, so every cell passes through the tool call.
To keep that small, only rows that changed since the last push are written: normally just the new
day's replies, plus any older row whose "Counted in NPS" flipped because the customer answered again.

  python scripts/nps_raw_push.py plan   <raw.json>   prints a JSON plan: tab, whether the tab is new,
                                                      the header row, and chunk files of {range, values}
  python scripts/nps_raw_push.py commit <raw.json>   after every chunk is written: records it as pushed

The pushed snapshot lives next to the raw file (git-ignored data/shopify_orders/, it holds PII).
"""
import json
import os
import sys

CHUNK = 60


def snapshot_path(raw):
    return os.path.join(os.path.dirname(os.path.abspath(raw)), 'nps_raw_pushed.json')


def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def cell(v):
    """The connector writes USER_ENTERED: keep '+91...' phones (and anything formula-like) as text."""
    if v is None:
        return ''
    if isinstance(v, str) and v[:1] in ('+', '=', '-', '@'):
        return "'" + v
    return v


def plan(raw):
    with open(raw, encoding='utf-8') as f:
        data = json.load(f)
    old = {}
    if os.path.exists(snapshot_path(raw)):
        with open(snapshot_path(raw), encoding='utf-8') as f:
            old = json.load(f)
    rows, width = data['rows'], len(data['header'])
    prev = old.get('rows', []) if old.get('tab') == data['tab'] else []
    first = next((i for i, r in enumerate(rows) if i >= len(prev) or prev[i] != r), len(rows))
    out_dir = os.path.dirname(os.path.abspath(raw))
    chunks = []
    for k, start in enumerate(range(first, len(rows), CHUNK)):
        part = rows[start:start + CHUNK]
        path = os.path.join(out_dir, 'nps_raw_chunk_%d.json' % (k + 1))
        rng = "'%s'!A%d:%s%d" % (data['tab'], start + 2, col(width), start + 1 + len(part))
        with open(path, 'w', encoding='utf-8') as f:
            json.dump({'range': rng, 'values': [[cell(v) for v in r] for r in part]}, f,
                      ensure_ascii=False)
        chunks.append({'file': path, 'range': rng, 'rows': len(part)})
    clear = None
    if len(prev) > len(rows):        # the month shrank (a reply removed from HO's sheet): blank the tail
        clear = "'%s'!A%d:%s%d" % (data['tab'], len(rows) + 2, col(width), len(prev) + 1)
    print(json.dumps({'tab': data['tab'], 'new_tab': old.get('tab') != data['tab'],
                      'header_range': "'%s'!A1:%s1" % (data['tab'], col(width + 1)),
                      'header': data['header'] + [data['stamp']], 'total_rows': len(rows),
                      'rows_to_write': len(rows) - first, 'clear': clear, 'chunks': chunks},
                     ensure_ascii=False, indent=1))


def commit(raw):
    with open(raw, encoding='utf-8') as f:
        data = json.load(f)
    with open(snapshot_path(raw), 'w', encoding='utf-8') as f:
        json.dump({'tab': data['tab'], 'rows': data['rows']}, f, ensure_ascii=False)
    print('OK. pushed snapshot: %s, %d rows' % (data['tab'], len(data['rows'])))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    if len(sys.argv) != 3 or sys.argv[1] not in ('plan', 'commit'):
        sys.exit(__doc__)
    (plan if sys.argv[1] == 'plan' else commit)(sys.argv[2])
