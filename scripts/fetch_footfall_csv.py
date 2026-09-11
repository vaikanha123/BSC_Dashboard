"""
fetch_footfall_csv.py -- Pulls one month's tab of the footfall tracker Google Sheet (see the
bsc-footfall-tracker-sheet memory for the URL) and reconstructs it into the New FF/Repeat FF/
TOTAL FF block-labeled CSV format that generate_briefing_data.py's load_footfall() expects.

The sheet is link-viewable but its "View only" mode hides the File menu, so a literal
File > Download > CSV isn't available without edit access. Its gviz text-query export works
without auth, but does NOT preserve the literal "New FF"/"Repeat FF"/"TOTAL FF" label rows --
instead the three 30-ish-row date blocks (New FF, Repeat FF, then a per-weekday TOTAL FF block)
sit back-to-back with only blank-label summary rows between them. This script detects the block
boundaries by finding every row where the date column resets to day 1, strips the header's
per-store trailing whitespace (the raw header has e.g. "Jaipur Store " with a trailing space,
which won't match REGION_MAP's canonical "Jaipur Store"), and inserts the expected label rows.

Usage:
  python3 fetch_footfall_csv.py --gid 1472501457 --out "footfall_sep.csv"

  --gid : the sheet tab's gid (from its URL when that tab is selected in the browser).
  --out : where to write the reconstructed CSV.
  --sheet-id : optional, defaults to the BSC "FF Master File" (1IGzh7GR2GmsJqPAVu6sCHLshi9CG43sxP57CZKpBE20).
"""
import argparse
import csv
import io
import urllib.request

DEFAULT_SHEET_ID = '1IGzh7GR2GmsJqPAVu6sCHLshi9CG43sxP57CZKpBE20'


def fetch_gviz_csv(sheet_id, gid):
    url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&gid={gid}'
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode('utf-8')


def reconstruct(raw_csv_text):
    rows = list(csv.reader(io.StringIO(raw_csv_text)))
    header = rows[0]
    stores = [s.strip() for s in header[2:-1]]
    new_header = ['', 'Date'] + stores + ['Day Wise Total']

    block_starts = [i for i, r in enumerate(rows) if len(r) > 1 and r[1].strip().startswith('1-')]
    if len(block_starts) != 3:
        raise RuntimeError(
            f"Expected 3 date blocks (New FF, Repeat FF, TOTAL FF) but found {len(block_starts)} "
            f"at rows {block_starts} -- sheet layout may have changed, inspect manually before trusting this.")
    b1, b2, b3 = block_starts
    block_len = b2 - b1  # assume all three blocks are the same length (days in month)

    def block_rows(start):
        return rows[start:start + block_len]

    out_rows = [new_header]
    out_rows.append(['', 'New FF'])
    out_rows.extend(block_rows(b1))
    out_rows.append(['', 'Repeat FF'])
    out_rows.extend(block_rows(b2))
    out_rows.append(['', 'TOTAL FF'])
    out_rows.extend(block_rows(b3))
    return out_rows


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--gid', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--sheet-id', default=DEFAULT_SHEET_ID)
    args = ap.parse_args()

    raw = fetch_gviz_csv(args.sheet_id, args.gid)
    out_rows = reconstruct(raw)
    with open(args.out, 'w', encoding='utf-8', newline='') as f:
        csv.writer(f).writerows(out_rows)
    print(f"OK. Wrote {len(out_rows)} rows to {args.out} ({len(out_rows[0]) - 3} stores).")


if __name__ == '__main__':
    main()
