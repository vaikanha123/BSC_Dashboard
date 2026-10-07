"""
build_sku_category_map.py — Rebuild data/sku_category_map.csv from the ENIGMA stock report.

The stock report's "Base" sheet carries merchandising's own hierarchy for every stock SKU
(Category -> Sub-Category 1). The Category tab of index.html groups sales by that hierarchy,
via bsc_common.load_sku_category_map(). Re-run this whenever a newer stock report is shared,
so newly launched SKUs get mapped:

    python scripts/build_sku_category_map.py --stock "ENIGMA STOCK REPORT dd.mm.yyyy ....xlsb"

Needs pyxlsb (pip install pyxlsb) for .xlsb files.
"""
import argparse
import csv
import re
import sys

OUT = 'data/sku_category_map.csv'


def clean_label(v, blank='Unspecified'):
    """Tidy a Base-sheet label: trim, treat 0/blank/#N/A as unspecified, prettify ALL_CAPS codes."""
    if v is None:
        return blank
    s = str(v).strip()
    if s in ('', '0', '0.0', '#N/A', '-') or s.startswith('0x'):
        return blank
    if s == s.upper() and re.search(r'[A-Z]', s):
        s = s.replace('_', ' ').title()   # MTM SHIRTS -> Mtm Shirts, METAL_CUFFLINKS -> Metal Cufflinks
        s = re.sub(r'\b(Mtm|Rtw|Bsc)\b', lambda m: m.group(1).upper(), s)
    return s


def read_base_rows(path):
    if path.lower().endswith('.xlsb'):
        import pyxlsb
        with pyxlsb.open_workbook(path) as wb:
            with wb.get_sheet('Base') as sh:
                for row in sh.rows():
                    yield [c.v for c in row]
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for row in wb['Base'].iter_rows(values_only=True):
            yield list(row)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--stock', required=True, help='ENIGMA stock report (.xlsb/.xlsx) with a "Base" sheet')
    ap.add_argument('--out', default=OUT)
    args = ap.parse_args()

    header, mapping, conflicts = None, {}, 0
    for row in read_base_rows(args.stock):
        if header is None:
            if 'sku_code' in row and 'Sub-Category 1' in row:
                header = {name: i for i, name in enumerate(row) if name is not None}
            continue
        sku = row[header['sku_code']]
        if sku is None or not str(sku).strip():
            continue
        sku = str(sku).strip().upper()
        pair = (clean_label(row[header['Category']]), clean_label(row[header['Sub-Category 1']]))
        if pair[0] == 'Unspecified':
            continue
        if sku in mapping and mapping[sku] != pair:
            conflicts += 1
            continue
        mapping[sku] = pair
    if header is None:
        print('ERROR: no header row with sku_code / Sub-Category 1 found in the Base sheet', file=sys.stderr)
        sys.exit(1)

    with open(args.out, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['sku', 'category', 'sub_category_1'])
        for sku in sorted(mapping):
            w.writerow([sku, *mapping[sku]])
    cats = sorted({c for c, _ in mapping.values()})
    print(f'OK. {len(mapping)} SKUs -> {args.out} ({len(cats)} categories: {", ".join(cats)})')
    if conflicts:
        print(f'NOTE: {conflicts} rows mapped an already-seen SKU to a different category; first mapping kept.')


if __name__ == '__main__':
    main()
