"""
bsc_common.py — Shared constants and helper functions for the BSC dashboard pipeline.

This module has no side effects on import (safe to import from either update script,
or from an interactive session for one-off analysis). See BSC_Dashboard_Runbook.md
for the full narrative explanation of what these pieces mean and why they exist.
"""
import json
import re
from collections import defaultdict

import pandas as pd


REGION_MAP = {
    'Ambience Vasant Kunj': 'North', 'DLF Midtown - Moti Nagar': 'North', 'Gurugram': 'North',
    'Khan Market': 'North', 'Mall of India, Noida': 'North', 'Select City': 'North', 'South Ex.': 'North',
    'Vegas Dwarka': 'North', 'Jaipur Store': 'North', 'Phoenix Palassio': 'North', 'Shakespearesarani': 'North',
    'Express Avenue': 'South', 'Indiranagar': 'South', 'Inorbit mall Hyderabad': 'South', 'Jayanagar': 'South',
    'Jubilee Hills': 'South', 'KNK Chennai': 'South', 'Kochi Store': 'South', 'LakeShore Mall': 'South',
    'Lavelle Road': 'South', 'Phoenix Marketcity, Whitefield': 'South', 'R.K. Salai': 'South',
    'Sarath City-Hyderabad': 'South',
    'Andheri': 'West', 'Inorbit Mall Malad West': 'West', 'Juhu Store': 'West',
    'Kalaghoda, Fort': 'West', 'Kemps Corner': 'West', 'Koregaon Park': 'West', 'Oberoi Mall Store': 'West',
    'Oberoi Sky City': 'West', 'Pali Hill, Bandra': 'West', 'Phoenix Marketcity Kurla': 'West',
    'Sindhu Bhavan Marg': 'West', 'Viviana Mall': 'West', 'PMC Viman Nagar Pune': 'West',
}
# Do NOT add "Online", "Corporate Orders", etc. to this dict for ad-hoc scripts -- a store not
# present here is correctly treated as non-offline. Adding a catch-all entry was a real bug before
# (a truthy-check on REGION_MAP.get(store) wrongly counted Online revenue as offline).

MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']


def categorize(product_type, sku):
    """MTM vs RTW shirt split, used for the Category tab's revenue/unit breakdown."""
    pt = str(product_type).strip() if pd.notna(product_type) else ''
    if pt.lower() == 'shirt':
        sku = str(sku) if pd.notna(sku) else ''
        return 'RTW Shirt' if '-' in sku else 'MTM Shirt'
    return pt if pt else 'Uncategorized'


def load_sales_csv(path):
    """Load a Shopify sales export, filtered to order rows, with the standard derived columns."""
    df = pd.read_csv(path, low_memory=False)
    df = df[df['Order or return'] == 'order'].copy()
    df['Day_str'] = df['Day'].str[:10]
    df['Revenue'] = df['Gross sales'].fillna(0) + df['Taxes'].fillna(0)  # established revenue basis -- do not change
    df['POS location name'] = df['POS location name'].fillna('Online').str.strip()
    df['Category'] = df.apply(lambda r: categorize(r['Product type'], r['Product variant SKU']), axis=1)
    df['Segment'] = df['New or returning customer'].fillna('').str.strip().str.lower()
    df['Stylist'] = df['Assisting staff member name'].fillna('').str.strip()
    df['StylistNorm'] = df['Stylist'].str.lower()
    df['Qty'] = df['Quantity ordered'].fillna(0)
    return df


def build_seed_days(df):
    """Build the SEED_DAYS dict (per-day store/category/stylist breakdown) from a loaded sales df."""
    seed_days = {}
    for dk, day_df in df.groupby('Day_str'):
        stores = day_df.groupby('POS location name')['Revenue'].sum().to_dict()
        bills = day_df.groupby('POS location name')['Order name'].nunique().to_dict()
        categories = day_df.groupby('Category')['Revenue'].sum().to_dict()
        categoryUnits = day_df.groupby('Category')['Qty'].sum().to_dict()
        storeUnits = day_df.groupby('POS location name')['Qty'].sum().to_dict()

        new_df = day_df[day_df['Segment'] == 'new']
        rep_df = day_df[day_df['Segment'] == 'returning']
        newRev = new_df.groupby('POS location name')['Revenue'].sum().to_dict()
        repRev = rep_df.groupby('POS location name')['Revenue'].sum().to_dict()
        newBills = new_df.groupby('POS location name')['Order name'].nunique().to_dict()
        repBills = rep_df.groupby('POS location name')['Order name'].nunique().to_dict()

        styl_df = day_df[day_df['Stylist'] != '']
        stylistRev = styl_df.groupby('Stylist')['Revenue'].sum().to_dict()
        stylistBills = styl_df.groupby('Stylist')['Order name'].nunique().to_dict()
        stylistUnits = styl_df.groupby('Stylist')['Qty'].sum().to_dict()

        stylistStoreRev, stylistStoreBills, stylistStoreUnits = {}, {}, {}
        for (styl, store), grp in styl_df.groupby(['Stylist', 'POS location name']):
            stylistStoreRev.setdefault(styl, {})[store] = float(grp['Revenue'].sum())
            stylistStoreBills.setdefault(styl, {})[store] = int(grp['Order name'].nunique())
            stylistStoreUnits.setdefault(styl, {})[store] = float(grp['Qty'].sum())

        seed_days[dk] = {
            'date': dk, 'stores': stores, 'bills': bills, 'categories': categories, 'categoryUnits': categoryUnits,
            'newRev': newRev, 'repRev': repRev, 'newBills': newBills, 'repBills': repBills,
            'stylistRev': stylistRev, 'stylistBills': stylistBills, 'storeUnits': storeUnits,
            'stylistStoreRev': stylistStoreRev, 'stylistUnits': stylistUnits,
            'stylistStoreBills': stylistStoreBills, 'stylistStoreUnits': stylistStoreUnits,
        }
    return seed_days


def month_key_from_date(v, current_month_fallback):
    """Parse a refund/CN 'Order Date'-style cell into a 'Mon-YYYY' key, falling back to the
    current month if the date can't be parsed (matches established handling of odd date formats)."""
    if v is None:
        return None
    if hasattr(v, 'year'):
        y = v.year
        return f"{MONTHS[v.month - 1]}-{y}" if 2020 <= y <= 2100 else None
    s = str(v).strip()
    m = re.match(r'(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s](\d{4})', s)
    if m:
        y = int(m.group(3))
        if 2020 <= y <= 2100:
            mon = m.group(2)[:3].lower()
            mon_names_l = [mn.lower() for mn in MONTHS]
            if mon in mon_names_l:
                return f"{MONTHS[mon_names_l.index(mon)]}-{y}"
    return None


def process_refund_sheet(ws, amount_idx, reason_idx, date_idx, current_month, cat_idx=6, channel_idx=0, store_idx=4):
    """Aggregate one refund or CN sheet (openpyxl worksheet) into the by-month structure baked
    into SEED_REFUNDS / SEED_CN. Column indices match the established Refund/CN export format --
    see the runbook section 3 if a new export has a different column order."""
    by_month = defaultdict(lambda: {
        'total': 0, 'onlineTotal': 0, 'storeTotal': 0, 'lineItems': 0, 'orders': set(),
        'byReason': defaultdict(float), 'byCategory': defaultdict(float), 'byStore': defaultdict(float),
        'byStoreCount': defaultdict(int),
        'byReasonOnline': defaultdict(float), 'byReasonOffline': defaultdict(float),
        'byCategoryOnline': defaultdict(float), 'byCategoryOffline': defaultdict(float),
    })
    for row in ws.iter_rows(min_row=2, values_only=True):
        amt = row[amount_idx]
        if not amt:
            continue
        mk = month_key_from_date(row[date_idx], current_month) or current_month
        channel = (row[channel_idx] or '').strip()
        is_online = channel.lower() == 'online'
        reason = (row[reason_idx] or 'Unspecified').strip() if row[reason_idx] else 'Unspecified'
        category = (row[cat_idx] or 'Uncategorized').strip() if row[cat_idx] else 'Uncategorized'
        store = (row[store_idx] or '').strip() if row[store_idx] else ''
        order_id = row[2] or ''
        m = by_month[mk]
        m['total'] += amt
        if is_online:
            m['onlineTotal'] += amt
        else:
            m['storeTotal'] += amt
        m['lineItems'] += 1
        if order_id:
            m['orders'].add(order_id)
        m['byReason'][reason] += amt
        m['byCategory'][category] += amt
        if store:
            m['byStore'][store] += amt
            m['byStoreCount'][store] += 1
        if is_online:
            m['byReasonOnline'][reason] += amt
            m['byCategoryOnline'][category] += amt
        else:
            m['byReasonOffline'][reason] += amt
            m['byCategoryOffline'][category] += amt
    out = {}
    for mk, m in by_month.items():
        out[mk] = {
            'total': m['total'], 'onlineTotal': m['onlineTotal'], 'storeTotal': m['storeTotal'],
            'lineItems': m['lineItems'], 'orders': len(m['orders']),
            'byReason': dict(m['byReason']), 'byCategory': dict(m['byCategory']), 'byStore': dict(m['byStore']),
            'byStoreCount': dict(m['byStoreCount']),
            'byReasonOnline': dict(m['byReasonOnline']), 'byReasonOffline': dict(m['byReasonOffline']),
            'byCategoryOnline': dict(m['byCategoryOnline']), 'byCategoryOffline': dict(m['byCategoryOffline']),
        }
    return out


def find_refund_and_cn_sheets(wb):
    """Detect which sheet is refund vs CN by header content, not sheet name (names vary a lot:
    refund/CN, refund/cn, CN/refund, etc.)."""
    refund_sheet, cn_sheet = None, None
    for sn in wb.sheetnames:
        ws = wb[sn]
        header = [str(c or '') for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
        joined = '|'.join(header).lower()
        if 'reason for refund' in joined:
            refund_sheet = sn
        elif 'reason for credit note' in joined or 'credit note amount' in joined:
            cn_sheet = sn
    return refund_sheet, cn_sheet


def replace_const(content, const_name, new_json_str):
    """Replace a `const NAME = {...}` or `const NAME = [...]` block in an HTML file's inline JS
    with new JSON content, using brace/bracket matching (not a naive string replace, since these
    blobs can be hundreds of KB and contain nested braces)."""
    start_marker = f"const {const_name} = "
    start_idx = content.index(start_marker)
    val_start = start_idx + len(start_marker)
    open_ch = '[' if new_json_str.lstrip().startswith('[') else '{'
    close_ch = ']' if open_ch == '[' else '}'
    i = val_start
    depth = 0
    started = False
    while i < len(content):
        if content[i] == open_ch:
            depth += 1
            started = True
        elif content[i] == close_ch:
            depth -= 1
            if started and depth == 0:
                break
        i += 1
    end_idx = i + 1
    # consume up to the trailing semicolon so we don't leave a stray one behind
    j = end_idx
    while content[j] != ';':
        j += 1
    return content[:start_idx] + f"const {const_name} = " + new_json_str + content[j:]


def extract_const(content, const_name, is_array=False):
    """Inverse of replace_const: pull out and json.loads() a `const NAME = ...` block."""
    marker = f"const {const_name} = "
    start = content.index(marker)
    val_start = start + len(marker)
    open_ch, close_ch = ('[', ']') if is_array else ('{', '}')
    i = val_start
    depth = 0
    started = False
    while i < len(content):
        if content[i] == open_ch:
            depth += 1
            started = True
        elif content[i] == close_ch:
            depth -= 1
            if started and depth == 0:
                break
        i += 1
    return json.loads(content[val_start:i + 1])


def syntax_check_html_js(path):
    """Extract the last <script> block from an HTML file and check it's valid JS via Node.
    Raises if invalid; call this after every edit, before pushing."""
    import re
    import subprocess
    import tempfile

    with open(path, encoding='utf-8') as f:
        content = f.read()
    scripts = re.findall(r'<script>(.*?)</script>', content, re.DOTALL)
    js = scripts[-1]
    with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as tmp:
        tmp.write(js)
        tmp_path = tmp.name
    result = subprocess.run(['node', '--check', tmp_path], capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"JS syntax check FAILED for {path}:\n{result.stderr}")
    return True


# ---- Established store-scoping overrides for individual stylists whose blended (all-store)
# totals are misleading because they've genuinely worked from different stores in different
# periods. Format: {stylist_name_lowercase: {month_key: store_name}}. Add new entries here if a
# similar situation is found for someone else -- don't build a one-off special case per person.
STORE_SCOPING_OVERRIDES = {
    'rahul yadav': {
        'jul': 'Andheri', 'aug': 'Andheri',
        'sep': 'Kemps Corner', 'oct': 'Kemps Corner', 'nov': 'Kemps Corner', 'dec': 'Kemps Corner',
        # extend forward month-by-month as needed; update if he moves again
    },
}


def scoped_stylist_metrics(df, stylist_name_lower, month_key):
    """Return (revenue, bills, units) for a stylist, applying any store-scoping override that
    applies to that month. df must already have Revenue/Qty/StylistNorm/POS location name columns
    (i.e. already passed through load_sales_csv)."""
    sub = df[df['StylistNorm'] == stylist_name_lower]
    override = STORE_SCOPING_OVERRIDES.get(stylist_name_lower, {})
    target_store = override.get(month_key)
    if target_store:
        sub = sub[sub['POS location name'] == target_store]
    rev = float(sub['Revenue'].sum())
    bills = int(sub['Order name'].nunique())
    units = float(sub['Qty'].sum())
    return rev, bills, units
