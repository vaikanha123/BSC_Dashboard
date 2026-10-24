"""update_sm_audit.py -- writes index.html's SM_AUDIT_BYMONTH (the SM Cohort tab's audit scores)
from the store audit report card's data (data/store_audit_data.js, built by build_store_audit.py).
Run by daily_refresh.py; rerun by hand after a fresh build_store_audit.py pull.

Usage (from the repo root):
  python scripts/update_sm_audit.py --html index.html

Every month in the audit data gets each SM Cohort store's compliance score, worked out exactly as
store-briefing.html and store-audit.html do (update_store_briefing.load_audit_scores: per
parameter, total score / total max over the month's weeks, then the mean over parameters), plus
the weeks it covers (e.g. 'Oct W1–W2'). The tab compares the month shown with August on this same
basis -- August is the cohort's fixed baseline.
"""
import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from update_store_briefing import AUDIT_DATA_PATH, load_audit_scores  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--html', required=True)
    a = ap.parse_args()

    with open(a.html, encoding='utf-8') as f:
        html = f.read()
    m = re.search(r'const SM_COHORT_STORES = (\[.*?\]);', html)
    if not m:
        sys.exit('STOP: SM_COHORT_STORES not found in ' + a.html)
    stores = json.loads(m.group(1))

    with open(AUDIT_DATA_PATH, encoding='utf-8') as f:
        js = f.read()
    months = [x['key'] for x in json.loads(js[js.index('{'):js.rindex('}') + 1])['months']]

    out = {}
    for month in months:
        _, audit = load_audit_scores(month)
        by_store = {s: audit[s]['current'] for s in stores if s in audit}
        if by_store:
            out[month] = {'label': next(iter(audit.values()))['currentLabel'], 'byStore': by_store}

    line = 'const SM_AUDIT_BYMONTH = ' + json.dumps(out, ensure_ascii=False) + ';'
    new, n = re.subn(r'^const SM_AUDIT_BYMONTH = [^\r\n]*;(?=\r?$)', lambda _: line, html, count=1, flags=re.M)
    if n != 1:
        sys.exit('STOP: const SM_AUDIT_BYMONTH line not found in ' + a.html)
    with open(a.html, 'w', encoding='utf-8', newline='') as f:
        f.write(new)
    print('OK. SM_AUDIT_BYMONTH: ' + ', '.join(f"{k} ({v['label']}, {len(v['byStore'])} stores)" for k, v in out.items()))


if __name__ == '__main__':
    main()
