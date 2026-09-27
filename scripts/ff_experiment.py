"""ff_experiment.py -- does footfall improve the forecast? Walk-forward test of method G with vs without
footfall features (28d/91d and 7d/28d footfall trend, new-customer share of walk-ins), scored per month.
Read-only: touches nothing in data/ or the pages.

First run 2026-09-27 (Jan-Sep 2026): NO gain -- day 12.4% -> 12.8%, week 7.1% -> 7.5%, month-ahead
4.4% -> 4.6%, landing@14 3.6% -> 3.5%, and inconsistent between Jan-May and Jun-Sep -- so footfall is
collected daily (daily_refresh.py -> footfall.py fetch) but not used. Re-run once footfall covers a full
year incl. a festive season (~Jan 2027):
    python scripts/ff_experiment.py [--months 2026-01 2026-02 ...]
Ship only if it wins on both halves of the period (see runbook 6b).
"""
import argparse
import os
import sys
import time

import numpy as np
import pandas as pd

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'scripts'))
import forecast as F  # noqa: E402

cfg = F.load_cfg()
hist = F.load_history()
Y, last, reg = F.build_matrix(hist, cfg, extend_days=F.MAX_H)
d = F.Data(Y, last, reg, cfg)
act = F.actuals(d)
pos = {x: i for i, x in enumerate(d.dates)}

ff = pd.read_csv(os.path.join(REPO, 'data', 'footfall_history_daily.csv'), parse_dates=['date'])
T = ff.pivot_table(index='date', columns='store', values='total').reindex(index=d.dates, columns=d.stores)
N = ff.pivot_table(index='date', columns='store', values='new').reindex(index=d.dates, columns=d.stores)
T.loc[T.index > last] = np.nan
N.loc[N.index > last] = np.nan
t7 = T.rolling(7, min_periods=5).mean()
t28 = T.rolling(28, min_periods=21).mean()
t91 = T.rolling(91, min_periods=60).mean()
n28 = N.rolling(28, min_periods=21).sum() / T.rolling(28, min_periods=21).sum()
E = np.stack([(t28 / t91).values, (t7 / t28).values, n28.values], -1)  # dates x stores x features
E[~np.isfinite(E)] = np.nan  # HistGradientBoosting handles NaN (days/stores without footfall)

orig_samples = F.Data.samples
USE_FF = {'on': False}


def samples(self, oi, hs, mode):
    sm = orig_samples(self, oi, hs, mode)
    if USE_FF['on']:
        sm['X'] = np.column_stack([sm['X'], E[sm['o'], sm['s']]])
    return sm


F.Data.samples = samples


def run(month, use):
    USE_FF['on'] = use
    m = pd.Period(month, 'M')
    ci = pos[m.start_time - pd.Timedelta(days=1)]
    model = F.fit(d, ci)
    oi = np.arange(ci, min(pos[m.end_time.normalize()], d.last_i))
    rows = F.predict_rows(d, model, oi, 'forward')
    return rows[rows.target <= last]


def metrics(rows, month):
    m = pd.Period(month, 'M')
    start, end = m.start_time, m.end_time.normalize()
    o0 = start - pd.Timedelta(days=1)
    net = rows.groupby(['origin', 'target', 'h'])[['A', 'B', 'G']].sum().reset_index()
    net['act'] = net.target.map(act['network'])
    net['blend'] = net[['A', 'B', 'G']].astype(float).mean(axis=1)
    out = {}
    for col in ('G', 'blend'):
        day = net[net.h == 1]
        out[col + '_day'] = float((abs(day[col] - day.act) / day.act).mean())
        wk = net[net.h <= 7].groupby('origin').filter(lambda g: len(g) == 7).groupby('origin')[[col, 'act']].sum()
        out[col + '_week'] = float((abs(wk[col] - wk.act) / wk.act).mean())
        full = end <= last
        mo = net[(net.origin == o0) & (net.target <= end)]
        out[col + '_month'] = float(abs(mo[col].sum() / mo.act.sum() - 1)) if full else np.nan
        o14 = start + pd.Timedelta(days=13)
        rest = net[(net.origin == o14) & (net.target <= end)]
        mtd = act['network'][start:o14].sum()
        out[col + '_land14'] = float(abs((mtd + rest[col].sum()) / (mtd + rest.act.sum()) - 1)) if full and len(rest) else np.nan
    st = rows[rows.h <= 7].copy()
    st['act'] = d.Y.values[[pos[t] for t in st.target], [list(d.stores).index(s) for s in st.store]]
    sw = st.groupby(['origin', 'store'])[['G', 'act']].sum()
    sw = sw[sw.act > 0]
    out['G_store_week_wape'] = float(abs(sw.G - sw.act).sum() / sw.act.sum())
    return out


ap = argparse.ArgumentParser()
ap.add_argument('--months', nargs='+',
                default=[str(p) for p in pd.period_range('2026-01', pd.Period(last, 'M'), freq='M')])
MONTHS = ap.parse_args().months
res = []
for month in MONTHS:
    for use in (False, True):
        t0 = time.time()
        r = metrics(run(month, use), month)
        r.update(month=month, ff=use)
        res.append(r)
        print(month, 'FF  ' if use else 'base', {k: round(v, 4) for k, v in r.items() if isinstance(v, float)},
              '%.0fs' % (time.time() - t0), flush=True)
R = pd.DataFrame(res)
cols = [c for c in R.columns if c not in ('month', 'ff')]
print('all months (columns: ff False = without footfall, True = with)')
print(R.groupby('ff')[cols].mean().T.round(4))
half = len(MONTHS) // 2 + 1
for label, ms in (('first half', MONTHS[:half]), ('second half', MONTHS[half:])):
    if ms:
        print(label, ms[0], '..', ms[-1])
        print(R[R.month.isin(ms)].groupby('ff')[cols].mean().T.round(4))
