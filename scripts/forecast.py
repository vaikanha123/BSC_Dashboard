"""forecast.py -- offline-store revenue forecasting ("learning layer") for the separate, unlinked forecast page
(forecast-1ad8b79c.html -- deliberately not linked from index.html).

Usage (from the repo root):
    python scripts/forecast.py init-history --sales "<multi-year sales CSV>"
        One-time: aggregate a full Shopify "Sales Data Report with Stylist Name" export into
        data/sales_history_daily.csv (date x POS location: revenue, bills, units -- all locations,
        Online included, so an online forecast can be added later without re-exporting).
    python scripts/forecast.py backtest [--months 18]
        Walk-forward backtest: for each of the last N months, train only on data before it and
        forecast every day of that month 1..MAX_H days ahead. Seeds data/forecast_log.csv with
        these "as-if" forecasts (source=backtest), scores the three blend policies, and records the
        winner in data/forecast_backtest.json. Re-run after changing the model or features -- a change
        only goes live if it wins here (champion/challenger).
    python scripts/forecast.py run --sales "<month-to-date CSV>" --html forecast-1ad8b79c.html --dashboard index.html
        At a month transition pass the complete previous-month file first, since the daily task skips
        the 1st/2nd and the month's last day would otherwise never reach the history:
            --sales "<full previous month CSV>" --sales "<new month-to-date CSV>"
        Daily (called by daily_refresh.py): merge the MTD file into the history, retrain, forecast the
        next MAX_H days, learn blend weights/bias from the scored log, append today's forecast to the
        log (source=live), and write `const FORECAST = {...}` into index.html.

How it learns:
  * Three forecasters per store per day, summed to region/network:
      A  mean of the store's last 4 same-weekday days
      B  A x how the network moved over the same weeks last year
      G  gradient boosting on target/level ratio (weekday, month, festival distance, store age,
         store's weekday profile, 3-month momentum, last-year seasonal factor, horizon)
  * Every forecast is logged (data/forecast_log.csv) and scored once actuals arrive.
  * Blend weights per horizon bucket are re-learned every run from the last 120 days of scored
    forecasts (inverse error), and an optional small bias correction -- but only in the policy the
    backtest chose (equal / adaptive / adaptive+bias).
  * Forecast ranges (80%) come from the logged error distribution for the same horizon window.
Revenue basis and store lists come from bsc_common (Gross sales + Taxes; REGION_MAP = offline).
Closed stores and the festival calendar live in data/forecast_config.json.
"""
import argparse
import datetime as dt
import json
import os
import sys

import numpy as np
import pandas as pd
from sklearn.ensemble import HistGradientBoostingRegressor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bsc_common import REGION_MAP, replace_const  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data')
HIST = os.path.join(DATA, 'sales_history_daily.csv')
LOG = os.path.join(DATA, 'forecast_log.csv')
CFG = os.path.join(DATA, 'forecast_config.json')
BT = os.path.join(DATA, 'forecast_backtest.json')

MAX_H = 62            # covers the rest of this month + all of next month from any origin
MIN_AGE = 28          # days of history before a store is modelled (younger stores use a run-rate fallback)
GAP_DAYS = 14         # a run of >= this many zero-sales days inside a store's life = temporary closure
POLICY_LOOKBACK = 120
BUCKETS = [('day', 1, 1), ('week', 2, 7), ('month', 8, MAX_H)]
SCOPES = ['network', 'North', 'South', 'West']
LOG_COLS = ['source', 'origin', 'target', 'h', 'scope', 'A', 'B', 'G', 'final']


def log(msg):
    print(msg, flush=True)


# ---------------------------------------------------------------- history

def aggregate_sales(path):
    cols = ['Day', 'Order name', 'Order or return', 'POS location name', 'Gross sales', 'Taxes', 'Quantity ordered']
    df = pd.read_csv(path, usecols=cols, low_memory=False)
    df = df[df['Order or return'] == 'order']
    df['date'] = df['Day'].str[:10]
    df['rev'] = df['Gross sales'].fillna(0) + df['Taxes'].fillna(0)  # established revenue basis
    df['loc'] = df['POS location name'].fillna('Online').str.strip()
    g = df.groupby(['date', 'loc']).agg(rev=('rev', 'sum'), bills=('Order name', 'nunique'),
                                        units=('Quantity ordered', 'sum')).reset_index()
    g['rev'] = g['rev'].round(2)
    return g


def load_history():
    h = pd.read_csv(HIST)
    h['date'] = pd.to_datetime(h['date'])
    return h


def merge_into_history(sales_csv):
    new = aggregate_sales(sales_csv)
    lo, hi = new['date'].min(), new['date'].max()
    hist = pd.read_csv(HIST)
    if hist['date'].max() < lo and (pd.Timestamp(lo) - pd.Timestamp(hist['date'].max())).days > 1:
        raise SystemExit('STOP: history ends %s but the new file starts %s -- gap; run init-history with a '
                         'file covering the gap' % (hist['date'].max(), lo))
    hist = hist[(hist['date'] < lo) | (hist['date'] > hi)]
    hist = pd.concat([hist, new]).sort_values(['date', 'loc'])
    hist.to_csv(HIST, index=False)
    log('history: replaced %s..%s from the MTD file; history now %s..%s' % (lo, hi, hist['date'].min(), hist['date'].max()))


# ---------------------------------------------------------------- matrices / features

def load_cfg():
    with open(CFG, encoding='utf-8') as f:
        return json.load(f)


def build_matrix(hist, cfg, extend_days=0):
    """Store x day revenue matrix. NaN = store not trading (before opening, after closing, or in a
    >= GAP_DAYS zero-sales gap such as a renovation). Future days (extend_days) are 0 placeholders
    for stores still trading, NaN for everything else."""
    reg = {**REGION_MAP, **cfg['closed_stores']}
    g = hist[hist['loc'].isin(reg)]
    piv = g.pivot_table(index='date', columns='loc', values='rev', aggfunc='sum')
    last = g['date'].max()
    idx = pd.date_range(piv.index.min(), last + pd.Timedelta(days=extend_days))
    piv = piv.reindex(idx)
    Y = pd.DataFrame(np.nan, index=idx, columns=piv.columns)
    for c in piv:
        s = piv[c].dropna()
        s = s[s > 0]
        if s.empty:
            continue
        start, end = s.index.min(), s.index.max()
        still_trading = end >= last - pd.Timedelta(days=7)
        if still_trading:
            end = last  # a quiet last day or two is a zero-sales day, not a closure
        seg = piv.loc[start:end, c].fillna(0)
        z = (seg <= 0).astype(int)
        run_id = (z.diff() != 0).cumsum()
        runlen = z.groupby(run_id).transform('sum')
        seg[(z == 1) & (runlen >= GAP_DAYS)] = np.nan
        Y.loc[start:end, c] = seg
        if still_trading and extend_days:
            Y.loc[last + pd.Timedelta(days=1):, c] = 0.0
    return Y, last, reg


def fest_matrix(dates, festivals, windows):
    cols = []
    for name, ds in festivals.items():
        ds = pd.to_datetime(ds).values
        diff = (dates.values[:, None] - ds[None, :]).astype('timedelta64[D]').astype(int)
        nearest = diff[np.arange(len(dates)), np.abs(diff).argmin(1)]
        lo, hi = windows.get(name, (-10, 5))
        cols.append(np.where((nearest >= lo) & (nearest <= hi), nearest, 99))
    return np.stack(cols, 1)


class Data:
    def __init__(self, Y, last, reg, cfg):
        self.Y, self.last, self.cfg = Y, last, cfg
        self.dates, self.stores = Y.index, Y.columns
        self.last_i = int(np.where(self.dates == last)[0][0])
        self.Ya = Y.values
        obs = Y.loc[:last]
        self.L = obs.rolling(28, min_periods=14).mean().ffill().reindex(self.dates).ffill().values
        self.L91 = obs.rolling(91, min_periods=28).mean().ffill().reindex(self.dates).ffill().values
        self.R = obs.rolling(28, min_periods=1).mean().reindex(self.dates).ffill().values
        D = np.full((7,) + self.Ya.shape, np.nan)
        for k in range(7):
            sub = obs[obs.index.dayofweek == k].rolling(4, min_periods=2).mean()
            D[k] = sub.reindex(self.dates).ffill().values
        self.D = D
        self.first = np.array([Y[c].first_valid_index() for c in Y]).astype('datetime64[D]')
        self.didx = self.dates.values.astype('datetime64[D]')
        self.dow = self.dates.dayofweek.values
        self.mon = self.dates.month.values
        self.dleft = (self.dates.days_in_month - self.dates.day).values
        wins = {k: tuple(v) for k, v in cfg.get('festival_windows', {}).items()}
        self.F = fest_matrix(self.dates, cfg['festivals'], wins)
        self.region_names = ['North', 'South', 'West']
        self.reg = np.array([self.region_names.index(reg[c]) for c in Y])
        per = obs.sum(axis=1) / obs.notna().sum(axis=1)
        self.P7 = per.rolling(7, center=True, min_periods=4).mean().reindex(self.dates).values
        self.P28 = per.rolling(28).mean().reindex(self.dates).values
        base = per.rolling(91, min_periods=28).mean()
        dowf = np.ones((len(self.dates), 7))
        for k in range(7):
            pk = per.where(per.index.dayofweek == k).rolling(91, min_periods=7).mean()
            dowf[:, k] = (pk / base).reindex(self.dates).ffill().fillna(1.0).values
        self.dowf = dowf

    def samples(self, oi, hs, mode):
        """mode: 'train' (model rows with actuals), 'eval' (all rows with actuals), 'forward'."""
        o, h, s = np.meshgrid(np.asarray(oi), np.asarray(hs), np.arange(len(self.stores)), indexing='ij')
        o, h, s = o.ravel(), h.ravel(), s.ravel()
        t = o + h
        keep = t < len(self.dates)
        o, h, s, t = o[keep], h[keep], s[keep], t[keep]
        age_o = (self.didx[o] - self.first[s]).astype(int)
        ok = ~np.isnan(self.Ya[o, s]) & ~np.isnan(self.Ya[t, s]) & (age_o >= 0)
        if mode != 'forward':
            ok &= t <= self.last_i
        o, h, s, t, age_o = o[ok], h[ok], s[ok], t[ok], age_o[ok]
        lev = self.L[o, s]
        model_row = (age_o >= MIN_AGE) & ~np.isnan(lev) & (lev > 0)
        if mode == 'train':
            o, h, s, t, age_o, lev, model_row = [a[model_row] for a in (o, h, s, t, age_o, lev, model_row)]
        lev = np.where(model_row, lev, self.R[o, s])
        dA = self.D[self.dow[t], o, s]
        fbA = lev * self.dowf[o, self.dow[t]]
        dA = np.where(model_row & ~np.isnan(dA), dA, fbA)
        ly_ok = (t - 364 >= 0) & (o - 364 >= 27)
        tb, ob = np.where(ly_ok, t - 364, 0), np.where(ly_ok, o - 364, 0)
        lyf = np.where(ly_ok, self.P7[tb] / self.P28[ob], np.nan)
        age_t = np.minimum((self.didx[t] - self.first[s]).astype(int), 730)
        l91 = np.where(np.isnan(self.L91[o, s]), lev, self.L91[o, s])
        safe = np.where(lev > 0, lev, 1.0)
        X = np.column_stack([self.dow[t], self.mon[t], self.dleft[t], h, age_t, dA / safe, np.log(safe),
                             self.reg[s], self.F[t], np.log(np.maximum(l91, 1) / safe), lyf])
        B = np.where(np.isnan(lyf), dA, dA * lyf)
        y = self.Ya[t, s] if mode != 'forward' else None
        return dict(o=o, h=h, s=s, t=t, lev=lev, X=X, y=y, A=dA, B=B, model_row=model_row)


def fit(d, cutoff_i):
    sm = d.samples(np.arange(60, cutoff_i, 3), np.arange(1, MAX_H + 1), 'train')
    m = sm['t'] <= cutoff_i
    ratio = np.clip(sm['y'][m] / sm['lev'][m], 0, 8)
    model = HistGradientBoostingRegressor(loss='poisson', learning_rate=0.05, max_iter=500, max_leaf_nodes=31,
                                          min_samples_leaf=3000, categorical_features=[7], random_state=0)
    model.fit(sm['X'][m], ratio, sample_weight=sm['lev'][m])
    return model


def predict_rows(d, model, oi, mode):
    sm = d.samples(oi, np.arange(1, MAX_H + 1), mode)
    G = model.predict(sm['X']) * sm['lev']
    G = np.where(sm['model_row'], G, sm['A'])
    B = np.where(sm['model_row'], sm['B'], sm['A'])
    return pd.DataFrame({'origin': d.dates[sm['o']], 'target': d.dates[sm['t']], 'h': sm['h'],
                         'store': d.stores[sm['s']], 'region': np.array(d.region_names)[d.reg[sm['s']]],
                         'A': sm['A'], 'B': B, 'G': G, 'fallback': ~sm['model_row']})


def aggregate(rows):
    net = rows.groupby(['origin', 'target', 'h'])[['A', 'B', 'G']].sum().reset_index().assign(scope='network')
    reg = rows.groupby(['origin', 'target', 'h', 'region'])[['A', 'B', 'G']].sum().reset_index()
    reg = reg.rename(columns={'region': 'scope'})
    return pd.concat([net, reg], ignore_index=True)


def actuals(d):
    """Actual offline revenue per date per scope, from the matrix (includes stores not forecast)."""
    obs = d.Y.loc[:d.last]
    out = {'network': obs.sum(axis=1, min_count=1)}
    regs = np.array(d.region_names)[d.reg]
    for r in d.region_names:
        out[r] = obs.loc[:, regs == r].sum(axis=1, min_count=1)
    return out


# ---------------------------------------------------------------- policy (the "learning" part)

def bucket_of(h):
    h = np.asarray(h)
    return np.select([h == 1, h <= 7], ['day', 'week'], 'month')


def scored(logdf, act, as_of, lookback=POLICY_LOOKBACK):
    """Network rows whose target has an actual and whose origin is within `lookback` days before as_of."""
    x = logdf[(logdf.scope == 'network') & (logdf.target <= as_of) &
              (logdf.origin > as_of - pd.Timedelta(days=lookback))].copy()
    x['actual'] = x['target'].map(act['network'])
    return x.dropna(subset=['actual'])


def learn_policy(logdf, act, as_of, mode):
    pol = {'mode': mode, 'weights': {}, 'bias': {}, 'n': {}}
    x = scored(logdf, act, as_of) if mode != 'equal' and len(logdf) else pd.DataFrame()
    for b, _, _ in BUCKETS:
        w = {'A': 1 / 3, 'B': 1 / 3, 'G': 1 / 3}
        bias = 1.0
        xb = x[bucket_of(x['h']) == b] if len(x) else x
        n = int(xb['origin'].nunique()) if len(xb) else 0
        if mode != 'equal' and n >= 20:
            err = {c: (np.abs(xb[c] - xb['actual']) / xb['actual']).mean() for c in 'ABG'}
            inv = {c: 1 / max(e, 1e-3) for c, e in err.items()}
            tot = sum(inv.values())
            w = {c: inv[c] / tot for c in 'ABG'}
            if mode == 'adaptive_bias':
                recent = xb[xb['origin'] > as_of - pd.Timedelta(days=28)]
                if recent['origin'].nunique() >= 14:
                    blend = sum(w[c] * recent[c] for c in 'ABG')
                    r = float(np.median(recent['actual'] / blend))
                    if abs(r - 1) > 0.02:
                        bias = float(np.clip(1 + 0.5 * (r - 1), 0.95, 1.05))
        pol['weights'][b] = {c: round(v, 4) for c, v in w.items()}
        pol['bias'][b] = round(bias, 4)
        pol['n'][b] = n
    return pol


def apply_policy(df, pol):
    b = bucket_of(df['h'])
    out = np.zeros(len(df))
    for name, _, _ in BUCKETS:
        m = b == name
        w = pol['weights'][name]
        out[m] = (w['A'] * df['A'].values[m] + w['B'] * df['B'].values[m] + w['G'] * df['G'].values[m]) * pol['bias'][name]
    return out


# ---------------------------------------------------------------- log I/O and error bands

def read_log():
    if not os.path.exists(LOG):
        return pd.DataFrame(columns=LOG_COLS)
    x = pd.read_csv(LOG, parse_dates=['origin', 'target'])
    return x


def write_log(x):
    x = x.sort_values(['origin', 'scope', 'h'])
    x.to_csv(LOG, index=False, float_format='%.0f', date_format='%Y-%m-%d')


def error_matrix(logdf, act, as_of):
    """origin x h matrices of network final forecast and actual (last 365 days of origins)."""
    x = logdf[(logdf.scope == 'network') & (logdf.origin > as_of - pd.Timedelta(days=365))].copy()
    x['actual'] = x['target'].map(act['network'])
    x.loc[x['target'] > as_of, 'actual'] = np.nan
    F = x.pivot_table(index='origin', columns='h', values='final', aggfunc='last').reindex(columns=range(1, MAX_H + 1))
    Ac = x.pivot_table(index='origin', columns='h', values='actual', aggfunc='last').reindex(columns=range(1, MAX_H + 1))
    Ac = Ac.reindex(F.index)
    return F.values, Ac.values


def band(F, Ac, h0, h1, q=(0.1, 0.9)):
    """Relative error quantiles of window sums over horizons h0..h1 (inclusive)."""
    f = F[:, h0 - 1:h1]
    a = Ac[:, h0 - 1:h1]
    full = ~np.isnan(f).any(1) & ~np.isnan(a).any(1)
    if full.sum() < 15:
        return None
    rel = a[full].sum(1) / f[full].sum(1) - 1
    return [float(np.quantile(rel, q[0])), float(np.quantile(rel, q[1])), int(full.sum())]


# ---------------------------------------------------------------- month landing (MTD actual + rest of month)
# Backtest findings (2026-09-26, 513 origins): the newest forecast over-reacts to a strong/weak start of the
# month -- first-half surprises mostly don't persist -- and method B is noisy over a part-month. Averaging the
# forecast made just before the month began with the newest one (A and G) helps clearly in days 8-21 (e.g.
# days 8-14: 4.63% -> 3.87%), but not early (the two forecasts are nearly the same) nor late in the month,
# when a start-of-month forecast that has already proven wrong only drags the landing (Sep-2026: the 31-Aug
# forecast was 14% high for Sep 1-25 and still pulled the day-25 landing up ~Rs10L).
# So the choice is made per phase of the month. Every run re-scores each candidate per phase on the last 365
# days of logged forecasts; a challenger replaces a phase's default only if it is better by
# LANDING_SWITCH_REL (relative), so the choice keeps learning without flip-flopping.

LANDING_METHODS = {
    'latest': 'newest forecast only',
    'avg_start_latest': 'average of the start-of-month forecast and the newest forecast',
    'avg_start_latest_noB': 'average of the start-of-month forecast and the newest forecast (methods A and G)',
}
LANDING_PHASES = {'early': (1, 7), 'mid': (8, 21), 'late': (22, 31)}  # day of month the data runs through
LANDING_DEFAULTS = {'early': 'latest', 'mid': 'avg_start_latest_noB', 'late': 'latest'}
LANDING_SWITCH_REL = 0.05


def landing_phase(day):
    return next(ph for ph, (a, b) in LANDING_PHASES.items() if a <= day <= b)


def start_origin(origins, month_start):
    """Latest forecast origin in the 7 days before the month began (the 'start-of-month' forecast)."""
    c = [o for o in origins if month_start - pd.Timedelta(days=7) <= o < month_start]
    return max(c) if c else None


def landing_rest(method, new, old_final):
    """Rest-of-month forecast. new: (final, A, G) sums from the newest origin; old_final: sum from the
    start-of-month origin (None if there wasn't one)."""
    final, A, G = new
    if method == 'latest' or old_final is None:
        return final
    newest = final if method == 'avg_start_latest' else (A + G) / 2
    return 0.5 * newest + 0.5 * old_final


def score_landing(lg, act, last, lookback=365):
    """Mean |landing miss| of each method per phase of the month, over logged origins in the last `lookback`
    days whose month has closed. Returns {phase: ({method: miss}, n)}."""
    x = lg[(lg.scope == 'network') & (lg.origin > last - pd.Timedelta(days=lookback))]
    out = {ph: ({}, 0) for ph in LANDING_PHASES}
    if x.empty:
        return out
    P = {c: x.pivot_table(index='origin', columns='target', values=c, aggfunc='last') for c in ('final', 'A', 'G')}
    origins = list(P['final'].index)
    errs = {ph: {m: [] for m in LANDING_METHODS} for ph in LANDING_PHASES}
    net = act['network']
    for o in origins:
        if o.day == o.days_in_month:
            continue  # month-ahead, not a landing
        per = pd.Period(o, 'M')
        ms, me = per.start_time, per.end_time.normalize()
        if me > last:
            continue
        rest = [t for t in pd.date_range(o + pd.Timedelta(days=1), me) if t in P['final'].columns]
        if len(rest) != (me - o).days:
            continue
        s0 = start_origin(origins, ms)
        if s0 is None or P['final'].loc[s0, rest].isna().any():
            continue
        new = tuple(float(P[c].loc[o, rest].sum()) for c in ('final', 'A', 'G'))
        old = float(P['final'].loc[s0, rest].sum())
        mtd, total = float(net[ms:o].sum()), float(net[ms:me].sum())
        ph = landing_phase(o.day)
        for m in LANDING_METHODS:
            errs[ph][m].append(abs((mtd + landing_rest(m, new, old)) / total - 1))
    for ph, e in errs.items():
        out[ph] = ({m: float(np.mean(v)) for m, v in e.items() if v}, len(e[LANDING_DEFAULTS[ph]]))
    return out


def choose_landing(phase_scores):
    """{phase: method}: each phase keeps its default unless another method beat it by LANDING_SWITCH_REL
    over at least 20 scored origins in that phase."""
    choice = {}
    for ph, (scores, n) in phase_scores.items():
        d = LANDING_DEFAULTS[ph]
        if n < 20 or d not in scores:
            choice[ph] = d
            continue
        best = min(scores, key=scores.get)
        choice[ph] = best if scores[best] < scores[d] * (1 - LANDING_SWITCH_REL) else d
    return choice


def landing14_by_month(lg, act, method, months):
    """Signed day-14 landing miss (forecast/actual - 1) per month with the given method, from the log."""
    x = lg[lg.scope == 'network']
    P = {c: x.pivot_table(index='origin', columns='target', values=c, aggfunc='last') for c in ('final', 'A', 'G')}
    origins = list(P['final'].index)
    out = {}
    for m in months:
        per = pd.Period(m, 'M')
        ms, me = per.start_time, per.end_time.normalize()
        o = ms + pd.Timedelta(days=13)
        if o not in P['final'].index or me > act['network'].index.max():
            continue
        rest = list(pd.date_range(o + pd.Timedelta(days=1), me))
        s0 = start_origin(origins, ms)
        new = tuple(float(P[c].loc[o, rest].sum()) for c in ('final', 'A', 'G'))
        old = float(P['final'].loc[s0, rest].sum()) if s0 is not None else None
        mtd, total = float(act['network'][ms:o].sum()), float(act['network'][ms:me].sum())
        out[m] = round((mtd + landing_rest(method, new, old)) / total - 1, 4)
    return out


# ---------------------------------------------------------------- backtest

def month_errors(df, act):
    df = df.copy()
    df['actual'] = df['target'].map(act['network'])
    return df


def cmd_backtest(a):
    cfg = load_cfg()
    hist = load_history()
    Y, last, reg = build_matrix(hist, cfg, extend_days=MAX_H)
    d = Data(Y, last, reg, cfg)
    act = actuals(d)
    pos = {x: i for i, x in enumerate(d.dates)}
    last_month = pd.Period(last, 'M')
    months = pd.period_range(last_month - a.months, last_month, freq='M')
    seeded = []
    for m in months:
        ci = pos[m.start_time - pd.Timedelta(days=1)]
        if ci >= d.last_i:
            break
        model = fit(d, ci)
        oi = np.arange(ci, min(pos[m.end_time.normalize()], d.last_i))
        rows = predict_rows(d, model, oi, 'forward')
        agg = aggregate(rows)
        agg = agg[agg.scope == 'network']  # only network rows are logged (regions are shown, not scored)
        agg['month'] = str(m)
        seeded.append(agg)
        log('backtest %s: trained on data through %s, %d origins' % (m, d.dates[ci].date(), len(oi)))
    allrows = pd.concat(seeded, ignore_index=True)
    allrows = allrows[allrows.target <= d.dates[-1]]

    # score the three policies month by month, each policy learning only from earlier months
    results, finals = {}, {}
    for mode in ['equal', 'adaptive', 'adaptive_bias']:
        parts = []
        for m in allrows['month'].unique():
            cur = allrows[allrows.month == m].copy()
            as_of = cur['origin'].min()
            prior = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=LOG_COLS)
            pol = learn_policy(prior, act, as_of, mode)
            cur['final'] = apply_policy(cur, pol)
            parts.append(cur)
        res = pd.concat(parts, ignore_index=True)
        finals[mode] = res
        results[mode] = evaluate(res, act, d)
    skip_first = sorted(allrows['month'].unique())[:2]  # adaptive needs history; compare on the rest
    for mode in results:
        results[mode] = evaluate(finals[mode][~finals[mode].month.isin(skip_first)], act, d)
    score = {m: (r['day'] + r['week'] + r['month'] + r['landing14']) for m, r in results.items()}
    champion = min(score, key=score.get)
    log('policy comparison (MAPE, months after warm-up):')
    for m, r in results.items():
        log('  %-14s day %.1f%%  week %.1f%%  month %.1f%%  landing@14 %.1f%%%s' % (
            m, r['day'] * 100, r['week'] * 100, r['month'] * 100, r['landing14'] * 100, '  <- champion' if m == champion else ''))
    chosen = finals[champion]
    per_month = evaluate(chosen, act, d, per_month=True)
    out = {'generated': dt.date.today().isoformat(), 'dataThrough': str(last.date()),
           'months': [str(m) for m in sorted(allrows['month'].unique())], 'champion': champion,
           'policies': {m: {k: round(v, 4) for k, v in r.items()} for m, r in results.items()},
           'perMonth': per_month}
    with open(BT, 'w', encoding='utf-8') as f:
        json.dump(out, f, indent=1)
    lg = read_log()
    lg = lg[lg.source != 'backtest']
    seed = chosen.assign(source='backtest')[LOG_COLS]
    if len(lg):
        seed = seed[~seed.origin.isin(lg.origin.unique())]
    write_log(pd.concat([lg, seed], ignore_index=True))
    log('seeded %d backtest log rows; champion policy = %s' % (len(seed), champion))


def evaluate(res, act, d, per_month=False):
    x = res[res.scope == 'network'].copy()
    x['actual'] = x['target'].map(act['network'])
    x = x[x.target <= d.last]
    rows = []
    for m in sorted(x['month'].unique()):
        xm = x[x.month == m]
        start = pd.Period(m).start_time
        o0 = start - pd.Timedelta(days=1)
        end = pd.Period(m).end_time.normalize()
        full_month = end <= d.last
        r = {'month': m}
        day = xm[xm.h == 1]
        r['day'] = float((np.abs(day.final - day.actual) / day.actual).mean())
        wk = xm[xm.h <= 7].groupby('origin').filter(lambda g: len(g) == 7).groupby('origin')[['final', 'actual']].sum()
        r['week'] = float((np.abs(wk.final - wk.actual) / wk.actual).mean()) if len(wk) else np.nan
        mo = xm[(xm.origin == o0) & (xm.target <= end)]
        r['monthAhead'] = float(mo.final.sum() / mo.actual.sum() - 1) if full_month else np.nan
        o14 = start + pd.Timedelta(days=13)
        rest = xm[(xm.origin == o14) & (xm.target <= end)]
        mtd = act['network'][start:o14].sum()
        r['landing14'] = float((mtd + rest.final.sum()) / (mtd + rest.actual.sum()) - 1) if full_month and len(rest) else np.nan
        r['full'] = bool(full_month)
        rows.append(r)
    pm = pd.DataFrame(rows)
    if per_month:
        return [{k: (round(v, 4) if isinstance(v, float) and not np.isnan(v) else (None if isinstance(v, float) else v))
                 for k, v in r.items()} for r in rows]
    return {'day': float(pm['day'].mean()), 'week': float(pm['week'].mean()),
            'month': float(pm['monthAhead'].abs().mean()), 'landing14': float(pm['landing14'].abs().mean())}


# ---------------------------------------------------------------- daily run

def cmd_run(a):
    cfg = load_cfg()
    for path in a.sales or []:  # oldest first, e.g. full previous month then the new MTD at a month transition
        merge_into_history(path)
    hist = load_history()
    Y, last, reg = build_matrix(hist, cfg, extend_days=MAX_H)
    d = Data(Y, last, reg, cfg)
    act = actuals(d)
    log('training on %s..%s (%d stores in history)' % (d.dates[0].date(), last.date(), len(d.stores)))
    model = fit(d, d.last_i)
    rows = predict_rows(d, model, [d.last_i], 'forward')

    bt = json.load(open(BT, encoding='utf-8')) if os.path.exists(BT) else {'champion': 'equal'}
    lg = read_log()
    pol = learn_policy(lg, act, last, bt['champion'])
    log('policy %s: weights %s bias %s' % (pol['mode'], pol['weights'], pol['bias']))
    rows['final'] = apply_policy(rows, pol)
    agg = aggregate(rows)
    agg['final'] = apply_policy(agg, pol)

    live = agg[agg.scope == 'network'].assign(source='live')[LOG_COLS]
    lg = lg[~((lg.origin == last) & (lg.source == 'live'))]
    lg = lg[~((lg.origin == last) & (lg.source == 'backtest'))]
    lg = pd.concat([lg, live], ignore_index=True)
    write_log(lg)

    payload = build_payload(d, rows, agg, lg, act, pol, bt, cfg, hist)
    payload['targets'] = dashboard_targets(os.path.join(ROOT, a.dashboard), pd.Period(last, 'M'))
    if a.html:
        p = os.path.join(ROOT, a.html) if not os.path.isabs(a.html) else a.html
        with open(p, encoding='utf-8', newline='') as f:
            content = f.read()
        if 'const FORECAST = ' not in content:
            raise SystemExit('STOP: %s has no `const FORECAST = {...};` placeholder' % a.html)
        content = replace_const(content, 'FORECAST', json.dumps(payload, separators=(',', ':')))
        with open(p, 'w', encoding='utf-8', newline='') as f:
            f.write(content)
    m = payload['month']
    log('OK: forecast through %s -- %s lands at Rs%.2fCr (80%%: %s), %s forecast Rs%.2fCr' % (
        last.date(), m['key'], m['landing'] / 1e7,
        '%.2f-%.2fCr' % (m['lo'] / 1e7, m['hi'] / 1e7) if m['lo'] else 'n/a',
        payload['nextMonth']['key'], payload['nextMonth']['forecast'] / 1e7))


def backtest_view(bt, lg, act, l_method):
    """Backtest summary for the page, with the day-14 landing column recomputed using the landing method
    currently in use (the stored backtest scored the older 'latest' method)."""
    view = {k: bt.get(k) for k in ('champion', 'policies', 'perMonth', 'months', 'generated')}
    if view.get('perMonth'):
        l14 = landing14_by_month(lg, act, l_method, [m['month'] for m in view['perMonth'] if m.get('full')])
        view['perMonth'] = [dict(m, landing14=l14.get(m['month'], m.get('landing14'))) for m in view['perMonth']]
        vals = [abs(v) for v in l14.values()]
        if vals and view.get('policies') and view.get('champion') in view['policies']:
            view['policies'] = {k: dict(v) for k, v in view['policies'].items()}
            view['policies'][view['champion']]['landing14'] = round(float(np.mean(vals)), 4)
    return view


def build_payload(d, rows, agg, lg, act, pol, bt, cfg, hist):
    last = d.last
    F, Ac = error_matrix(lg, act, last)
    net = agg[agg.scope == 'network'].set_index('target')
    cur = pd.Period(last, 'M')
    nxt = cur + 1
    cur_end, nxt_start, nxt_end = cur.end_time.normalize(), nxt.start_time, nxt.end_time.normalize()
    days_left = (cur_end - last).days
    mtd = float(act['network'][cur.start_time:last].sum())
    rest_latest = float(net.loc[last + pd.Timedelta(days=1):cur_end, 'final'].sum())
    # month landing: learned choice of how to combine the start-of-month and newest forecasts
    l_phase_scores = score_landing(lg, act, last)
    l_choice = choose_landing(l_phase_scores)
    l_phase = landing_phase(last.day)
    l_method = l_choice[l_phase]
    l_scores, l_n = l_phase_scores[l_phase]
    rest_dates = pd.date_range(last + pd.Timedelta(days=1), cur_end)
    s0 = start_origin(list(lg.loc[lg.scope == 'network', 'origin'].unique()), cur.start_time)
    old = None
    if s0 is not None and days_left:
        o_rows = lg[(lg.scope == 'network') & (lg.origin == s0) & (lg.target.isin(rest_dates))]
        old = float(o_rows['final'].sum()) if o_rows['target'].nunique() == days_left else None
    new = tuple(float(net.loc[rest_dates, c].sum()) if days_left else 0.0 for c in ('final', 'A', 'G'))
    rest = landing_rest(l_method, new, old) if days_left else 0.0
    k_rest = rest / rest_latest if rest_latest else 1.0  # applied to regions/stores so they add up
    log('landing: %s-month phase -> %s (%s); phase scores %s over %d origins; start-of-month origin %s' % (
        l_phase, l_method, LANDING_METHODS[l_method], {m: round(v * 100, 2) for m, v in l_scores.items()}, l_n,
        s0.date() if s0 is not None else None))
    b_rest = band(F, Ac, 1, days_left) if days_left else [0.0, 0.0, 0]
    h0, h1 = (nxt_start - last).days, (nxt_end - last).days
    nm = float(net.loc[nxt_start:nxt_end, 'final'].sum())
    b_nm = band(F, Ac, h0, h1)
    ly = pd.Period(nxt - 12)
    ly_actual = float(act['network'][ly.start_time:ly.end_time.normalize()].sum())

    def rng(v, b, base=0.0):
        return (round(base + v * (1 + b[0])), round(base + v * (1 + b[1]))) if b else (None, None)

    lo, hi = rng(rest, b_rest, mtd)
    nlo, nhi = rng(nm, b_nm)

    daily = []
    for t in pd.date_range(cur.start_time, nxt_end):
        h = (t - last).days
        e = {'date': str(t.date())}
        if h <= 0:
            e['actual'] = round(float(act['network'].get(t, np.nan) or 0))
        else:
            e['forecast'] = round(float(net.loc[t, 'final'])) if t in net.index else None
            bb = band(F, Ac, h, h)
            if bb and e['forecast'] is not None:
                e['lo'], e['hi'] = round(e['forecast'] * (1 + bb[0])), round(e['forecast'] * (1 + bb[1]))
        daily.append(e)

    # per store / region landing for the current month
    stores = []
    mtd_store = d.Y.loc[cur.start_time:last].sum(min_count=1)
    rest_store = rows[rows.target <= cur_end].groupby('store')['final'].sum()
    nm_store = rows[(rows.target >= nxt_start) & (rows.target <= nxt_end)].groupby('store')['final'].sum()
    fb = rows.groupby('store')['fallback'].any()
    region_of = dict(zip(d.stores, np.array(d.region_names)[d.reg]))
    for s in sorted(set(rest_store.index) | set(mtd_store.dropna().index)):
        if s not in REGION_MAP:
            continue
        m_ = float(mtd_store.get(s, 0) or 0)
        r_ = float(rest_store.get(s, 0)) * k_rest
        stores.append({'store': s, 'region': region_of[s], 'mtd': round(m_), 'rest': round(r_), 'landing': round(m_ + r_),
                       'nextMonth': round(float(nm_store.get(s, 0))), 'runRate': bool(fb.get(s, False)),
                       'forecast': s in rest_store.index or s in nm_store.index})

    # accuracy
    x = lg[(lg.scope == 'network') & (lg.h == 1)].copy()
    x['actual'] = x['target'].map(act['network'])
    x = x.dropna(subset=['actual']).sort_values('target')
    x['err'] = x['final'] / x['actual'] - 1
    recent = [{'date': str(r.target.date()), 'forecast': round(r.final), 'actual': round(r.actual),
               'err': round(r.err, 4), 'source': r.source} for r in x.tail(14).itertuples()]

    def acc(src, days):
        y = x[(x.source == src) & (x.target > last - pd.Timedelta(days=days))]
        return {'n': int(len(y)), 'mape': round(float(y.err.abs().mean()), 4) if len(y) else None}

    net_log = lg[lg.scope == 'network']

    def acc_month(src):
        """Calendar-month miss (1st..last day), using the latest forecast of that source made in the
        7 days before the month started, once the whole month has actuals."""
        y = net_log[net_log.source == src]
        misses = []
        if y.empty:
            return {'n': 0, 'mape': None}
        for per in pd.period_range(pd.Period(y.origin.min(), 'M'), pd.Period(last, 'M')):
            ms, me = per.start_time, per.end_time.normalize()
            if me > last:
                continue
            cand = y[(y.origin < ms) & (y.origin >= ms - pd.Timedelta(days=7))]
            if cand.empty:
                continue
            z = cand[(cand.origin == cand.origin.max()) & (cand.target >= ms) & (cand.target <= me)]
            if z.target.nunique() != per.days_in_month:
                continue
            misses.append(abs(z.final.sum() / act['network'][ms:me].sum() - 1))
        return {'n': len(misses), 'mape': round(float(np.mean(misses)), 4) if misses else None}

    all_locs = set(hist['loc'])
    temp = {k: v for k, v in cfg.get('temporarily_closed', {}).items()
            if k in all_locs and k not in set(rows['store'])}
    new_stores = sorted(set(rows.loc[rows.fallback, 'store']))
    festivals = []
    for name, ds in cfg['festivals'].items():
        for x_ in ds:
            t = pd.Timestamp(x_)
            if last < t <= nxt_end + pd.Timedelta(days=20):
                festivals.append({'name': cfg.get('festival_labels', {}).get(name, name), 'date': str(t.date())})
    festivals.sort(key=lambda f: f['date'])

    return {
        'generatedOn': dt.date.today().isoformat(), 'dataThrough': str(last.date()),
        'policy': pol, 'backtest': backtest_view(bt, lg, act, l_choice['mid']),
        'month': {'key': str(cur), 'mtd': round(mtd), 'rest': round(rest), 'landing': round(mtd + rest),
                  'lo': lo, 'hi': hi, 'daysLeft': days_left},
        'nextMonth': {'key': str(nxt), 'forecast': round(nm), 'lo': nlo, 'hi': nhi, 'lastYear': round(ly_actual),
                      'lastYearKey': str(ly)},
        'regions': {r: {'mtd': round(float(act[r][cur.start_time:last].sum())),
                        'landing': round(float(act[r][cur.start_time:last].sum()) +
                                         k_rest * float(agg[(agg.scope == r) & (agg.target <= cur_end)]['final'].sum())),
                        'nextMonth': round(float(agg[(agg.scope == r) & (agg.target >= nxt_start) &
                                                     (agg.target <= nxt_end)]['final'].sum()))}
                    for r in d.region_names},
        'daily': daily, 'stores': stores,
        'accuracy': {'liveDay30': acc('live', 30), 'liveMonth': acc_month('live'),
                     'backtestDay': acc('backtest', 400), 'backtestMonth': acc_month('backtest'), 'recent': recent},
        'notes': {'temporarilyClosed': temp, 'runRateStores': new_stores, 'festivals': festivals,
                  'storeEvents': [e for e in cfg.get('store_events', [])
                                  if last - pd.Timedelta(days=60) <= pd.Timestamp(e['from']) <= last + pd.Timedelta(days=30)]},
        'insights': compute_insights(d, agg, rows, act, cfg, cur, nxt, mtd + rest, nm, k_rest),
        'landingMethod': {'method': l_method, 'label': LANDING_METHODS[l_method], 'n': l_n, 'phase': l_phase,
                          'scores': {m: round(v, 4) for m, v in l_scores.items()},
                          'labels': LANDING_METHODS, 'choice': l_choice, 'phases': LANDING_PHASES,
                          'phaseScores': {ph: {'n': n_, 'scores': {m: round(v, 4) for m, v in sc.items()}}
                                          for ph, (sc, n_) in l_phase_scores.items()}},
    }


# ---------------------------------------------------------------- targets (read from the main dashboard)

def dashboard_targets(path, cur):
    """Month / daily / store targets for month `cur`, read from index.html's TARGETS + DAILY_TARGETS
    constants with the same rule as its storeTargetForDay(): exact day-wise target if present, else the
    store's weekday/weekend figure. None when the dashboard's targets aren't for this month yet."""
    if not path or not os.path.exists(path):
        return None
    with open(path, encoding='utf-8') as f:
        content = f.read()
    dec = json.JSONDecoder()

    def const(name):
        i = content.find('const %s = ' % name)
        return dec.raw_decode(content, i + len('const %s = ' % name))[0] if i >= 0 else None

    T, DT = const('TARGETS'), const('DAILY_TARGETS') or {}
    days = pd.date_range(cur.start_time, cur.end_time.normalize())
    keys = [str(x.date()) for x in days]
    if not T or not any(k in v for v in DT.values() for k in keys):
        return None
    daily = dict.fromkeys(keys, 0.0)
    for rec in T:
        dt_ = DT.get(rec['pos_name'], {})
        for x, k in zip(days, keys):
            daily[k] += dt_[k] if k in dt_ else (rec['weekend_target'] if x.dayofweek >= 5 else rec['weekday_target'])
    return {'monthKey': str(cur), 'month': round(sum(daily.values())), 'daily': {k: round(v) for k, v in daily.items()},
            'stores': {r['pos_name']: r['month_target'] for r in T}}


# ---------------------------------------------------------------- insights ("why this forecast")

def _month_slice(obs, per):
    return obs.loc[per.start_time:per.end_time.normalize()]


def _weekend_days(per):
    ds = pd.date_range(per.start_time, per.end_time.normalize())
    we = int((ds.dayofweek >= 5).sum())
    return len(ds) - we, we


def compute_insights(d, agg, rows, act, cfg, cur, nxt, landing, nm_fc, k_rest=1.0):
    """Data-derived explanations for the forecast page. Everything here is computed from the sales
    history on each run -- no hand-written claims -- so it stays true as the data moves."""
    obs = d.Y.loc[:d.last]
    last = d.last
    regs = pd.Series(np.array(d.region_names)[d.reg], index=d.stores)
    groups = {'network': list(d.stores), **{r: list(regs[regs == r].index) for r in d.region_names}}

    # weekend multiplier (Sat/Sun vs Mon-Fri revenue per trading store-day), last 91 days
    win = obs.loc[last - pd.Timedelta(days=90):last]
    wk = win.index.dayofweek >= 5
    wmult = {g: float(np.nanmean(win.loc[wk, c].values) / np.nanmean(win.loc[~wk, c].values)) for g, c in groups.items()}

    cal_n, cal_c = _weekend_days(nxt), _weekend_days(cur)

    def cal_factor(pn, pc, m):
        (wn, en), (wc, ec) = _weekend_days(pn), _weekend_days(pc)
        return ((wn + en * m) / (wn + en)) / ((wc + ec * m) / (wc + ec))

    # same-store seasonality: next-month / this-month in prior years (stores trading every day of both)
    seasonal = {}
    for g, cols in groups.items():
        out = []
        for k in (1, 2, 3):
            pc, pn = cur - 12 * k, nxt - 12 * k
            if pc.start_time < obs.index[0]:
                continue
            a, b = _month_slice(obs, pc)[cols], _month_slice(obs, pn)[cols]
            keep = [c for c in cols if a[c].notna().all() and b[c].notna().all()]
            if len(keep) < 3:
                continue
            raw = float(b[keep].sum().sum() / a[keep].sum().sum())
            adj = raw / cal_factor(pn, pc, wmult[g])
            out.append({'year': pc.year, 'raw': round(raw, 3), 'calAdj': round(adj, 3), 'stores': len(keep)})
        seasonal[g] = out

    # same-store YoY growth, last 91 days vs the same weekdays a year earlier
    yoy = {}
    for g, cols in groups.items():
        now = obs.loc[last - pd.Timedelta(days=90):last, cols]
        ly = obs.loc[last - pd.Timedelta(days=90 + 364):last - pd.Timedelta(days=364), cols]
        keep = [c for c in cols if now[c].notna().mean() > 0.9 and ly[c].notna().mean() > 0.9]
        yoy[g] = {'growth': round(float(now[keep].sum().sum() / ly[keep].sum().sum() - 1), 4), 'stores': len(keep)} if len(keep) >= 3 else None

    # festival effects: festival window vs the 28 days just before it, per past occurrence, per region
    wins = {k: tuple(v) for k, v in cfg.get('festival_windows', {}).items()}
    labels = cfg.get('festival_labels', {})
    fests = []
    for name, ds in cfg['festivals'].items():
        if name in ('gandhi', 'independence', 'republic'):
            continue  # one-day holidays that overlap the big festivals; their lift can't be separated out
        lo, hi = wins.get(name, (-10, 5))
        upcoming = [pd.Timestamp(x) for x in ds if last < pd.Timestamp(x) + pd.Timedelta(days=hi)
                    and pd.Timestamp(x) + pd.Timedelta(days=lo) <= nxt.end_time.normalize()]
        if not upcoming:
            continue
        up = upcoming[0]
        days_in_next = int(((pd.date_range(up + pd.Timedelta(days=lo), up + pd.Timedelta(days=hi)) >= nxt.start_time) &
                            (pd.date_range(up + pd.Timedelta(days=lo), up + pd.Timedelta(days=hi)) <= nxt.end_time.normalize())).sum())
        past = []
        for x in ds:
            D = pd.Timestamp(x)
            w0, w1 = D + pd.Timedelta(days=lo), D + pd.Timedelta(days=hi)
            if w1 > last or w0 - pd.Timedelta(days=28) < obs.index[0]:
                continue
            e = {'date': str(D.date())}
            for g, cols in groups.items():
                w = obs.loc[w0:w1, cols]
                b = obs.loc[w0 - pd.Timedelta(days=28):w0 - pd.Timedelta(days=1), cols]
                e[g] = round(float(np.nanmean(w.values) / np.nanmean(b.values) - 1), 3)
            past.append(e)
        prev = [str(pd.Timestamp(x).date()) for x in ds if pd.Timestamp(x) < up][-2:]
        fests.append({'name': labels.get(name, name), 'date': str(up.date()), 'window': [lo, hi], 'previous': prev,
                      'daysInNextMonth': days_in_next, 'windowDays': hi - lo + 1, 'past': past})
    fests.sort(key=lambda f: f['date'])

    # the three methods for next month (why the blend lands where it does)
    net = agg[agg.scope == 'network']
    nmask = (net.target >= nxt.start_time) & (net.target <= nxt.end_time.normalize())
    methods = {c: round(float(net.loc[nmask, c].sum())) for c in 'ABG'}

    # model-implied next/this month ratio per group, calendar-adjusted, for comparison with history
    implied = {}
    cur_mask = lambda s: (agg.scope == s) & (agg.target <= cur.end_time.normalize())  # noqa: E731
    for g in groups:
        mtd_g = float(act[g][cur.start_time:last].sum())
        land_g = mtd_g + k_rest * float(agg.loc[cur_mask(g), 'final'].sum())
        nm_g = float(agg[(agg.scope == g) & (agg.target >= nxt.start_time) & (agg.target <= nxt.end_time.normalize())]['final'].sum())
        if land_g > 0:
            dn, dc = nxt.days_in_month, cur.days_in_month
            per_day = (nm_g / dn) / (land_g / dc)
            implied[g] = {'perDay': round(per_day, 3), 'calAdj': round(per_day / cal_factor(nxt, cur, wmult[g]), 3)}

    # bridge: this month's landing -> next month's forecast
    pace = landing / cur.days_in_month
    same_pace = pace * nxt.days_in_month
    calf = cal_factor(nxt, cur, wmult['network'])
    bridge = {'landing': round(landing), 'daysCur': cur.days_in_month, 'daysNext': nxt.days_in_month,
              'samePace': round(same_pace), 'calendar': round(same_pace * (calf - 1)),
              'learned': round(nm_fc - same_pace * calf), 'forecast': round(nm_fc),
              'weekendsCur': cal_c[1], 'weekendsNext': cal_n[1]}
    hist = [x['calAdj'] for x in seasonal['network']]
    if hist:
        # "if the past pattern repeats": this month's pace x calendar x average past next/this-month ratio
        bridge['historyRatio'] = round(float(np.mean(hist)), 3)
        bridge['historyScenario'] = round(same_pace * calf * float(np.mean(hist)))

    # stores drifting: last 28 days vs the same 28 days a year earlier (removes seasonality), per trading day
    r28 = obs.loc[last - pd.Timedelta(days=27):last].mean()
    p28 = obs.loc[last - pd.Timedelta(days=27 + 364):last - pd.Timedelta(days=364)].mean()
    live = set(rows['store'])
    drift = []
    for s in d.stores:
        if s in live and s in REGION_MAP and p28.get(s, np.nan) > 20000 and not np.isnan(r28.get(s, np.nan)):
            drift.append({'store': s, 'region': regs[s], 'recent': round(float(r28[s])), 'prior': round(float(p28[s])),
                          'change': round(float(r28[s] / p28[s] - 1), 3)})
    drift.sort(key=lambda x: x['change'])

    return {'weekendMultiplier': {g: round(v, 2) for g, v in wmult.items()}, 'seasonal': seasonal, 'implied': implied,
            'yoy': yoy, 'festivals': fests, 'methods': methods, 'bridge': bridge,
            'drift': {'down': drift[:4], 'up': drift[::-1][:4]}}


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)
    p = sub.add_parser('init-history')
    p.add_argument('--sales', required=True)
    p = sub.add_parser('backtest')
    p.add_argument('--months', type=int, default=18)
    p = sub.add_parser('run')
    p.add_argument('--sales', action='append', help='sales CSV to merge into the history first; repeatable (oldest first)')
    p.add_argument('--html', help='forecast page to write const FORECAST into')
    p.add_argument('--dashboard', default='index.html', help='main dashboard to read TARGETS / DAILY_TARGETS from')
    a = ap.parse_args()
    os.makedirs(DATA, exist_ok=True)
    if a.cmd == 'init-history':
        g = aggregate_sales(a.sales)
        g.to_csv(HIST, index=False)
        log('wrote %s: %d rows, %s..%s, %d locations' % (HIST, len(g), g['date'].min(), g['date'].max(), g['loc'].nunique()))
    elif a.cmd == 'backtest':
        cmd_backtest(a)
    else:
        cmd_run(a)


if __name__ == '__main__':
    main()
