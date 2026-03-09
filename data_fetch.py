"""
Fetch and prepare all data for the oil/REIT/rates analysis.
Cached to avoid re-downloading on every request.
"""

import pandas as pd
import numpy as np
import yfinance as yf
import requests
from io import StringIO
from scipy.stats import mstats
import statsmodels.api as sm
import os, pickle, time

CACHE_PATH = os.path.join(os.path.dirname(__file__), 'cached_data.pkl')
CACHE_TTL = 86400  # 24 hours


def get_fred(series_id, start='1995-01-01', end='2025-12-31'):
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}&cosd={start}&coed={end}"
    for attempt in range(3):
        try:
            resp = requests.get(url, timeout=60)
            resp.raise_for_status()
            break
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
            if attempt == 2:
                raise
            time.sleep(2)
    df = pd.read_csv(StringIO(resp.text), parse_dates=['observation_date'], index_col='observation_date')
    df.columns = [series_id]
    df[series_id] = pd.to_numeric(df[series_id], errors='coerce')
    return df


def fetch_all():
    """Return dict with all DataFrames needed for the analysis."""
    # Check cache
    if os.path.exists(CACHE_PATH):
        age = time.time() - os.path.getmtime(CACHE_PATH)
        if age < CACHE_TTL:
            with open(CACHE_PATH, 'rb') as f:
                return pickle.load(f)

    start, end = '1995-01-01', '2025-12-31'

    # FRED data
    oil_raw = get_fred('DCOILWTICO', start, end)
    t3m_raw = get_fred('DTB3', start, end)
    t10y_raw = get_fred('DGS10', start, end)

    # yfinance
    spx = yf.download('^GSPC', start=start, end=end, progress=False)
    rmz = yf.download('^RMZ', start=start, end=end, progress=False)
    iyr = yf.download('IYR', start=start, end=end, progress=False)

    # Monthly aggregation
    oil_monthly = oil_raw['DCOILWTICO'].resample('ME').mean()
    oil_pct = oil_monthly.pct_change() * 100
    t3m_monthly = t3m_raw['DTB3'].resample('ME').mean()
    t10y_monthly = t10y_raw['DGS10'].resample('ME').mean()

    def extract_close(df):
        if isinstance(df.columns, pd.MultiIndex):
            return df['Close'].iloc[:, 0]
        return df['Close']

    # Splice RMZ + IYR (level-adjusted so no fake return at splice point)
    rmz_close = extract_close(rmz)
    iyr_close = extract_close(iyr)
    iyr_start = iyr_close.index[0]
    # Find the last RMZ price on or before IYR's first date to compute scale factor
    rmz_before = rmz_close[rmz_close.index <= iyr_start]
    if len(rmz_before) > 0:
        scale = float(rmz_before.iloc[-1]) / float(iyr_close.iloc[0])
        iyr_scaled = iyr_close * scale
    else:
        iyr_scaled = iyr_close
    reit_close = pd.concat([rmz_close[rmz_close.index < iyr_start], iyr_scaled]).sort_index()
    reit_close = reit_close[~reit_close.index.duplicated(keep='last')]

    reit_monthly = reit_close.resample('ME').last()
    reit_ret = reit_monthly.pct_change() * 100

    spx_close = extract_close(spx)
    spx_monthly = spx_close.resample('ME').last()
    spx_ret = spx_monthly.pct_change() * 100

    # === Main DataFrame ===
    df = pd.DataFrame({
        'oil_price': oil_monthly,
        'oil_chg': oil_pct,
        'reit_ret': reit_ret,
        'spx_ret': spx_ret,
        't3m': t3m_monthly,
        't10y': t10y_monthly,
    })
    df['excess_ret'] = df['reit_ret'] - df['spx_ret']
    df['term_spread'] = df['t10y'] - df['t3m']
    df['d_t3m'] = df['t3m'].diff()
    df['d_t10y'] = df['t10y'].diff()
    df['d_spread'] = df['term_spread'].diff()
    # 3-month trailing oil price change (computed on daily, resampled to monthly max)
    oil_daily = oil_raw['DCOILWTICO'].dropna()
    oil_3m_chg_daily = (oil_daily / oil_daily.shift(63) - 1) * 100  # ~63 trading days = 3 months
    df['oil_3m_chg'] = oil_3m_chg_daily.resample('ME').max()
    df = df.dropna()

    # Winsorized version
    df_w = df.copy()
    for col in ['excess_ret', 'oil_chg']:
        df_w[col] = mstats.winsorize(df_w[col], limits=[0.01, 0.01])

    # Daily series for trough-to-peak calculations
    t3m_daily = t3m_raw['DTB3'].dropna()
    t10y_daily = t10y_raw['DGS10'].dropna()

    # Build daily DataFrame
    df_daily = pd.DataFrame({
        'oil_price': oil_daily,
        'oil_3m_chg': oil_3m_chg_daily,
        'reit_close': reit_close,
        'spx_close': spx_close,
        't3m': t3m_daily,
        't10y': t10y_daily,
    }).dropna()

    result = {'df': df, 'df_w': df_w, 'df_daily': df_daily,
              'oil_monthly': oil_monthly, 'oil_daily': oil_daily,
              'reit_close': reit_close, 'spx_close': spx_close,
              't3m_daily': t3m_daily, 't10y_daily': t10y_daily,
              'reit_monthly': reit_monthly, 'spx_monthly': spx_monthly,
              't3m_monthly': t3m_monthly, 't10y_monthly': t10y_monthly}

    with open(CACHE_PATH, 'wb') as f:
        pickle.dump(result, f)

    return result


def detect_oil_shocks(oil_daily):
    """Auto-detect oil shock episodes using 3-month trailing daily oil price change > 30%.

    Args:
        oil_daily: Daily oil price series (pd.Series with DatetimeIndex)

    Returns list of dicts: [{label, start, end, trough_price, peak_price, pct_change}, ...]
    """
    oil = oil_daily.dropna()
    chg_3m = (oil / oil.shift(63) - 1) * 100  # 63 trading days ~ 3 months
    flagged = chg_3m > 30

    # Cluster consecutive flagged days into episodes (allow up to 30 calendar day gaps)
    flagged_dates = flagged[flagged].index.tolist()
    if not flagged_dates:
        return []

    clusters = [[flagged_dates[0]]]
    for dt in flagged_dates[1:]:
        gap = (dt - clusters[-1][-1]).days
        if gap <= 30:
            clusters[-1].append(dt)
        else:
            clusters.append([dt])

    shocks = []
    for cluster in clusters:
        cluster_start = cluster[0]
        cluster_end = cluster[-1]

        # Peak = highest daily oil price within the cluster
        cluster_prices = oil.loc[cluster_start:cluster_end]
        peak_date = cluster_prices.idxmax()
        peak_price = float(cluster_prices.loc[peak_date])

        # Trough = lowest daily oil price in 4 months before peak
        trough_search_start = peak_date - pd.DateOffset(months=4)
        trough_window = oil.loc[trough_search_start:peak_date]
        if len(trough_window) == 0:
            trough_window = oil.loc[:peak_date].tail(84)
        trough_date = trough_window.idxmin()
        trough_price = float(trough_window.loc[trough_date])

        pct_change = (peak_price / trough_price - 1) * 100 if trough_price > 0 else 0

        # Convert to month boundaries for regression masking
        start_month = trough_date.strftime('%Y-%m')
        end_month = peak_date.strftime('%Y-%m')

        label = f"{start_month} to {end_month}"
        shocks.append({
            'label': label,
            'start': start_month,
            'end': end_month,
            'start_date': trough_date.strftime('%Y-%m-%d'),
            'end_date': peak_date.strftime('%Y-%m-%d'),
            'trough_price': round(trough_price, 2),
            'peak_price': round(peak_price, 2),
            'pct_change': round(pct_change, 1),
        })

    # Merge overlapping episodes (same or overlapping month ranges)
    if len(shocks) <= 1:
        return shocks

    merged = [shocks[0]]
    for s in shocks[1:]:
        prev = merged[-1]
        # Overlap if new start <= previous end (month-level comparison)
        if s['start'] <= prev['end']:
            # Keep the wider window with the bigger move
            if s['pct_change'] > prev['pct_change']:
                merged[-1] = s
            # Extend end if new episode goes further
            if s['end'] > prev['end']:
                merged[-1]['end'] = s['end']
                merged[-1]['end_date'] = s['end_date']
                merged[-1]['label'] = f"{merged[-1]['start']} to {s['end']}"
        else:
            merged.append(s)

    return merged


def run_regressions(data):
    """Run all regressions and return results dict."""
    df = data['df']
    df_w = data['df_w']
    results = {}

    # --- REIT vs S&P regressions ---

    # Model 1: Base (oil change + rate levels)
    X = sm.add_constant(df[['oil_chg', 't3m', 't10y']])
    results['reit_m1'] = sm.OLS(df['excess_ret'], X).fit(cov_type='HC1')

    # Model 2: Winsorized + rate changes (best spec)
    X = sm.add_constant(df_w[['oil_chg', 'd_t3m', 'd_t10y']])
    results['reit_m2'] = sm.OLS(df_w['excess_ret'], X).fit(cov_type='HC1')

    # Model 3: Asymmetric oil + rate changes (winsorized)
    df_w2 = df_w.copy()
    df_w2['oil_up'] = df_w2['oil_chg'].clip(lower=0)
    df_w2['oil_down'] = df_w2['oil_chg'].clip(upper=0)
    X = sm.add_constant(df_w2[['oil_up', 'oil_down', 'd_t3m', 'd_t10y']])
    results['reit_m3'] = sm.OLS(df_w2['excess_ret'], X).fit(cov_type='HC1')

    # Full sample HC1 model (same spec as shock regressions, for apples-to-apples comparison)
    X = sm.add_constant(df[['oil_chg', 'd_t3m', 'd_t10y']])
    results['reit_full_hc1'] = sm.OLS(df['excess_ret'], X).fit(cov_type='HC1')

    # Models for Excel LINEST cross-check (no HC1 — LINEST uses standard OLS SEs)
    X = sm.add_constant(df[['oil_chg', 'd_t3m', 'd_t10y']])
    results['reit_m1_levels_chg'] = sm.OLS(df['excess_ret'], X).fit()

    X = sm.add_constant(df[['oil_chg']])
    results['t10y_on_oil_ols'] = sm.OLS(df['d_t10y'], X).fit()
    results['t3m_on_oil_ols'] = sm.OLS(df['d_t3m'], X).fit()

    X = sm.add_constant(df[['d_t3m', 'd_t10y']])
    results['oil_rates_changes_ols'] = sm.OLS(df['oil_chg'], X).fit()

    # --- Oil vs Interest Rates regressions ---

    # Model A: Oil change ~ rate levels
    X = sm.add_constant(df[['t3m', 't10y']])
    results['oil_rates_levels'] = sm.OLS(df['oil_chg'], X).fit(cov_type='HC1')

    # Model B: Oil change ~ rate changes
    X = sm.add_constant(df[['d_t3m', 'd_t10y']])
    results['oil_rates_changes'] = sm.OLS(df['oil_chg'], X).fit(cov_type='HC1')

    # Model C: Oil change ~ rate changes + spread
    X = sm.add_constant(df[['d_t3m', 'd_t10y', 'term_spread']])
    results['oil_rates_spread'] = sm.OLS(df['oil_chg'], X).fit(cov_type='HC1')

    # Model D: Rate changes ~ oil (reverse causality check)
    X = sm.add_constant(df[['oil_chg']])
    results['t10y_on_oil'] = sm.OLS(df['d_t10y'], X).fit(cov_type='HC1')
    results['t3m_on_oil'] = sm.OLS(df['d_t3m'], X).fit(cov_type='HC1')

    # Model E: Lagged oil -> rate changes (does oil predict future rate moves?)
    df_lag = df.copy()
    df_lag['oil_chg_lag1'] = df_lag['oil_chg'].shift(1)
    df_lag['oil_chg_lag2'] = df_lag['oil_chg'].shift(2)
    df_lag['oil_chg_lag3'] = df_lag['oil_chg'].shift(3)
    df_lag = df_lag.dropna()
    X = sm.add_constant(df_lag[['oil_chg_lag1', 'oil_chg_lag2', 'oil_chg_lag3']])
    results['t10y_oil_lagged'] = sm.OLS(df_lag['d_t10y'], X).fit(cov_type='HC1')
    results['t3m_oil_lagged'] = sm.OLS(df_lag['d_t3m'], X).fit(cov_type='HC1')

    # --- Shock episode analysis (auto-detected via 3-month trailing daily oil change > 30%) ---
    detected_shocks = detect_oil_shocks(data['oil_daily'])

    # Compute trough-to-peak returns using exact daily price data
    reit_close = data['reit_close']
    spx_close = data['spx_close']
    t3m_daily = data['t3m_daily']
    t10y_daily = data['t10y_daily']

    def _nearest(series, date):
        """Get value at or nearest prior date in series."""
        ts = pd.Timestamp(date)
        idx = series.index.searchsorted(ts)
        if idx >= len(series):
            idx = len(series) - 1
        # If exact match, use it; otherwise use prior trading day
        if series.index[idx] == ts:
            return float(series.iloc[idx]), series.index[idx]
        if idx > 0:
            return float(series.iloc[idx - 1]), series.index[idx - 1]
        return float(series.iloc[idx]), series.index[idx]

    for shock in detected_shocks:
        trough_dt = shock['start_date']
        peak_dt = shock['end_date']

        reit_t, _ = _nearest(reit_close, trough_dt)
        reit_p, _ = _nearest(reit_close, peak_dt)
        spx_t, _ = _nearest(spx_close, trough_dt)
        spx_p, _ = _nearest(spx_close, peak_dt)
        t3m_t, _ = _nearest(t3m_daily, trough_dt)
        t3m_p, _ = _nearest(t3m_daily, peak_dt)
        t10y_t, _ = _nearest(t10y_daily, trough_dt)
        t10y_p, _ = _nearest(t10y_daily, peak_dt)

        shock['reit_ret'] = round((reit_p / reit_t - 1) * 100, 2) if reit_t else None
        shock['spx_ret'] = round((spx_p / spx_t - 1) * 100, 2) if spx_t else None
        shock['excess_ret'] = round(shock['reit_ret'] - shock['spx_ret'], 2) if shock['reit_ret'] is not None and shock['spx_ret'] is not None else None
        shock['d_t10y'] = round(t10y_p - t10y_t, 3)
        shock['d_t3m'] = round(t3m_p - t3m_t, 3)
        # Trading days in the episode
        oil_daily = data['oil_daily']
        shock['trading_days'] = len(oil_daily.loc[trough_dt:peak_dt])

    results['_detected_shocks'] = detected_shocks

    # --- Post-shock recovery: 3, 6, 12 months after each shock peak using daily data ---
    post_shock = []
    for shock in detected_shocks:
        peak_dt = pd.Timestamp(shock['end_date'])

        for horizon, cal_days in [('3M', 91), ('6M', 182), ('12M', 365)]:
            end_dt = peak_dt + pd.DateOffset(days=cal_days)

            reit_p, _ = _nearest(reit_close, peak_dt)
            reit_e, actual_end = _nearest(reit_close, end_dt)
            spx_p, _ = _nearest(spx_close, peak_dt)
            spx_e, _ = _nearest(spx_close, end_dt)
            oil_p, _ = _nearest(data['oil_daily'], peak_dt)
            oil_e, _ = _nearest(data['oil_daily'], end_dt)
            t10y_p, _ = _nearest(t10y_daily, peak_dt)
            t10y_e, _ = _nearest(t10y_daily, end_dt)
            t3m_p, _ = _nearest(t3m_daily, peak_dt)
            t3m_e, _ = _nearest(t3m_daily, end_dt)

            if reit_p and spx_p and oil_p:
                reit_cum = round((reit_e / reit_p - 1) * 100, 2)
                spx_cum = round((spx_e / spx_p - 1) * 100, 2)
                post_shock.append({
                    'shock': shock['label'],
                    'shock_end': shock['end_date'],
                    'horizon': horizon,
                    'end_date': actual_end.strftime('%Y-%m-%d'),
                    'reit_ret': reit_cum,
                    'spx_ret': spx_cum,
                    'excess_ret': round(reit_cum - spx_cum, 2),
                    'oil_chg': round((oil_e / oil_p - 1) * 100, 2),
                    'd_t10y': round(t10y_e - t10y_p, 3),
                    'd_t3m': round(t3m_e - t3m_p, 3),
                })

    # Averages across all shocks by horizon
    post_shock_avg = {}
    for horizon in ['3M', '6M', '12M']:
        rows = [r for r in post_shock if r['horizon'] == horizon]
        if rows:
            post_shock_avg[horizon] = {
                'n': len(rows),
                'reit_ret': round(np.mean([r['reit_ret'] for r in rows]), 2),
                'spx_ret': round(np.mean([r['spx_ret'] for r in rows]), 2),
                'excess_ret': round(np.mean([r['excess_ret'] for r in rows]), 2),
                'oil_chg': round(np.mean([r['oil_chg'] for r in rows]), 2),
                'd_t10y': round(np.mean([r['d_t10y'] for r in rows]), 3),
                'd_t3m': round(np.mean([r['d_t3m'] for r in rows]), 3),
            }

    results['_post_shock'] = post_shock
    results['_post_shock_avg'] = post_shock_avg

    return results


if __name__ == '__main__':
    data = fetch_all()
    print(f"Data: {len(data['df'])} months, {data['df'].index[0].strftime('%Y-%m')} to {data['df'].index[-1].strftime('%Y-%m')}")
    regs = run_regressions(data)
    print(f"\nDetected {len(regs['_detected_shocks'])} oil shock episodes:")
    for s in regs['_detected_shocks']:
        print(f"  {s['label']}: ${s['trough_price']} -> ${s['peak_price']} (+{s['pct_change']}%)")
    for name, model in regs.items():
        if name.startswith('_'):
            continue
        print(f"\n{'='*60}\n{name}  (R²={model.rsquared:.3f})")
        print(model.summary().tables[1])
