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
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
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

    # Splice RMZ + IYR
    rmz_close = extract_close(rmz)
    iyr_close = extract_close(iyr)
    iyr_start = iyr_close.index[0]
    reit_close = pd.concat([rmz_close[rmz_close.index < iyr_start], iyr_close]).sort_index()
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
    df = df.dropna()

    # Winsorized version
    df_w = df.copy()
    for col in ['excess_ret', 'oil_chg']:
        df_w[col] = mstats.winsorize(df_w[col], limits=[0.01, 0.01])

    result = {'df': df, 'df_w': df_w, 'oil_monthly': oil_monthly,
              'reit_monthly': reit_monthly, 'spx_monthly': spx_monthly,
              't3m_monthly': t3m_monthly, 't10y_monthly': t10y_monthly}

    with open(CACHE_PATH, 'wb') as f:
        pickle.dump(result, f)

    return result


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

    return results


if __name__ == '__main__':
    data = fetch_all()
    print(f"Data: {len(data['df'])} months, {data['df'].index[0].strftime('%Y-%m')} to {data['df'].index[-1].strftime('%Y-%m')}")
    regs = run_regressions(data)
    for name, model in regs.items():
        print(f"\n{'='*60}\n{name}  (R²={model.rsquared:.3f})")
        print(model.summary().tables[1])
