"""
Oil Shock Analysis: Same regressions but only during oil shock periods.
Defines shocks multiple ways to see if the relationship strengthens.
"""

import warnings
warnings.filterwarnings('ignore')

import pandas as pd
import numpy as np
import statsmodels.api as sm
from data_fetch import fetch_all, run_regressions

print("Loading data...")
data = fetch_all()
df = data['df'].copy()

print(f"Full sample: {len(df)} months ({df.index[0].strftime('%Y-%m')} to {df.index[-1].strftime('%Y-%m')})")
print(f"Oil monthly change: mean={df['oil_chg'].mean():.1f}%, std={df['oil_chg'].std():.1f}%")

# ================================================================
# Define oil shocks multiple ways
# ================================================================
oil_std = df['oil_chg'].std()
oil_mean = df['oil_chg'].mean()

shock_defs = {
    '1 SD move (|oil chg| > 1 std dev)': df['oil_chg'].abs() > oil_std,
    '1.5 SD move': df['oil_chg'].abs() > 1.5 * oil_std,
    'Top/Bottom 10% of oil moves': (df['oil_chg'] <= df['oil_chg'].quantile(0.10)) | (df['oil_chg'] >= df['oil_chg'].quantile(0.90)),
    'Top/Bottom 20% of oil moves': (df['oil_chg'] <= df['oil_chg'].quantile(0.20)) | (df['oil_chg'] >= df['oil_chg'].quantile(0.80)),
    'Oil spike (>10% monthly)': df['oil_chg'] > 10,
    'Oil crash (<-10% monthly)': df['oil_chg'] < -10,
    'Any big move (|chg| > 10%)': df['oil_chg'].abs() > 10,
}

# Also define named historical shock windows
historical_shocks = {
    'Gulf War (1990-08 to 1991-03)': ('1990-08', '1991-03'),
    'Asian Crisis (1997-10 to 1998-12)': ('1997-10', '1998-12'),
    'Dot-com / 9-11 (2001-01 to 2002-01)': ('2001-01', '2002-01'),
    'Oil Super-Spike (2007-01 to 2008-07)': ('2007-01', '2008-07'),
    'GFC Crash (2008-07 to 2009-04)': ('2008-07', '2009-04'),
    'Oil Glut (2014-07 to 2016-02)': ('2014-07', '2016-02'),
    'COVID Crash (2020-02 to 2020-06)': ('2020-02', '2020-06'),
    'Post-COVID Spike (2021-01 to 2022-06)': ('2021-01', '2022-06'),
}

# Build a combined "any named shock" mask
any_historical = pd.Series(False, index=df.index)
for name, (start, end) in historical_shocks.items():
    mask = (df.index >= start) & (df.index <= end)
    any_historical = any_historical | mask

shock_defs['Any historical shock window'] = any_historical


def run_shock_regression(subset, label):
    """Run the key regressions on a subset and print results."""
    n = len(subset)
    if n < 10:
        print(f"  {label}: only {n} obs, skipping\n")
        return None

    print(f"\n{'-'*70}")
    print(f"  {label}")
    print(f"  {n} months | Oil chg range: {subset['oil_chg'].min():.1f}% to {subset['oil_chg'].max():.1f}%")
    print(f"  Avg REIT excess return: {subset['excess_ret'].mean():.2f}% | Avg S&P: {subset['spx_ret'].mean():.2f}%")
    print(f"{'-'*70}")

    results = {}

    # Regression 1: REIT excess ~ oil change + rate changes
    try:
        X = sm.add_constant(subset[['oil_chg', 'd_t3m', 'd_t10y']])
        m = sm.OLS(subset['excess_ret'], X).fit(cov_type='HC1')
        results['reit'] = m
        print(f"\n  REIT Excess Return ~ Oil + Rate Changes")
        print(f"  R² = {m.rsquared:.3f} | N = {n}")
        for var in m.params.index:
            p = m.pvalues[var]
            sig = '***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.1 else ''
            print(f"    {var:15s}  coef={m.params[var]:8.4f}  p={p:.4f} {sig}")
    except Exception as e:
        print(f"  REIT regression failed: {e}")

    # Regression 2: 10Y change ~ oil change
    try:
        X = sm.add_constant(subset[['oil_chg']])
        m2 = sm.OLS(subset['d_t10y'], X).fit(cov_type='HC1')
        results['t10y'] = m2
        p_oil = m2.pvalues['oil_chg']
        sig = '***' if p_oil < 0.01 else '**' if p_oil < 0.05 else '*' if p_oil < 0.1 else ''
        print(f"\n  10Y Rate Change ~ Oil Change")
        print(f"  R² = {m2.rsquared:.3f} | oil coef = {m2.params['oil_chg']:.4f}  p={p_oil:.4f} {sig}")
    except Exception as e:
        print(f"  10Y regression failed: {e}")

    # Regression 3: 3M change ~ oil change
    try:
        X = sm.add_constant(subset[['oil_chg']])
        m3 = sm.OLS(subset['d_t3m'], X).fit(cov_type='HC1')
        results['t3m'] = m3
        p_oil = m3.pvalues['oil_chg']
        sig = '***' if p_oil < 0.01 else '**' if p_oil < 0.05 else '*' if p_oil < 0.1 else ''
        print(f"\n  3M Rate Change ~ Oil Change")
        print(f"  R² = {m3.rsquared:.3f} | oil coef = {m3.params['oil_chg']:.4f}  p={p_oil:.4f} {sig}")
    except Exception as e:
        print(f"  3M regression failed: {e}")

    return results


# ================================================================
# Run it
# ================================================================
print(f"\n{'='*70}")
print("  FULL SAMPLE (baseline for comparison)")
print(f"{'='*70}")
run_shock_regression(df, f"All {len(df)} months")

print(f"\n\n{'='*70}")
print("  STATISTICAL SHOCK DEFINITIONS")
print(f"{'='*70}")
for label, mask in shock_defs.items():
    subset = df[mask]
    run_shock_regression(subset, label)

print(f"\n\n{'='*70}")
print("  INDIVIDUAL HISTORICAL SHOCK WINDOWS")
print(f"{'='*70}")
for label, (start, end) in historical_shocks.items():
    mask = (df.index >= start) & (df.index <= end)
    subset = df[mask]
    if len(subset) > 0:
        avg_oil = subset['oil_chg'].mean()
        avg_excess = subset['excess_ret'].mean()
        avg_10y = subset['d_t10y'].mean()
        print(f"\n  {label}")
        print(f"    {len(subset)} months | Avg oil chg: {avg_oil:+.1f}% | Avg REIT excess: {avg_excess:+.2f}% | Avg 10Y chg: {avg_10y:+.3f}pp")

# ================================================================
# Comparison table: Full sample vs shock months
# ================================================================
print(f"\n\n{'='*70}")
print("  COMPARISON: FULL SAMPLE vs OIL SHOCK MONTHS")
print(f"{'='*70}")

# Use 1 SD as the primary shock definition
shock_mask = df['oil_chg'].abs() > oil_std
shock = df[shock_mask]
calm = df[~shock_mask]

print(f"\n{'':20s} {'Full Sample':>15s} {'Shock Months':>15s} {'Calm Months':>15s}")
print(f"{'-'*70}")
print(f"{'N months':20s} {len(df):>15d} {len(shock):>15d} {len(calm):>15d}")
print(f"{'Avg REIT excess %':20s} {df['excess_ret'].mean():>15.2f} {shock['excess_ret'].mean():>15.2f} {calm['excess_ret'].mean():>15.2f}")
print(f"{'Std REIT excess %':20s} {df['excess_ret'].std():>15.2f} {shock['excess_ret'].std():>15.2f} {calm['excess_ret'].std():>15.2f}")
print(f"{'Avg oil chg %':20s} {df['oil_chg'].mean():>15.2f} {shock['oil_chg'].mean():>15.2f} {calm['oil_chg'].mean():>15.2f}")
print(f"{'Avg 10Y chg pp':20s} {df['d_t10y'].mean():>15.3f} {shock['d_t10y'].mean():>15.3f} {calm['d_t10y'].mean():>15.3f}")
print(f"{'Avg 3M chg pp':20s} {df['d_t3m'].mean():>15.3f} {shock['d_t3m'].mean():>15.3f} {calm['d_t3m'].mean():>15.3f}")

# Correlation during shocks vs calm
print(f"\n  Correlations:")
print(f"{'':20s} {'Full':>12s} {'Shock':>12s} {'Calm':>12s}")
print(f"{'Oil vs REIT excess':20s} {df['oil_chg'].corr(df['excess_ret']):>12.3f} {shock['oil_chg'].corr(shock['excess_ret']):>12.3f} {calm['oil_chg'].corr(calm['excess_ret']):>12.3f}")
print(f"{'Oil vs 10Y chg':20s} {df['oil_chg'].corr(df['d_t10y']):>12.3f} {shock['oil_chg'].corr(shock['d_t10y']):>12.3f} {calm['oil_chg'].corr(calm['d_t10y']):>12.3f}")
print(f"{'Oil vs 3M chg':20s} {df['oil_chg'].corr(df['d_t3m']):>12.3f} {shock['oil_chg'].corr(shock['d_t3m']):>12.3f} {calm['oil_chg'].corr(calm['d_t3m']):>12.3f}")
print(f"{'10Y chg vs excess':20s} {df['d_t10y'].corr(df['excess_ret']):>12.3f} {shock['d_t10y'].corr(shock['excess_ret']):>12.3f} {calm['d_t10y'].corr(calm['excess_ret']):>12.3f}")

# ================================================================
# Key question: during oil shocks, does oil -> rates -> REITs chain strengthen?
# ================================================================
print(f"\n\n{'='*70}")
print("  BOTTOM LINE")
print(f"{'='*70}")

# Full sample regressions for comparison
X_full = sm.add_constant(df[['oil_chg', 'd_t3m', 'd_t10y']])
m_full = sm.OLS(df['excess_ret'], X_full).fit(cov_type='HC1')

X_shock = sm.add_constant(shock[['oil_chg', 'd_t3m', 'd_t10y']])
m_shock = sm.OLS(shock['excess_ret'], X_shock).fit(cov_type='HC1')

print(f"""
REIT Excess Return ~ Oil + Rate Changes:
                        Full Sample          Shock Months Only
  Oil coefficient:      {m_full.params['oil_chg']:>8.4f} (p={m_full.pvalues['oil_chg']:.3f})    {m_shock.params['oil_chg']:>8.4f} (p={m_shock.pvalues['oil_chg']:.3f})
  10Y chg coefficient:  {m_full.params['d_t10y']:>8.4f} (p={m_full.pvalues['d_t10y']:.3f})    {m_shock.params['d_t10y']:>8.4f} (p={m_shock.pvalues['d_t10y']:.3f})
  R-squared:            {m_full.rsquared:>8.3f}                {m_shock.rsquared:>8.3f}
  N:                    {int(m_full.nobs):>8d}                {int(m_shock.nobs):>8d}
""")

X_full2 = sm.add_constant(df[['oil_chg']])
m_full2 = sm.OLS(df['d_t10y'], X_full2).fit(cov_type='HC1')
X_shock2 = sm.add_constant(shock[['oil_chg']])
m_shock2 = sm.OLS(shock['d_t10y'], X_shock2).fit(cov_type='HC1')

print(f"""10Y Rate Change ~ Oil Change:
                        Full Sample          Shock Months Only
  Oil coefficient:      {m_full2.params['oil_chg']:>8.4f} (p={m_full2.pvalues['oil_chg']:.3f})    {m_shock2.params['oil_chg']:>8.4f} (p={m_shock2.pvalues['oil_chg']:.3f})
  R-squared:            {m_full2.rsquared:>8.3f}                {m_shock2.rsquared:>8.3f}
""")

print("═"*70)
print("  DONE")
print("═"*70)
