"""
Build Excel workbook with:
  Tab 1: Raw monthly data
  Tab 2: Shock Periods (flags each month, lists historical windows)
  Tab 3: Correlations (CORREL formulas)
  Tab 4: Full Sample Regressions (LINEST formulas + cross-check)
  Tab 5: Shock Regressions (all shock definitions side by side)
  Tab 6: Charts
"""

import io
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series, LineChart
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FONT = Font(bold=True, color='FFFFFF', size=11)
HEADER_FILL = PatternFill('solid', fgColor='2F5496')
RED_FILL = PatternFill('solid', fgColor='C00000')
RED_FONT = Font(bold=True, color='FFFFFF', size=11)
LIGHT_FILL = PatternFill('solid', fgColor='D6E4F0')
SHOCK_FILL = PatternFill('solid', fgColor='FCE4EC')
GREEN_FILL = PatternFill('solid', fgColor='E2EFDA')
BOLD = Font(bold=True)
BLUE = Font(color='2F5496')
BLUE_BOLD = Font(bold=True, color='2F5496')
ITALIC_GRAY = Font(italic=True, color='888888')
THIN_BORDER = Border(bottom=Side(style='thin', color='B4C6E7'))


def _header_row(ws, row, headers, fill=HEADER_FILL, font=HEADER_FONT):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = font
        c.fill = fill
        c.alignment = Alignment(horizontal='center', wrap_text=True)


def _auto_width(ws):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        max_len = max((len(str(c.value or '')) for c in col), default=8)
        ws.column_dimensions[letter].width = min(max_len + 3, 26)


def _write_model_row(ws, row, var_name, coef, se, tstat, pval):
    """Write one variable row with coef, SE, t-stat, p-value, significance."""
    sig = '***' if pval < 0.01 else '**' if pval < 0.05 else '*' if pval < 0.1 else ''
    ws.cell(row=row, column=1, value=var_name).font = BOLD
    ws.cell(row=row, column=2, value=round(float(coef), 4)).number_format = '0.0000'
    ws.cell(row=row, column=3, value=round(float(se), 4)).number_format = '0.0000'
    ws.cell(row=row, column=4, value=round(float(tstat), 3)).number_format = '0.000'
    ws.cell(row=row, column=5, value=round(float(pval), 4)).number_format = '0.0000'
    ws.cell(row=row, column=6, value=sig)
    if pval < 0.05:
        ws.cell(row=row, column=6).font = Font(bold=True, color='C00000')
    for c in range(1, 7):
        ws.cell(row=row, column=c).border = THIN_BORDER


def _write_reg_block(ws, start_row, title, model, include_linest=False, linest_base=None):
    """Write a full regression block. Returns next free row."""
    r = start_row
    ws.cell(row=r, column=1, value=title).font = Font(bold=True, size=12, color='2F5496')
    r += 1

    # Summary stats
    ws.cell(row=r, column=1, value='Observations').font = BOLD
    ws.cell(row=r, column=2, value=int(model.nobs))
    r += 1
    ws.cell(row=r, column=1, value='R-squared').font = BOLD
    ws.cell(row=r, column=2, value=round(float(model.rsquared), 4)).number_format = '0.0000'
    if include_linest and linest_base:
        ws.cell(row=r, column=3, value=f'=INDEX({linest_base},3,1)')
        ws.cell(row=r, column=3).number_format = '0.0000'
        ws.cell(row=r, column=3).font = BLUE
    r += 1
    ws.cell(row=r, column=1, value='Adj. R-squared').font = BOLD
    ws.cell(row=r, column=2, value=round(float(model.rsquared_adj), 4)).number_format = '0.0000'
    r += 1
    ws.cell(row=r, column=1, value='F-statistic').font = BOLD
    ws.cell(row=r, column=2, value=round(float(model.fvalue), 3)).number_format = '0.000'
    r += 1
    ws.cell(row=r, column=1, value='F p-value').font = BOLD
    ws.cell(row=r, column=2, value=round(float(model.f_pvalue), 4)).number_format = '0.0000'
    r += 2

    _header_row(ws, r, ['Variable', 'Coefficient', 'Std Error', 't-stat', 'p-value', 'Sig'])
    r += 1

    # LINEST cross-check column headers if needed
    if include_linest and linest_base:
        ws.cell(row=r-1, column=7, value='Coef (formula)').font = BLUE_BOLD
        ws.cell(row=r-1, column=7).fill = LIGHT_FILL

    n_x = len(model.params) - 1  # excluding constant
    for i, var in enumerate(model.params.index):
        _write_model_row(ws, r, var,
                         model.params[var], model.bse[var],
                         model.tvalues[var], model.pvalues[var])
        # LINEST cross-check
        if include_linest and linest_base:
            # LINEST order: xn, xn-1, ..., x1, intercept
            if var == 'const':
                li = n_x + 1
            else:
                # Find position of this var among non-const params (0-indexed)
                non_const = [v for v in model.params.index if v != 'const']
                pos = non_const.index(var)
                li = n_x - pos  # LINEST reverses
            ws.cell(row=r, column=7, value=f'=INDEX({linest_base},1,{li})')
            ws.cell(row=r, column=7).number_format = '0.0000'
            ws.cell(row=r, column=7).font = BLUE
        r += 1

    r += 1
    return r


def build_excel(data, regressions):
    wb = Workbook()
    df = data['df']
    n = len(df)
    last_row = n + 1

    shock_defs = regressions.get('_shock_defs', {})
    historical_shocks = regressions.get('_historical_shocks', {})
    oil_std = regressions.get('_oil_std', df['oil_chg'].std())
    shock_results = regressions.get('_shock_results', {})
    shock_counts = regressions.get('_shock_counts', {})

    # ==================================================================
    # TAB 1: DATA
    # ==================================================================
    ws = wb.active
    ws.title = 'Data'
    ws.sheet_properties.tabColor = '2F5496'

    headers = ['Date', 'Oil Price', 'Oil Chg %', 'REIT Return %',
               'S&P 500 Return %', 'REIT Excess Return %',
               '3M Rate %', '10Y Rate %', 'Term Spread %',
               'Chg in 3M Rate', 'Chg in 10Y Rate']
    _header_row(ws, 1, headers)

    for r, (date, row) in enumerate(df.iterrows(), 2):
        ws.cell(row=r, column=1, value=date.strftime('%Y-%m'))
        ws.cell(row=r, column=2, value=round(row['oil_price'], 2))
        ws.cell(row=r, column=3, value=round(row['oil_chg'], 2))
        ws.cell(row=r, column=4, value=round(row['reit_ret'], 2))
        ws.cell(row=r, column=5, value=round(row['spx_ret'], 2))
        ws.cell(row=r, column=6, value=round(row['excess_ret'], 2))
        ws.cell(row=r, column=7, value=round(row['t3m'], 2))
        ws.cell(row=r, column=8, value=round(row['t10y'], 2))
        ws.cell(row=r, column=9, value=round(row['term_spread'], 2))
        ws.cell(row=r, column=10, value=round(row['d_t3m'], 4))
        ws.cell(row=r, column=11, value=round(row['d_t10y'], 4))

    _auto_width(ws)
    ws.freeze_panes = 'A2'

    def drange(col_letter):
        return f"Data!{col_letter}2:{col_letter}{last_row}"

    # ==================================================================
    # TAB 2: SHOCK PERIODS
    # ==================================================================
    ws2 = wb.create_sheet('Shock Periods')
    ws2.sheet_properties.tabColor = 'C00000'

    ws2.cell(row=1, column=1, value='Oil Shock Period Identification').font = Font(bold=True, size=14, color='2F5496')
    ws2.cell(row=2, column=1, value=f'1 Std Dev of monthly oil change = {oil_std:.1f}%. Months highlighted in pink are shock months.').font = ITALIC_GRAY

    # Monthly shock flags
    r = 4
    flag_headers = ['Date', 'Oil Chg %', '|Oil Chg| > 1 SD', '>1.5 SD', '|Chg| > 10%',
                    'Spike (>10%)', 'Crash (<-10%)', 'Historical Window']
    _header_row(ws2, r, flag_headers)
    r += 1

    shock_1sd = df['oil_chg'].abs() > oil_std
    shock_15sd = df['oil_chg'].abs() > 1.5 * oil_std
    shock_10 = df['oil_chg'].abs() > 10
    spike_10 = df['oil_chg'] > 10
    crash_10 = df['oil_chg'] < -10
    any_hist = shock_defs.get('Any historical window', pd.Series(False, index=df.index))

    for date, row in df.iterrows():
        dt = date.strftime('%Y-%m')
        oil = row['oil_chg']
        is_1sd = abs(oil) > oil_std
        is_15sd = abs(oil) > 1.5 * oil_std
        is_10 = abs(oil) > 10
        is_spike = oil > 10
        is_crash = oil < -10
        is_hist = bool(any_hist.get(date, False))

        ws2.cell(row=r, column=1, value=dt)
        ws2.cell(row=r, column=2, value=round(oil, 2)).number_format = '0.00'
        ws2.cell(row=r, column=3, value='YES' if is_1sd else '')
        ws2.cell(row=r, column=4, value='YES' if is_15sd else '')
        ws2.cell(row=r, column=5, value='YES' if is_10 else '')
        ws2.cell(row=r, column=6, value='YES' if is_spike else '')
        ws2.cell(row=r, column=7, value='YES' if is_crash else '')
        ws2.cell(row=r, column=8, value='YES' if is_hist else '')

        if is_1sd:
            for c in range(1, 9):
                ws2.cell(row=r, column=c).fill = SHOCK_FILL
        r += 1

    # Historical shock windows list
    r += 2
    ws2.cell(row=r, column=1, value='Historical Shock Windows').font = Font(bold=True, size=13, color='2F5496')
    r += 1
    _header_row(ws2, r, ['Period', 'Start', 'End', 'Months', 'Avg Oil Chg %',
                          'Avg REIT Excess %', 'Avg 10Y Chg (pp)'])
    r += 1
    for label, (start, end) in historical_shocks.items():
        mask = (df.index >= start) & (df.index <= end)
        sub = df[mask]
        if len(sub) == 0:
            continue
        ws2.cell(row=r, column=1, value=label)
        ws2.cell(row=r, column=2, value=start)
        ws2.cell(row=r, column=3, value=end)
        ws2.cell(row=r, column=4, value=len(sub))
        ws2.cell(row=r, column=5, value=round(float(sub['oil_chg'].mean()), 1))
        ws2.cell(row=r, column=6, value=round(float(sub['excess_ret'].mean()), 2))
        ws2.cell(row=r, column=7, value=round(float(sub['d_t10y'].mean()), 3))
        for c in range(1, 8):
            ws2.cell(row=r, column=c).border = THIN_BORDER
        r += 1

    # Shock definition counts
    r += 2
    ws2.cell(row=r, column=1, value='Shock Definition Summary').font = Font(bold=True, size=13, color='2F5496')
    r += 1
    _header_row(ws2, r, ['Definition', 'Months', '% of Sample'])
    r += 1
    for label, count in shock_counts.items():
        ws2.cell(row=r, column=1, value=label)
        ws2.cell(row=r, column=2, value=count)
        ws2.cell(row=r, column=3, value=round(count / n * 100, 1))
        ws2.cell(row=r, column=3).number_format = '0.0'
        for c in range(1, 4):
            ws2.cell(row=r, column=c).border = THIN_BORDER
        r += 1

    _auto_width(ws2)

    # ==================================================================
    # TAB 3: CORRELATIONS (CORREL formulas)
    # ==================================================================
    ws3 = wb.create_sheet('Correlations')
    ws3.sheet_properties.tabColor = '548235'

    ws3.cell(row=1, column=1, value='Correlation Matrix (Full Sample)').font = Font(bold=True, size=14, color='2F5496')
    ws3.cell(row=2, column=1, value='All values are CORREL() formulas referencing the Data tab').font = ITALIC_GRAY

    corr_vars = [
        ('Oil Chg %', 'C'), ('REIT Excess Ret %', 'F'), ('3M Rate %', 'G'),
        ('10Y Rate %', 'H'), ('Term Spread %', 'I'), ('Chg in 3M', 'J'), ('Chg in 10Y', 'K'),
    ]

    r0 = 4
    _header_row(ws3, r0, [''] + [v[0] for v in corr_vars])
    for ri, (row_name, row_col) in enumerate(corr_vars):
        r = r0 + 1 + ri
        ws3.cell(row=r, column=1, value=row_name).font = BOLD
        for ci, (_, col_col) in enumerate(corr_vars, 2):
            if row_col == col_col:
                ws3.cell(row=r, column=ci, value=1.0)
            else:
                ws3.cell(row=r, column=ci, value=f'=CORREL({drange(row_col)},{drange(col_col)})')
            ws3.cell(row=r, column=ci).number_format = '0.000'

    _auto_width(ws3)

    # ==================================================================
    # TAB 4: FULL SAMPLE REGRESSIONS
    # ==================================================================
    ws4 = wb.create_sheet('Full Sample Regressions')
    ws4.sheet_properties.tabColor = 'BF8F00'

    r = 1
    ws4.cell(row=r, column=1, value='Full Sample Regression Results').font = Font(bold=True, size=14, color='2F5496')
    r += 1
    ws4.cell(row=r, column=1, value='Black = values. Blue = LINEST formula cross-check. *** p<0.01  ** p<0.05  * p<0.1').font = ITALIC_GRAY
    r += 2

    # LINEST base for cross-check (non-adjacent cols C,J,K need CHOOSE)
    linest_reit = f'LINEST({drange("F")},CHOOSE({{1,2,3}},{drange("C")},{drange("J")},{drange("K")}),TRUE,TRUE)'
    r = _write_reg_block(ws4, r, 'REIT Excess Return ~ Oil + Rate Changes',
                         regressions['reit_m1_levels_chg'],
                         include_linest=True, linest_base=linest_reit)

    linest_10y = f'LINEST({drange("K")},{drange("C")},TRUE,TRUE)'
    r = _write_reg_block(ws4, r, '10Y Rate Change ~ Oil Change',
                         regressions['t10y_on_oil_ols'],
                         include_linest=True, linest_base=linest_10y)

    linest_3m = f'LINEST({drange("J")},{drange("C")},TRUE,TRUE)'
    r = _write_reg_block(ws4, r, '3M Rate Change ~ Oil Change',
                         regressions['t3m_on_oil_ols'],
                         include_linest=True, linest_base=linest_3m)

    linest_oil = f'LINEST({drange("C")},{drange("J")}:{drange("K").split("!")[1]},TRUE,TRUE)'
    # Adjacent cols J,K so simpler
    linest_oil = f'LINEST({drange("C")},Data!J2:K{last_row},TRUE,TRUE)'
    r = _write_reg_block(ws4, r, 'Oil Change ~ Rate Changes',
                         regressions['oil_rates_changes_ols'],
                         include_linest=True, linest_base=linest_oil)

    # Additional HC1 models (no LINEST cross-check)
    r = _write_reg_block(ws4, r, 'REIT Excess ~ Oil + Rate Changes (Winsorized, Robust SE)',
                         regressions['reit_m2'])
    r = _write_reg_block(ws4, r, 'REIT Excess ~ Asymmetric Oil + Rate Changes (Winsorized)',
                         regressions['reit_m3'])

    _auto_width(ws4)

    # ==================================================================
    # TAB 5: SHOCK REGRESSIONS
    # ==================================================================
    ws5 = wb.create_sheet('Shock Regressions')
    ws5.sheet_properties.tabColor = 'C00000'

    r = 1
    ws5.cell(row=r, column=1, value='Regressions During Oil Shock Periods').font = Font(bold=True, size=14, color='C00000')
    r += 1
    ws5.cell(row=r, column=1, value='Same models as Full Sample tab, but restricted to shock months only. *** p<0.01  ** p<0.05  * p<0.1').font = ITALIC_GRAY
    r += 2

    # Comparison table: all shock defs side by side for the REIT regression
    ws5.cell(row=r, column=1, value='REIT Excess Return ~ Oil + Rate Changes: Shock vs Full Sample').font = Font(bold=True, size=13, color='2F5496')
    r += 1

    comp_headers = ['Shock Definition', 'N', 'R-sq',
                    'Oil Coef', 'Oil t-stat', 'Oil p-val',
                    'd_10Y Coef', 'd_10Y t-stat', 'd_10Y p-val',
                    'd_3M Coef', 'd_3M t-stat', 'd_3M p-val']
    _header_row(ws5, r, comp_headers)
    r += 1

    # Full sample first
    m_full = regressions['reit_m1_levels_chg']
    ws5.cell(row=r, column=1, value='FULL SAMPLE').font = Font(bold=True, color='2F5496')
    ws5.cell(row=r, column=2, value=int(m_full.nobs))
    ws5.cell(row=r, column=3, value=round(float(m_full.rsquared), 4)).number_format = '0.0000'
    for vi, var in enumerate(['oil_chg', 'd_t10y', 'd_t3m']):
        base_col = 4 + vi * 3
        ws5.cell(row=r, column=base_col, value=round(float(m_full.params[var]), 4)).number_format = '0.0000'
        ws5.cell(row=r, column=base_col+1, value=round(float(m_full.tvalues[var]), 3)).number_format = '0.000'
        ws5.cell(row=r, column=base_col+2, value=round(float(m_full.pvalues[var]), 4)).number_format = '0.0000'
    for c in range(1, 13):
        ws5.cell(row=r, column=c).border = THIN_BORDER
        ws5.cell(row=r, column=c).fill = LIGHT_FILL
    r += 1

    # Each shock definition
    ordered_labels = [
        '1 SD (|oil| > 1 std dev)', '1.5 SD (|oil| > 1.5 std dev)',
        'Top/Bottom 10%', 'Big move (|chg| > 10%)',
        'Oil spike (>10%)', 'Oil crash (<-10%)',
        'Any historical window', 'Calm months (|oil| < 1 SD)',
    ]
    for label in ordered_labels:
        sr = shock_results.get(label)
        if sr is None or 'reit' not in sr:
            ws5.cell(row=r, column=1, value=label)
            ws5.cell(row=r, column=2, value=shock_counts.get(label, '<10'))
            ws5.cell(row=r, column=3, value='N/A (too few obs)')
            r += 1
            continue

        m = sr['reit']
        ws5.cell(row=r, column=1, value=label).font = BOLD
        ws5.cell(row=r, column=2, value=int(m.nobs))
        ws5.cell(row=r, column=3, value=round(float(m.rsquared), 4)).number_format = '0.0000'
        for vi, var in enumerate(['oil_chg', 'd_t10y', 'd_t3m']):
            base_col = 4 + vi * 3
            ws5.cell(row=r, column=base_col, value=round(float(m.params[var]), 4)).number_format = '0.0000'
            ws5.cell(row=r, column=base_col+1, value=round(float(m.tvalues[var]), 3)).number_format = '0.000'
            pval = float(m.pvalues[var])
            cell_p = ws5.cell(row=r, column=base_col+2, value=round(pval, 4))
            cell_p.number_format = '0.0000'
            if pval < 0.05:
                cell_p.font = Font(bold=True, color='C00000')
        for c in range(1, 13):
            ws5.cell(row=r, column=c).border = THIN_BORDER
        if label == 'Calm months (|oil| < 1 SD)':
            for c in range(1, 13):
                ws5.cell(row=r, column=c).fill = GREEN_FILL
        r += 1

    # Same comparison for Oil -> 10Y regression
    r += 2
    ws5.cell(row=r, column=1, value='10Y Rate Change ~ Oil Change: Shock vs Full Sample').font = Font(bold=True, size=13, color='2F5496')
    r += 1
    comp_h2 = ['Shock Definition', 'N', 'R-sq', 'Oil Coef', 'Oil t-stat', 'Oil p-val']
    _header_row(ws5, r, comp_h2)
    r += 1

    m_full2 = regressions['t10y_on_oil_ols']
    ws5.cell(row=r, column=1, value='FULL SAMPLE').font = Font(bold=True, color='2F5496')
    ws5.cell(row=r, column=2, value=int(m_full2.nobs))
    ws5.cell(row=r, column=3, value=round(float(m_full2.rsquared), 4)).number_format = '0.0000'
    ws5.cell(row=r, column=4, value=round(float(m_full2.params['oil_chg']), 4)).number_format = '0.0000'
    ws5.cell(row=r, column=5, value=round(float(m_full2.tvalues['oil_chg']), 3)).number_format = '0.000'
    ws5.cell(row=r, column=6, value=round(float(m_full2.pvalues['oil_chg']), 4)).number_format = '0.0000'
    for c in range(1, 7):
        ws5.cell(row=r, column=c).border = THIN_BORDER
        ws5.cell(row=r, column=c).fill = LIGHT_FILL
    r += 1

    for label in ordered_labels:
        sr = shock_results.get(label)
        if sr is None or 't10y' not in sr:
            ws5.cell(row=r, column=1, value=label)
            ws5.cell(row=r, column=2, value=shock_counts.get(label, '<10'))
            ws5.cell(row=r, column=3, value='N/A')
            r += 1
            continue
        m = sr['t10y']
        ws5.cell(row=r, column=1, value=label).font = BOLD
        ws5.cell(row=r, column=2, value=int(m.nobs))
        ws5.cell(row=r, column=3, value=round(float(m.rsquared), 4)).number_format = '0.0000'
        ws5.cell(row=r, column=4, value=round(float(m.params['oil_chg']), 4)).number_format = '0.0000'
        ws5.cell(row=r, column=5, value=round(float(m.tvalues['oil_chg']), 3)).number_format = '0.000'
        pval = float(m.pvalues['oil_chg'])
        cell_p = ws5.cell(row=r, column=6, value=round(pval, 4))
        cell_p.number_format = '0.0000'
        if pval < 0.05:
            cell_p.font = Font(bold=True, color='C00000')
        for c in range(1, 7):
            ws5.cell(row=r, column=c).border = THIN_BORDER
        if label == 'Calm months (|oil| < 1 SD)':
            for c in range(1, 7):
                ws5.cell(row=r, column=c).fill = GREEN_FILL
        r += 1

    # Full detail blocks for key shock definitions
    r += 2
    ws5.cell(row=r, column=1, value='Detailed Regression Output by Shock Definition').font = Font(bold=True, size=14, color='2F5496')
    r += 2

    for label in ['1 SD (|oil| > 1 std dev)', '1.5 SD (|oil| > 1.5 std dev)',
                   'Oil crash (<-10%)', 'Any historical window',
                   'Calm months (|oil| < 1 SD)']:
        sr = shock_results.get(label)
        if sr is None:
            continue
        for reg_name, model_key, desc in [
            ('reit', 'reit', 'REIT Excess ~ Oil + Rate Changes'),
            ('t10y', 't10y', '10Y Change ~ Oil Change'),
            ('t3m', 't3m', '3M Change ~ Oil Change'),
        ]:
            if model_key in sr:
                r = _write_reg_block(ws5, r, f'{desc} [{label}]', sr[model_key])

    _auto_width(ws5)

    # ==================================================================
    # TAB 6: KEY FINDINGS
    # ==================================================================
    wsf = wb.create_sheet('Key Findings')
    wsf.sheet_properties.tabColor = '548235'
    wsf.column_dimensions['A'].width = 4
    wsf.column_dimensions['B'].width = 90

    r = 1
    wsf.cell(row=r, column=2, value='Key Findings: Oil, REITs & Interest Rates').font = Font(bold=True, size=16, color='2F5496')
    r += 2

    findings = [
        ("1. Oil doesn't directly move REITs vs S&P -- even during shocks.",
         "Across every definition of 'oil shock' (1 SD moves, >10% swings, historical crisis windows), "
         "the oil coefficient on REIT excess returns is statistically insignificant. Oil spikes and crashes "
         "hit REITs and the S&P roughly equally."),

        ("2. Long-term rates are what actually drive the wedge.",
         "A 1 percentage point rise in the 10Y yield in a given month is associated with ~5% REIT "
         "underperformance vs the S&P (p<0.001). REITs behave like long-duration bonds."),

        ("3. Oil's effect on rates gets stronger during shocks.",
         "Full sample: oil->10Y R-sq = 8.3%. Shock months (1 SD): R-sq = 18.1%. "
         "Extreme shocks (1.5 SD): R-sq = 32.1%. The bigger the oil move, the more it pushes the "
         "10-year rate -- likely through inflation expectations."),

        ("4. Oil crashes and spikes work differently.",
         "Oil crashes (<-10%): The 10Y rate effect on REITs becomes highly significant (p<0.01). "
         "Oil crashes pull rates down, which helps REITs relative to the S&P. The oil->3M rate "
         "relationship also strengthens (p=0.02), suggesting the Fed responds to oil-driven deflation risk.\n\n"
         "Oil spikes (>10%): Noisy -- nothing is significant. REITs underperform by -2.3% on average "
         "during spike months, but the regression can't attribute it cleanly to oil or rates."),

        ("5. The transmission chain is indirect.",
         "Oil shock -> inflation expectations -> rates move -> REITs react to rates. "
         "The direct oil->REIT channel is basically zero. The oil->rates->REITs chain is real "
         "but only explains ~6% of monthly variation. Most REIT vs S&P performance comes from "
         "other factors (sector rotation, cap rates, property fundamentals)."),
    ]

    for title, body in findings:
        wsf.cell(row=r, column=2, value=title).font = Font(bold=True, size=12, color='2F5496')
        r += 1
        wsf.cell(row=r, column=2, value=body).alignment = Alignment(wrap_text=True)
        # Auto-height: roughly 1 row per 90 chars
        lines = max(len(body) // 85 + body.count('\n') + 1, 2)
        wsf.row_dimensions[r].height = lines * 16
        r += 2

    r += 1
    wsf.cell(row=r, column=2, value='Supporting Evidence (from Shock Regressions tab)').font = Font(bold=True, size=13, color='2F5496')
    r += 1

    evidence = [
        ['Metric', 'Full Sample', 'Shock (1 SD)', 'Extreme (1.5 SD)', 'Oil Crash (<-10%)'],
    ]
    # Pull actual numbers
    m_full = regressions['reit_m1_levels_chg']
    sr_1sd = shock_results.get('1 SD (|oil| > 1 std dev)', {})
    sr_15sd = shock_results.get('1.5 SD (|oil| > 1.5 std dev)', {})
    sr_crash = shock_results.get('Oil crash (<-10%)', {})

    def _safe(sr, key, attr, var=None):
        m = sr.get(key)
        if m is None:
            return 'N/A'
        if var:
            return round(float(getattr(m, attr)[var]), 4)
        return round(float(getattr(m, attr)), 4)

    evidence.append(['REIT reg: Oil coef',
                     round(float(m_full.params['oil_chg']), 4),
                     _safe(sr_1sd, 'reit', 'params', 'oil_chg'),
                     _safe(sr_15sd, 'reit', 'params', 'oil_chg'),
                     _safe(sr_crash, 'reit', 'params', 'oil_chg')])
    evidence.append(['REIT reg: Oil p-value',
                     round(float(m_full.pvalues['oil_chg']), 4),
                     _safe(sr_1sd, 'reit', 'pvalues', 'oil_chg'),
                     _safe(sr_15sd, 'reit', 'pvalues', 'oil_chg'),
                     _safe(sr_crash, 'reit', 'pvalues', 'oil_chg')])
    evidence.append(['REIT reg: 10Y coef',
                     round(float(m_full.params['d_t10y']), 4),
                     _safe(sr_1sd, 'reit', 'params', 'd_t10y'),
                     _safe(sr_15sd, 'reit', 'params', 'd_t10y'),
                     _safe(sr_crash, 'reit', 'params', 'd_t10y')])
    evidence.append(['REIT reg: 10Y p-value',
                     round(float(m_full.pvalues['d_t10y']), 4),
                     _safe(sr_1sd, 'reit', 'pvalues', 'd_t10y'),
                     _safe(sr_15sd, 'reit', 'pvalues', 'd_t10y'),
                     _safe(sr_crash, 'reit', 'pvalues', 'd_t10y')])

    m_full_10y = regressions['t10y_on_oil_ols']
    evidence.append(['Oil->10Y: R-squared',
                     round(float(m_full_10y.rsquared), 4),
                     _safe(sr_1sd, 't10y', 'rsquared'),
                     _safe(sr_15sd, 't10y', 'rsquared'),
                     _safe(sr_crash, 't10y', 'rsquared')])
    evidence.append(['Oil->10Y: Oil p-value',
                     round(float(m_full_10y.pvalues['oil_chg']), 4),
                     _safe(sr_1sd, 't10y', 'pvalues', 'oil_chg'),
                     _safe(sr_15sd, 't10y', 'pvalues', 'oil_chg'),
                     _safe(sr_crash, 't10y', 'pvalues', 'oil_chg')])
    evidence.append(['N months',
                     int(m_full.nobs),
                     shock_counts.get('1 SD (|oil| > 1 std dev)', 0),
                     shock_counts.get('1.5 SD (|oil| > 1.5 std dev)', 0),
                     shock_counts.get('Oil crash (<-10%)', 0)])

    _header_row(wsf, r, [''] + evidence[0])
    r += 1
    for row_data in evidence[1:]:
        wsf.cell(row=r, column=2, value=row_data[0]).font = BOLD
        for ci, val in enumerate(row_data[1:], 3):
            cell = wsf.cell(row=r, column=ci, value=val)
            if isinstance(val, float):
                cell.number_format = '0.0000'
        for c in range(2, 7):
            wsf.cell(row=r, column=c).border = THIN_BORDER
        r += 1

    # ==================================================================
    # TAB 7: SOURCES
    # ==================================================================
    wss = wb.create_sheet('Sources')
    wss.sheet_properties.tabColor = '333333'
    wss.column_dimensions['A'].width = 4
    wss.column_dimensions['B'].width = 30
    wss.column_dimensions['C'].width = 65
    wss.column_dimensions['D'].width = 55

    r = 1
    wss.cell(row=r, column=2, value='Data Sources').font = Font(bold=True, size=16, color='2F5496')
    r += 2
    _header_row(wss, r, ['', 'Source', 'URL', 'Used For'])
    r += 1

    sources = [
        ('FRED: DCOILWTICO',
         'https://fred.stlouisfed.org/series/DCOILWTICO',
         'WTI Crude Oil spot price (daily). Resampled to monthly average, then converted to % change for the oil shock variable.'),
        ('FRED: DTB3',
         'https://fred.stlouisfed.org/series/DTB3',
         '3-Month Treasury Bill rate (daily). Resampled to monthly average. Used as the short-term interest rate variable (level and month-over-month change).'),
        ('FRED: DGS10',
         'https://fred.stlouisfed.org/series/DGS10',
         '10-Year Treasury Constant Maturity rate (daily). Resampled to monthly average. Used as the long-term interest rate variable (level and month-over-month change).'),
        ('Yahoo Finance: ^RMZ',
         'https://finance.yahoo.com/quote/%5ERMZ/',
         'MSCI US REIT Index (daily close, Jun 1995 - Sep 2021). Used as the REIT benchmark for the earlier portion of the sample period.'),
        ('Yahoo Finance: IYR',
         'https://finance.yahoo.com/quote/IYR/',
         'iShares U.S. Real Estate ETF (daily close, Jun 2000 - present). Spliced with ^RMZ to create a continuous REIT return series through 2025.'),
        ('Yahoo Finance: ^GSPC',
         'https://finance.yahoo.com/quote/%5EGSPC/',
         'S&P 500 Index (daily close). Resampled to month-end values, then converted to monthly % return. Used as the broad equity benchmark.'),
    ]

    for source, url, desc in sources:
        wss.cell(row=r, column=2, value=source).font = BOLD
        wss.cell(row=r, column=3, value=url).font = Font(color='0563C1', underline='single')
        wss.hyperlink = url
        wss.cell(row=r, column=4, value=desc).alignment = Alignment(wrap_text=True)
        wss.row_dimensions[r].height = max(32, len(desc) // 55 * 16 + 16)
        for c in range(2, 5):
            wss.cell(row=r, column=c).border = THIN_BORDER
        r += 1

    r += 2
    wss.cell(row=r, column=2, value='Methodology Notes').font = Font(bold=True, size=13, color='2F5496')
    r += 1
    notes = [
        ('REIT Index Splice',
         '^RMZ covers Jun 1995 - Sep 2021. IYR covers Jun 2000 - present. We use ^RMZ for months before IYR began, then IYR from Jun 2000 onward, creating a continuous series from 1995 to 2025.'),
        ('Monthly Returns',
         'All return series use month-end closing prices. Monthly return = (close_t / close_{t-1} - 1) * 100.'),
        ('Excess Return',
         'REIT excess return = REIT monthly return - S&P 500 monthly return. Positive = REITs outperformed.'),
        ('Rate Changes',
         'Month-over-month change in the monthly average rate level (in percentage points). Used in regressions instead of rate levels to avoid spurious correlation.'),
        ('Robust Standard Errors',
         'Website regressions use HC1 (heteroscedasticity-consistent) standard errors. Excel LINEST formulas use classical OLS standard errors (cross-check column).'),
        ('Winsorization',
         'Some models trim excess returns and oil changes at the 1st/99th percentile to reduce the influence of extreme outliers (e.g., COVID March 2020, GFC).'),
    ]
    for title, desc in notes:
        wss.cell(row=r, column=2, value=title).font = BOLD
        wss.cell(row=r, column=3, value=desc).alignment = Alignment(wrap_text=True)
        wss.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        wss.row_dimensions[r].height = max(32, len(desc) // 90 * 16 + 16)
        for c in range(2, 5):
            wss.cell(row=r, column=c).border = THIN_BORDER
        r += 1

    # ==================================================================
    # TAB 8: CHARTS
    # ==================================================================
    ws6 = wb.create_sheet('Charts')
    ws6.sheet_properties.tabColor = '7030A0'

    # Chart 1: Oil vs REIT excess
    c1 = ScatterChart()
    c1.title = 'Oil Price Swings vs REIT Outperformance'
    c1.x_axis.title = 'Monthly Oil Price Move (%)'
    c1.y_axis.title = 'REIT vs S&P (%)'
    c1.width = 20
    c1.height = 14
    s1 = Series(Reference(ws, min_col=6, min_row=2, max_row=last_row),
                Reference(ws, min_col=3, min_row=2, max_row=last_row), title='Monthly')
    s1.graphicalProperties.noFill = True
    c1.series.append(s1)
    ws6.add_chart(c1, 'A1')

    # Chart 2: 10Y change vs REIT excess
    c2 = ScatterChart()
    c2.title = 'When Long-Term Rates Rise, REITs Underperform'
    c2.x_axis.title = '10Y Rate Change (pp)'
    c2.y_axis.title = 'REIT vs S&P (%)'
    c2.width = 20
    c2.height = 14
    s2 = Series(Reference(ws, min_col=6, min_row=2, max_row=last_row),
                Reference(ws, min_col=11, min_row=2, max_row=last_row), title='Monthly')
    s2.graphicalProperties.noFill = True
    c2.series.append(s2)
    ws6.add_chart(c2, 'L1')

    # Chart 3: Oil vs 10Y change
    c3 = ScatterChart()
    c3.title = 'Oil and Long-Term Rates Move Together'
    c3.x_axis.title = 'Monthly Oil Price Move (%)'
    c3.y_axis.title = '10Y Rate Change (pp)'
    c3.width = 20
    c3.height = 14
    s3 = Series(Reference(ws, min_col=11, min_row=2, max_row=last_row),
                Reference(ws, min_col=3, min_row=2, max_row=last_row), title='Monthly')
    s3.graphicalProperties.noFill = True
    c3.series.append(s3)
    ws6.add_chart(c3, 'A18')

    # Chart 4: Oil + 10Y over time
    c4 = LineChart()
    c4.title = 'Oil Price and 10Y Rate Over Time'
    c4.y_axis.title = 'Oil Price ($/bbl)'
    c4.width = 20
    c4.height = 14
    cats = Reference(ws, min_col=1, min_row=2, max_row=last_row)
    c4.add_data(Reference(ws, min_col=2, min_row=1, max_row=last_row), titles_from_data=True)
    c4.set_categories(cats)
    c4.series[0].graphicalProperties.line.width = 15000
    c4b = LineChart()
    c4b.y_axis.title = '10Y Rate (%)'
    c4b.add_data(Reference(ws, min_col=8, min_row=1, max_row=last_row), titles_from_data=True)
    c4b.set_categories(cats)
    c4b.y_axis.axId = 200
    c4b.series[0].graphicalProperties.line.width = 15000
    c4.y_axis.crosses = 'min'
    c4 += c4b
    ws6.add_chart(c4, 'L18')

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
