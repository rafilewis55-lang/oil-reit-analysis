"""
Build Excel workbook with:
  Tab 1: Raw monthly data
  Tab 2: Correlation matrix (CORREL formulas referencing Tab 1)
  Tab 3: Regressions (INDEX/LINEST formulas referencing Tab 1) + Python cross-check
  Tab 4: Charts (simplified)
"""

import io
import numpy as np
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series, LineChart
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FONT = Font(bold=True, color='FFFFFF', size=11)
HEADER_FILL = PatternFill('solid', fgColor='2F5496')
LIGHT_FILL = PatternFill('solid', fgColor='D6E4F0')
GREEN_FILL = PatternFill('solid', fgColor='E2EFDA')
BOLD = Font(bold=True)
BLUE = Font(color='2F5496')
BLUE_BOLD = Font(bold=True, color='2F5496')
ITALIC_GRAY = Font(italic=True, color='888888')
THIN_BORDER = Border(bottom=Side(style='thin', color='B4C6E7'))


def _header_row(ws, row, headers):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal='center', wrap_text=True)


def _auto_width(ws):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        max_len = max((len(str(c.value or '')) for c in col), default=8)
        ws.column_dimensions[letter].width = min(max_len + 3, 24)


def build_excel(data, regressions):
    wb = Workbook()
    df = data['df']
    n = len(df)
    last_row = n + 1  # data starts row 2

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

    # Helper: data range strings referencing the Data tab
    def drange(col_letter):
        return f"Data!{col_letter}2:{col_letter}{last_row}"

    # Column mapping for formulas:
    # B=Oil Price, C=Oil Chg, D=REIT Ret, E=SPX Ret, F=Excess Ret,
    # G=3M Rate, H=10Y Rate, I=Spread, J=d_3M, K=d_10Y

    # ==================================================================
    # TAB 2: CORRELATIONS (all formulas)
    # ==================================================================
    ws2 = wb.create_sheet('Correlations')
    ws2.sheet_properties.tabColor = '548235'

    ws2.cell(row=1, column=1, value='Correlation Matrix').font = Font(bold=True, size=14, color='2F5496')
    ws2.cell(row=2, column=1, value='All values are CORREL() formulas referencing the Data tab').font = ITALIC_GRAY

    corr_vars = [
        ('Oil Chg %', 'C'),
        ('REIT Excess Ret %', 'F'),
        ('3M Rate %', 'G'),
        ('10Y Rate %', 'H'),
        ('Term Spread %', 'I'),
        ('Chg in 3M', 'J'),
        ('Chg in 10Y', 'K'),
    ]

    r0 = 4
    # Header row
    for ci, (name, _) in enumerate(corr_vars, 2):
        ws2.cell(row=r0, column=ci, value=name).font = BOLD
    _header_row(ws2, r0, [''] + [v[0] for v in corr_vars])

    # Fill correlation formulas
    for ri, (row_name, row_col) in enumerate(corr_vars):
        r = r0 + 1 + ri
        ws2.cell(row=r, column=1, value=row_name).font = BOLD
        for ci, (col_name, col_col) in enumerate(corr_vars, 2):
            if row_col == col_col:
                ws2.cell(row=r, column=ci, value=1.0)
            else:
                formula = f'=CORREL({drange(row_col)},{drange(col_col)})'
                ws2.cell(row=r, column=ci, value=formula)
            ws2.cell(row=r, column=ci).number_format = '0.000'

    _auto_width(ws2)

    # ==================================================================
    # TAB 3: REGRESSIONS (LINEST formulas + Python cross-check)
    # ==================================================================
    ws3 = wb.create_sheet('Regressions')
    ws3.sheet_properties.tabColor = 'BF8F00'

    r = 1
    ws3.cell(row=r, column=1, value='Regression Analysis').font = Font(bold=True, size=14, color='2F5496')
    r += 1
    ws3.cell(row=r, column=1, value='Black text = Excel formulas (LINEST) pulling from the Data tab.  Blue text = cross-check values.  Match? column verifies they agree.').font = ITALIC_GRAY
    r += 2

    def write_linest_regression(ws, start_row, title, y_col, x_cols, x_names, model):
        """
        Write a regression using INDEX(LINEST()) formulas.
        LINEST returns coefficients in REVERSE order: last X first, then ..., intercept.
        """
        r = start_row
        ws.cell(row=r, column=1, value=title).font = Font(bold=True, size=12, color='2F5496')
        r += 1

        y_range = drange(y_col)
        if len(x_cols) == 1:
            x_range = drange(x_cols[0])
        else:
            x_range = ','.join(drange(c) for c in x_cols)
            # For multi-column LINEST, we need to use CHOOSE or concatenate columns
            # Actually LINEST accepts a multi-column array: Data!C2:C367,Data!J2:J367 won't work
            # We need: Data!J2:K367 if columns are adjacent, or use CHOOSE
            # Let's check if columns are adjacent
            col_nums = [ord(c) - ord('A') + 1 for c in x_cols]
            if col_nums == list(range(col_nums[0], col_nums[0] + len(col_nums))):
                # Adjacent columns
                x_range = f"Data!{x_cols[0]}2:{x_cols[-1]}{last_row}"
            else:
                # Non-adjacent: use CHOOSE trick
                choose_parts = ','.join(f'{drange(c)}' for c in x_cols)
                n_x = len(x_cols)
                choose_seq = ','.join(str(i+1) for i in range(n_x))
                x_range = f'CHOOSE({{{choose_seq}}},{choose_parts})'

        linest_base = f'LINEST({y_range},{x_range},TRUE,TRUE)'

        # LINEST output layout (nX+1 columns, 5 rows):
        # Row 1: coef_Xn, coef_Xn-1, ..., coef_X1, intercept
        # Row 2: se_Xn, se_Xn-1, ..., se_X1, se_intercept
        # Row 3: R², se_y
        # Row 4: F-stat, df
        # Row 5: SS_reg, SS_resid

        n_x = len(x_cols)
        all_names = list(reversed(x_names)) + ['Intercept']  # LINEST order

        # Black = Excel formula, Blue = Python-computed cross-check
        # Header
        _header_row(ws, r, ['', 'Value', 'Cross-Check', 'Match?'])
        r += 1

        # R-squared
        ws.cell(row=r, column=1, value='R-squared').font = BOLD
        ws.cell(row=r, column=2, value=f'=INDEX({linest_base},3,1)')
        ws.cell(row=r, column=2).number_format = '0.0000'
        c3 = ws.cell(row=r, column=3, value=round(float(model.rsquared), 4))
        c3.font = BLUE
        c3.number_format = '0.0000'
        ws.cell(row=r, column=4, value=f'=ROUND(B{r},3)=ROUND(C{r},3)')
        ws.cell(row=r, column=4).font = ITALIC_GRAY
        r += 1

        # F-statistic
        ws.cell(row=r, column=1, value='F-statistic').font = BOLD
        ws.cell(row=r, column=2, value=f'=INDEX({linest_base},4,1)')
        ws.cell(row=r, column=2).number_format = '0.000'
        c3 = ws.cell(row=r, column=3, value=round(float(model.fvalue), 3))
        c3.font = BLUE
        c3.number_format = '0.000'
        ws.cell(row=r, column=4, value=f'=ROUND(B{r},2)=ROUND(C{r},2)')
        ws.cell(row=r, column=4).font = ITALIC_GRAY
        r += 1

        # Observations
        ws.cell(row=r, column=1, value='Observations').font = BOLD
        ws.cell(row=r, column=2, value=f'=INDEX({linest_base},4,2)+{n_x}+1')
        ws.cell(row=r, column=3, value=int(model.nobs)).font = BLUE
        r += 1

        r += 1
        _header_row(ws, r, ['Variable', 'Coefficient', 'Std Error',
                            'Coefficient', 'Std Error', 'Match?'])
        r += 1

        # Coefficients in LINEST order (reversed X, then intercept)
        python_params = list(model.params)
        python_bse = list(model.bse)
        linest_param_order = list(reversed(python_params[1:])) + [python_params[0]]
        linest_bse_order = list(reversed(python_bse[1:])) + [python_bse[0]]
        linest_name_order = list(reversed(list(model.params.index[1:]))) + ['Intercept']

        for i, (name, py_coef, py_se) in enumerate(zip(linest_name_order, linest_param_order, linest_bse_order)):
            col_idx = i + 1
            ws.cell(row=r, column=1, value=name).font = BOLD
            # Black = formula
            ws.cell(row=r, column=2, value=f'=INDEX({linest_base},1,{col_idx})')
            ws.cell(row=r, column=2).number_format = '0.0000'
            ws.cell(row=r, column=3, value=f'=INDEX({linest_base},2,{col_idx})')
            ws.cell(row=r, column=3).number_format = '0.0000'
            # Blue = python
            c4 = ws.cell(row=r, column=4, value=round(float(py_coef), 4))
            c4.font = BLUE
            c4.number_format = '0.0000'
            c5 = ws.cell(row=r, column=5, value=round(float(py_se), 4))
            c5.font = BLUE
            c5.number_format = '0.0000'
            ws.cell(row=r, column=6, value=f'=ROUND(B{r},3)=ROUND(D{r},3)')
            ws.cell(row=r, column=6).font = ITALIC_GRAY
            for c in range(1, 7):
                ws.cell(row=r, column=c).border = THIN_BORDER
            r += 1

        return r + 2

    # --- Regression 1: REIT Excess Return ~ Oil Chg + d_3M + d_10Y ---
    r = write_linest_regression(
        ws3, r,
        'How Oil & Rate Changes Affect REIT vs S&P Performance',
        'F',  # Y = Excess Return
        ['C', 'J', 'K'],  # X = Oil Chg, d_3M, d_10Y
        ['Oil Chg %', 'Chg in 3M Rate', 'Chg in 10Y Rate'],
        regressions['reit_m1_levels_chg']  # we'll create this
    )

    # --- Regression 2: 10Y Change ~ Oil Change ---
    r = write_linest_regression(
        ws3, r,
        'Does Oil Move Long-Term Interest Rates?',
        'K',  # Y = d_10Y
        ['C'],  # X = Oil Chg
        ['Oil Chg %'],
        regressions['t10y_on_oil_ols']
    )

    # --- Regression 3: 3M Change ~ Oil Change ---
    r = write_linest_regression(
        ws3, r,
        'Does Oil Move Short-Term Interest Rates?',
        'J',  # Y = d_3M
        ['C'],  # X = Oil Chg
        ['Oil Chg %'],
        regressions['t3m_on_oil_ols']
    )

    # --- Regression 4: Oil ~ Rate Changes ---
    r = write_linest_regression(
        ws3, r,
        'Do Rate Changes Predict Oil Moves?',
        'C',  # Y = Oil Chg
        ['J', 'K'],  # X = d_3M, d_10Y
        ['Chg in 3M Rate', 'Chg in 10Y Rate'],
        regressions['oil_rates_changes_ols']
    )

    # Interpretation block
    r += 1
    ws3.cell(row=r, column=1, value='How to Read These Results').font = Font(bold=True, size=12, color='2F5496')
    r += 1
    notes = [
        'R-squared: What % of the ups and downs in Y are explained by the X variables (0-100%).',
        'Coefficient: How much Y moves when X goes up by 1 unit. Negative = they move in opposite directions.',
        'Std Error: How uncertain we are about the coefficient. Smaller = more confident.',
        'Match?: TRUE means the Excel formula matches the Python calculation (they should all be TRUE).',
        '',
        'Key takeaway: The "Chg in 10Y Rate" coefficient in the first regression is large and negative,',
        'meaning when long-term rates rise, REITs underperform the S&P 500 significantly.',
        'Oil price changes have almost no effect on the REIT vs S&P spread.',
    ]
    for line in notes:
        ws3.cell(row=r, column=1, value=line)
        r += 1

    _auto_width(ws3)

    # ==================================================================
    # TAB 4: CHARTS
    # ==================================================================
    ws4 = wb.create_sheet('Charts')
    ws4.sheet_properties.tabColor = '7030A0'

    # Chart 1: Oil Change vs REIT Excess Return
    c1 = ScatterChart()
    c1.title = 'Do Oil Price Swings Affect REITs vs S&P?'
    c1.x_axis.title = 'Monthly Oil Price Move (%)'
    c1.y_axis.title = 'How Much REITs Beat/Trail S&P (%)'
    c1.width = 20
    c1.height = 14
    xv = Reference(ws, min_col=3, min_row=2, max_row=last_row)
    yv = Reference(ws, min_col=6, min_row=2, max_row=last_row)
    s = Series(yv, xv, title='Each dot = 1 month')
    s.graphicalProperties.noFill = True
    c1.series.append(s)
    ws4.add_chart(c1, 'A1')

    # Chart 2: 10Y Rate Change vs REIT Excess Return
    c2 = ScatterChart()
    c2.title = 'When Long-Term Rates Rise, REITs Underperform'
    c2.x_axis.title = 'Monthly Change in 10Y Treasury Rate (pp)'
    c2.y_axis.title = 'How Much REITs Beat/Trail S&P (%)'
    c2.width = 20
    c2.height = 14
    xv2 = Reference(ws, min_col=11, min_row=2, max_row=last_row)
    yv2 = Reference(ws, min_col=6, min_row=2, max_row=last_row)
    s2 = Series(yv2, xv2, title='Each dot = 1 month')
    s2.graphicalProperties.noFill = True
    c2.series.append(s2)
    ws4.add_chart(c2, 'L1')

    # Chart 3: Oil Change vs 10Y Change
    c3 = ScatterChart()
    c3.title = 'Oil Prices and Long-Term Rates Move Together (Slightly)'
    c3.x_axis.title = 'Monthly Oil Price Move (%)'
    c3.y_axis.title = 'Monthly Change in 10Y Rate (pp)'
    c3.width = 20
    c3.height = 14
    xv3 = Reference(ws, min_col=3, min_row=2, max_row=last_row)
    yv3 = Reference(ws, min_col=11, min_row=2, max_row=last_row)
    s3 = Series(yv3, xv3, title='Each dot = 1 month')
    s3.graphicalProperties.noFill = True
    c3.series.append(s3)
    ws4.add_chart(c3, 'A18')

    # Chart 4: Oil Price + 10Y over time (dual axis)
    c4 = LineChart()
    c4.title = 'Oil Price and 10Y Treasury Rate Over Time'
    c4.y_axis.title = 'Oil Price ($/bbl)'
    c4.width = 20
    c4.height = 14
    cats = Reference(ws, min_col=1, min_row=2, max_row=last_row)
    v_oil = Reference(ws, min_col=2, min_row=1, max_row=last_row)
    v_10y = Reference(ws, min_col=8, min_row=1, max_row=last_row)
    c4.add_data(v_oil, titles_from_data=True)
    c4.set_categories(cats)
    c4.series[0].graphicalProperties.line.width = 15000
    c4b = LineChart()
    c4b.y_axis.title = '10Y Rate (%)'
    c4b.add_data(v_10y, titles_from_data=True)
    c4b.set_categories(cats)
    c4b.y_axis.axId = 200
    c4b.series[0].graphicalProperties.line.width = 15000
    c4.y_axis.crosses = 'min'
    c4 += c4b
    ws4.add_chart(c4, 'L18')

    # Save
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
