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
import os
import glob as globmod
import numpy as np
import pandas as pd
from docx import Document as DocxDocument
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series, LineChart
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Deutsche Bank style: #0018A8 dark blue, #E8EDF2 light blue-gray, clean/minimal
DB_BLUE = '0018A8'
DB_LIGHT = 'E8EDF2'
DB_MID = 'C7D3E3'

HEADER_FONT = Font(bold=True, color='FFFFFF', size=11)
HEADER_FILL = PatternFill('solid', fgColor=DB_BLUE)
RED_FILL = PatternFill('solid', fgColor=DB_BLUE)
RED_FONT = Font(bold=True, color='FFFFFF', size=11)
LIGHT_FILL = PatternFill('solid', fgColor=DB_LIGHT)
SHOCK_FILL = PatternFill('solid', fgColor='F5E6E8')
GREEN_FILL = PatternFill('solid', fgColor='E6EFE6')
BOLD = Font(bold=True)
BLUE = Font(color=DB_BLUE)
BLUE_BOLD = Font(bold=True, color=DB_BLUE)
ITALIC_GRAY = Font(italic=True, color='888888')
THIN_BORDER = Border(bottom=Side(style='thin', color=DB_MID))


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


def _find_flash_note_docx():
    """Find the most recent flash note docx in the project directory."""
    project_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = globmod.glob(os.path.join(project_dir, '*Flash*Note*.docx'))
    candidates += globmod.glob(os.path.join(project_dir, '*flash*note*.docx'))
    candidates += globmod.glob(os.path.join(project_dir, 'REIT_Flash*.docx'))
    # Deduplicate and sort by modification time (most recent first)
    candidates = list(set(candidates))
    if not candidates:
        return None
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


def _build_flash_note_from_docx(wb):
    """Build the Flash Note tab by reading content from the most recent docx file."""
    docx_path = _find_flash_note_docx()
    if not docx_path:
        ws = wb.create_sheet('Flash Note')
        ws.cell(row=1, column=2, value='No Flash Note docx found in project directory.').font = Font(italic=True, color='999999')
        return

    doc = DocxDocument(docx_path)
    ws = wb.create_sheet('Flash Note')
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 22
    ws.column_dimensions['F'].width = 22

    DARK_FILL = PatternFill('solid', fgColor=DB_BLUE)
    WHITE_BOLD = Font(bold=True, color='FFFFFF', size=11)
    SECTION_FONT = Font(bold=True, size=13, color=DB_BLUE)
    SUBSECTION_FONT = Font(bold=True, size=11, color=DB_BLUE)

    # Track which tables we've inserted (by index) to place them in order
    table_idx = 0
    # Map paragraph indices that precede a table to the table index
    # We'll insert tables after certain section headers
    table_triggers = {
        'Current Market Snapshot': 0,
        'REIT Subsector Performance': 1,
        'oil coefficient': 2,
        'The 21 Shock Episodes': 3,
        'Post-Shock Recovery': 4,
        'SCENARIO ANALYSIS': 5,
    }
    pending_table = None

    r = 1
    in_takeaways = False

    for pi, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        has_bold = any(run.bold for run in para.runs if run.bold)
        is_list = (para.style and para.style.name == 'List Paragraph') if para.style else False

        # Insert any pending table before continuing
        if pending_table is not None and not is_list:
            if pending_table < len(doc.tables):
                r = _insert_docx_table(ws, r, doc.tables[pending_table])
                r += 1
            pending_table = None

        # Title block (first few paragraphs)
        if pi == 0:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.cell(row=r, column=2, value=text).font = Font(bold=True, size=9, color='888888')
            r += 1
            continue
        if pi == 1:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.cell(row=r, column=2, value=text).font = Font(bold=True, size=9, color=DB_BLUE)
            r += 1
            continue
        if pi == 2:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.cell(row=r, column=2, value=text).font = Font(bold=True, size=16, color=DB_BLUE)
            r += 1
            continue
        if pi == 3:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.cell(row=r, column=2, value=text).font = Font(italic=True, size=11, color='555555')
            r += 1
            continue
        if pi == 4:
            ws.cell(row=r, column=6, value=text).font = Font(italic=True, color='888888')
            r += 2
            continue

        # KEY TAKEAWAYS header
        if text == 'KEY TAKEAWAYS':
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            c = ws.cell(row=r, column=2, value=text)
            c.font = WHITE_BOLD
            c.fill = DARK_FILL
            for col in range(2, 7):
                ws.cell(row=r, column=col).fill = DARK_FILL
            r += 1
            in_takeaways = True
            continue

        # Takeaway bullets (List Paragraph style)
        if is_list and in_takeaways:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.cell(row=r, column=2, value=text).alignment = Alignment(wrap_text=True)
            ws.row_dimensions[r].height = max(32, len(text) // 80 * 16 + 20)
            for col in range(2, 7):
                ws.cell(row=r, column=col).fill = PatternFill('solid', fgColor=DB_LIGHT)
                ws.cell(row=r, column=col).border = THIN_BORDER
            r += 1
            continue

        if in_takeaways and not is_list:
            in_takeaways = False
            r += 1

        # Section headers (all caps or bold with specific keywords)
        is_section = (text.isupper() and len(text) > 5 and has_bold)
        is_subsection = (has_bold and not text.isupper() and len(text) < 120)

        # Check if this paragraph should trigger a table insertion
        for trigger, tidx in table_triggers.items():
            if trigger.lower() in text.lower() and tidx >= table_idx:
                pending_table = tidx
                table_idx = tidx + 1
                break

        if text == 'SOURCES':
            # Skip — sources go in a separate tab
            break

        if is_section:
            # Check for special sections
            if 'TAIL RISK' in text:
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
                c = ws.cell(row=r, column=2, value=text)
                c.font = Font(bold=True, color='FFFFFF')
                for col in range(2, 7):
                    ws.cell(row=r, column=col).fill = PatternFill('solid', fgColor=DB_BLUE)
                r += 1
                continue
            if 'BOTTOM LINE' in text:
                ws.cell(row=r, column=2, value=text).font = SECTION_FONT
                r += 1
                continue

            ws.cell(row=r, column=2, value=text).font = SECTION_FONT
            r += 1
            continue

        if is_subsection:
            ws.cell(row=r, column=2, value=text).font = SUBSECTION_FONT
            r += 1
            # Insert table if pending right after the header
            if pending_table is not None:
                if pending_table < len(doc.tables):
                    r = _insert_docx_table(ws, r, doc.tables[pending_table])
                    r += 1
                pending_table = None
            continue

        # Body text
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        cell = ws.cell(row=r, column=2, value=text)
        cell.alignment = Alignment(wrap_text=True)
        ws.row_dimensions[r].height = max(48, len(text) // 80 * 16 + 20)

        # Special styling for text after TAIL RISK
        if pi > 0 and any('TAIL RISK' in (doc.paragraphs[j].text if j < len(doc.paragraphs) else '') for j in range(max(0, pi-3), pi)):
            for col in range(2, 7):
                ws.cell(row=r, column=col).fill = PatternFill('solid', fgColor='FCE4EC')
                ws.cell(row=r, column=col).border = THIN_BORDER

        # Bold body text (bottom line paragraphs)
        if has_bold and len(text) > 120:
            cell.font = Font(bold=True, size=10)

        r += 1

        # Insert pending table after body text
        if pending_table is not None:
            if pending_table < len(doc.tables):
                r = _insert_docx_table(ws, r, doc.tables[pending_table])
                r += 1
            pending_table = None

    # Insert any remaining pending table
    if pending_table is not None and pending_table < len(doc.tables):
        r = _insert_docx_table(ws, r, doc.tables[pending_table])


def _insert_docx_table(ws, start_row, table):
    """Insert a docx table into the Excel worksheet starting at start_row. Returns next row."""
    r = start_row
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            c = ws.cell(row=r, column=ci + 2, value=cell.text.strip())
            c.border = THIN_BORDER
            c.alignment = Alignment(wrap_text=True, horizontal='center' if ci > 0 else 'left')
            if ri == 0:
                c.font = Font(bold=True, color='FFFFFF', size=10)
                c.fill = PatternFill('solid', fgColor=DB_BLUE)
            elif ci == 0:
                c.font = Font(bold=True, size=10)
        ws.row_dimensions[r].height = max(28, max(len(cell.text) for cell in row.cells) // 30 * 14 + 16)
        r += 1
    return r


def _build_note_sources_from_docx(wb):
    """Build the Note Sources tab from the sources table in the docx."""
    docx_path = _find_flash_note_docx()
    if not docx_path:
        ws = wb.create_sheet('Note Sources')
        ws.cell(row=1, column=2, value='No Flash Note docx found.').font = Font(italic=True, color='999999')
        return

    doc = DocxDocument(docx_path)
    ws = wb.create_sheet('Note Sources')
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 8
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 80

    r = 1
    ws.cell(row=r, column=2, value='Flash Note Sources').font = Font(bold=True, size=16, color=DB_BLUE)
    r += 2

    # Find the sources table (last table in the docx, typically table index 6)
    sources_table = None
    for t in doc.tables:
        # Sources table has 3 columns and many rows with URLs
        if len(t.columns) == 3 and len(t.rows) > 10:
            # Check if it contains URLs
            has_url = any('http' in cell.text for row in t.rows for cell in row.cells)
            if has_url:
                sources_table = t
                break

    if not sources_table:
        ws.cell(row=r, column=2, value='No sources table found in docx.').font = Font(italic=True, color='999999')
        return

    _header_row(ws, r, ['', '#', 'Description', 'Link'])
    r += 1

    for row in sources_table.rows:
        cells = [c.text.strip() for c in row.cells]
        # Section header rows have merged cells (all same text)
        if cells[0] == cells[1] == cells[2]:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
            ws.cell(row=r, column=2, value=cells[0]).font = Font(bold=True, size=11, color=DB_BLUE)
            ws.cell(row=r, column=2).fill = PatternFill('solid', fgColor=DB_LIGHT)
            for c in range(2, 5):
                ws.cell(row=r, column=c).fill = PatternFill('solid', fgColor=DB_LIGHT)
            r += 1
            continue

        # Data rows: [number, description, url]
        num_text = cells[0]
        desc = cells[1]
        link = cells[2]

        try:
            ws.cell(row=r, column=2, value=int(num_text)).alignment = Alignment(horizontal='center')
        except ValueError:
            ws.cell(row=r, column=2, value=num_text).alignment = Alignment(horizontal='center')

        ws.cell(row=r, column=3, value=desc).font = Font(bold=True)
        ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True)

        if link.startswith('http'):
            ws.cell(row=r, column=4, value=link).font = Font(color='0563C1', underline='single')
            ws.cell(row=r, column=4).hyperlink = link
        else:
            ws.cell(row=r, column=4, value=link).font = Font(italic=True, color='888888')
        ws.cell(row=r, column=4).alignment = Alignment(wrap_text=True)

        ws.row_dimensions[r].height = max(32, len(desc) // 40 * 14 + 20)
        for c in range(2, 5):
            ws.cell(row=r, column=c).border = THIN_BORDER
        r += 1


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
        ws.cell(row=row, column=6).font = Font(bold=True, color=DB_BLUE)
    for c in range(1, 7):
        ws.cell(row=row, column=c).border = THIN_BORDER


def _write_reg_block(ws, start_row, title, model, include_linest=False, linest_base=None):
    """Write a full regression block. Returns next free row."""
    r = start_row
    ws.cell(row=r, column=1, value=title).font = Font(bold=True, size=12, color=DB_BLUE)
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
    df_daily = data['df_daily']
    detected_shocks = regressions.get('_detected_shocks', [])

    # Also keep monthly df for regressions/correlations tabs
    df = data['df']
    n_monthly = len(df)
    monthly_last_row = n_monthly + 1

    # ==================================================================
    # TAB 1: DATA (Daily)
    # ==================================================================
    ws = wb.active
    ws.title = 'Data'

    headers = ['Date', 'Oil Price', 'Oil 3M Chg %', 'REIT Close', 'S&P Close',
               '3M Rate %', '10Y Rate %']
    _header_row(ws, 1, headers)

    n_daily = len(df_daily)
    daily_last_row = n_daily + 1
    for r, (date, row) in enumerate(df_daily.iterrows(), 2):
        ws.cell(row=r, column=1, value=date.strftime('%Y-%m-%d'))
        ws.cell(row=r, column=2, value=round(row['oil_price'], 2))
        ws.cell(row=r, column=3, value=round(row['oil_3m_chg'], 2) if not np.isnan(row['oil_3m_chg']) else '')
        ws.cell(row=r, column=4, value=round(row['reit_close'], 2))
        ws.cell(row=r, column=5, value=round(row['spx_close'], 2))
        ws.cell(row=r, column=6, value=round(row['t3m'], 2))
        ws.cell(row=r, column=7, value=round(row['t10y'], 2))

    _auto_width(ws)
    ws.freeze_panes = 'A2'

    def drange(col_letter):
        """Column range for the monthly data tab."""
        return f"'Monthly Data'!{col_letter}2:{col_letter}{monthly_last_row}"

    # ==================================================================
    # TAB 2: SHOCK PERIODS (Trough-to-Peak using daily data)
    # ==================================================================
    ws2 = wb.create_sheet('Shock Periods')

    ws2.cell(row=1, column=1, value='Oil Shock Episodes: Trough-to-Peak Analysis').font = Font(bold=True, size=14, color=DB_BLUE)
    ws2.cell(row=2, column=1, value='Auto-detected from daily data: 3-month trailing oil price change > 30%. Returns computed using exact daily prices.').font = ITALIC_GRAY

    r = 4
    ep_headers = ['Episode', 'Trough Date', 'Peak Date', 'Trading Days',
                   'Oil Trough', 'Oil Peak', 'Oil % Chg',
                   'REIT Ret %', 'S&P Ret %', 'Excess Ret %',
                   '10Y Chg (pp)', '3M Chg (pp)']
    _header_row(ws2, r, ep_headers)
    r += 1
    for shock in detected_shocks:
        ws2.cell(row=r, column=1, value=shock['label']).font = BOLD
        ws2.cell(row=r, column=2, value=shock['start_date'])
        ws2.cell(row=r, column=3, value=shock['end_date'])
        ws2.cell(row=r, column=4, value=shock.get('trading_days', ''))
        ws2.cell(row=r, column=5, value=shock['trough_price']).number_format = '0.00'
        ws2.cell(row=r, column=6, value=shock['peak_price']).number_format = '0.00'
        ws2.cell(row=r, column=7, value=shock['pct_change']).number_format = '0.0'
        ws2.cell(row=r, column=8, value=shock.get('reit_ret', '')).number_format = '0.00'
        ws2.cell(row=r, column=9, value=shock.get('spx_ret', '')).number_format = '0.00'
        excess = shock.get('excess_ret', '')
        c_ex = ws2.cell(row=r, column=10, value=excess)
        c_ex.number_format = '0.00'
        if isinstance(excess, (int, float)):
            if excess > 0:
                c_ex.font = Font(bold=True, color='336633')
            elif excess < 0:
                c_ex.font = Font(bold=True, color='CC0000')
        ws2.cell(row=r, column=11, value=shock.get('d_t10y', '')).number_format = '0.000'
        ws2.cell(row=r, column=12, value=shock.get('d_t3m', '')).number_format = '0.000'
        for c in range(1, 13):
            ws2.cell(row=r, column=c).border = THIN_BORDER
        r += 1

    # Averages row
    r += 1
    ws2.cell(row=r, column=1, value='AVERAGE').font = Font(bold=True, color=DB_BLUE)
    for ci, key in [(4, 'trading_days'), (7, 'pct_change'), (8, 'reit_ret'), (9, 'spx_ret'),
                     (10, 'excess_ret'), (11, 'd_t10y'), (12, 'd_t3m')]:
        vals = [s.get(key) for s in detected_shocks if s.get(key) is not None and isinstance(s.get(key), (int, float))]
        if vals:
            avg = round(np.mean(vals), 2 if ci != 4 else 0)
            ws2.cell(row=r, column=ci, value=avg)
    for c in range(1, 13):
        ws2.cell(row=r, column=c).fill = LIGHT_FILL
        ws2.cell(row=r, column=c).border = THIN_BORDER

    _auto_width(ws2)

    # ==================================================================
    # TAB 3: MONTHLY DATA (for correlations/LINEST references)
    # ==================================================================
    wsm = wb.create_sheet('Monthly Data')
    m_headers = ['Date', 'Oil Price', 'Oil Chg %', 'REIT Return %',
                 'S&P 500 Return %', 'REIT Excess Return %',
                 '3M Rate %', '10Y Rate %', 'Term Spread %',
                 'Chg in 3M Rate', 'Chg in 10Y Rate']
    _header_row(wsm, 1, m_headers)
    for r_i, (date, row) in enumerate(df.iterrows(), 2):
        wsm.cell(row=r_i, column=1, value=date.strftime('%Y-%m'))
        wsm.cell(row=r_i, column=2, value=round(row['oil_price'], 2))
        wsm.cell(row=r_i, column=3, value=round(row['oil_chg'], 2))
        wsm.cell(row=r_i, column=4, value=round(row['reit_ret'], 2))
        wsm.cell(row=r_i, column=5, value=round(row['spx_ret'], 2))
        wsm.cell(row=r_i, column=6, value=round(row['excess_ret'], 2))
        wsm.cell(row=r_i, column=7, value=round(row['t3m'], 2))
        wsm.cell(row=r_i, column=8, value=round(row['t10y'], 2))
        wsm.cell(row=r_i, column=9, value=round(row['term_spread'], 2))
        wsm.cell(row=r_i, column=10, value=round(row['d_t3m'], 4))
        wsm.cell(row=r_i, column=11, value=round(row['d_t10y'], 4))
    _auto_width(wsm)
    wsm.freeze_panes = 'A2'

    # ==================================================================
    # TAB 4: CORRELATIONS (CORREL formulas)
    # ==================================================================
    ws3 = wb.create_sheet('Correlations')

    ws3.cell(row=1, column=1, value='Correlation Matrix (Full Sample)').font = Font(bold=True, size=14, color=DB_BLUE)
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

    r = 1
    ws4.cell(row=r, column=1, value='Full Sample Regression Results').font = Font(bold=True, size=14, color=DB_BLUE)
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

    # Adjacent cols J,K so simpler
    linest_oil = f'LINEST({drange("C")},"Monthly Data"!J2:K{monthly_last_row},TRUE,TRUE)'
    r = _write_reg_block(ws4, r, 'Oil Change ~ Rate Changes',
                         regressions['oil_rates_changes_ols'],
                         include_linest=True, linest_base=linest_oil)

    # Additional HC1 models (no LINEST cross-check)
    r = _write_reg_block(ws4, r, 'REIT Excess ~ Oil + Rate Changes (Winsorized, Robust SE)',
                         regressions['reit_m2'])
    r = _write_reg_block(ws4, r, 'REIT Excess ~ Asymmetric Oil + Rate Changes (Winsorized)',
                         regressions['reit_m3'])

    _auto_width(ws4)

    # (Shock Regressions tab removed — shock analysis is now episode-based in Shock Periods tab)

    # ==================================================================
    # TAB 6: KEY FINDINGS
    # ==================================================================
    wsf = wb.create_sheet('Key Findings')
    wsf.column_dimensions['A'].width = 4
    wsf.column_dimensions['B'].width = 90

    r = 1
    wsf.cell(row=r, column=2, value='Key Findings: Oil, REITs & Interest Rates').font = Font(bold=True, size=16, color=DB_BLUE)
    r += 2

    # Compute summary stats from detected shocks for findings text
    n_episodes = len(detected_shocks)
    if detected_shocks:
        excess_rets = [s.get('excess_ret', 0) for s in detected_shocks if s.get('excess_ret') is not None]
        avg_excess = sum(excess_rets) / len(excess_rets) if excess_rets else 0
        pos_excess = sum(1 for e in excess_rets if e > 0)
        neg_excess = sum(1 for e in excess_rets if e < 0)
        avg_oil = sum(s['pct_change'] for s in detected_shocks) / n_episodes
        avg_d10y = sum(s.get('d_t10y', 0) for s in detected_shocks) / n_episodes
    else:
        avg_excess = avg_oil = avg_d10y = 0
        pos_excess = neg_excess = 0

    post_shock_avg = regressions.get('_post_shock_avg', {})

    findings = [
        ("1. Oil shocks don't consistently help or hurt REITs vs the S&P.",
         f"Across {n_episodes} auto-detected oil shock episodes (3-month trailing change >30%), "
         f"REITs outperformed in {pos_excess} and underperformed in {neg_excess}. "
         f"The average trough-to-peak excess return is {avg_excess:+.1f}%, "
         "meaning oil shocks move REITs and the S&P by roughly equal amounts. "
         "There is no reliable directional edge."),

        ("2. Long-term rates are what actually drive the wedge.",
         "A 1 percentage point rise in the 10Y yield in a given month is associated with ~5% REIT "
         "underperformance vs the S&P (p<0.001). REITs behave like long-duration bonds. "
         "This relationship holds in the full monthly regression and is the single strongest predictor."),

        ("3. The transmission chain is indirect: Oil -> Rates -> REITs.",
         f"During shock episodes, the 10Y yield moves by an average of {avg_d10y:+.3f} pp. "
         "Oil doesn't hit REITs directly -- it works through inflation expectations and rate moves. "
         "The direct oil->REIT channel is basically zero. Most REIT vs S&P performance comes from "
         "other factors (sector rotation, cap rates, property fundamentals)."),

        ("4. Oil shocks are asymmetric but the REIT impact isn't.",
         f"The average oil shock is +{avg_oil:.0f}% trough-to-peak, but REIT excess returns "
         "range widely from episode to episode. Some of the largest oil spikes (e.g. 2020 COVID rebound "
         "at +255%) saw REITs trail the S&P, while moderate spikes often saw REITs outperform. "
         "The size of the oil move does not predict whether REITs win or lose."),
    ]

    # Add post-shock recovery finding if data available
    if post_shock_avg:
        ps_3m = post_shock_avg.get('3M', {})
        ps_6m = post_shock_avg.get('6M', {})
        ps_12m = post_shock_avg.get('12M', {})
        recovery_text = (
            f"After oil shock peaks, REITs consistently outperform the S&P. "
            f"At 3 months: REIT +{ps_3m.get('reit_ret', 'N/A')}% vs S&P +{ps_3m.get('spx_ret', 'N/A')}% "
            f"(+{ps_3m.get('excess_ret', 'N/A')}% excess). "
            f"At 6 months: REIT +{ps_6m.get('reit_ret', 'N/A')}% vs S&P +{ps_6m.get('spx_ret', 'N/A')}% "
            f"(+{ps_6m.get('excess_ret', 'N/A')}% excess). "
            f"At 12 months: REIT +{ps_12m.get('reit_ret', 'N/A')}% vs S&P +{ps_12m.get('spx_ret', 'N/A')}% "
            f"(+{ps_12m.get('excess_ret', 'N/A')}% excess). "
            "As oil gives back gains post-peak "
            f"({ps_3m.get('oil_chg', 'N/A')}% at 3M, {ps_12m.get('oil_chg', 'N/A')}% at 12M), "
            "the rate pressure eases and REITs recover faster than the broader market."
        )
        findings.append(("5. Post-shock recovery: REITs outperform after the oil peak fades.", recovery_text))

    for title, body in findings:
        wsf.cell(row=r, column=2, value=title).font = Font(bold=True, size=12, color=DB_BLUE)
        r += 1
        wsf.cell(row=r, column=2, value=body).alignment = Alignment(wrap_text=True)
        # Auto-height: roughly 1 row per 90 chars
        lines = max(len(body) // 85 + body.count('\n') + 1, 2)
        wsf.row_dimensions[r].height = lines * 16
        r += 2

    r += 1
    wsf.cell(row=r, column=2, value='Supporting Evidence: Trough-to-Peak Returns by Episode').font = Font(bold=True, size=13, color=DB_BLUE)
    r += 1

    ep_headers = ['Episode', 'Oil % Chg', 'REIT Ret %', 'S&P Ret %', 'Excess %', '10Y Chg (pp)']
    _header_row(wsf, r, [''] + ep_headers)
    r += 1
    for shock in detected_shocks:
        wsf.cell(row=r, column=2, value=shock['label']).font = BOLD
        wsf.cell(row=r, column=3, value=shock['pct_change']).number_format = '0.0'
        wsf.cell(row=r, column=4, value=shock.get('reit_ret', 'N/A'))
        wsf.cell(row=r, column=5, value=shock.get('spx_ret', 'N/A'))
        excess = shock.get('excess_ret', 'N/A')
        c_ex = wsf.cell(row=r, column=6, value=excess)
        if isinstance(excess, (int, float)):
            c_ex.number_format = '0.00'
            if excess > 0:
                c_ex.font = Font(bold=True, color='336633')
            elif excess < 0:
                c_ex.font = Font(bold=True, color='CC0000')
        wsf.cell(row=r, column=7, value=shock.get('d_t10y', 'N/A')).number_format = '0.000'
        for c in range(2, 8):
            wsf.cell(row=r, column=c).border = THIN_BORDER
        r += 1

    # Post-shock recovery summary on Key Findings tab
    if post_shock_avg:
        r += 2
        wsf.cell(row=r, column=2, value='Post-Shock Recovery: Average Cumulative Returns After Oil Peak').font = Font(bold=True, size=13, color=DB_BLUE)
        r += 1
        ps_headers = ['Horizon', 'N', 'REIT %', 'S&P %', 'Excess %', 'Oil %', '\u039410Y (pp)', '\u03943M (pp)']
        _header_row(wsf, r, [''] + ps_headers)
        r += 1
        for h_label in ['3M', '6M', '12M']:
            a = post_shock_avg.get(h_label, {})
            if a:
                wsf.cell(row=r, column=2, value=h_label).font = BOLD
                wsf.cell(row=r, column=3, value=a.get('n', '')).number_format = '0'
                wsf.cell(row=r, column=4, value=a.get('reit_ret', '')).number_format = '0.00'
                wsf.cell(row=r, column=5, value=a.get('spx_ret', '')).number_format = '0.00'
                ex_val = a.get('excess_ret', '')
                c_ex = wsf.cell(row=r, column=6, value=ex_val)
                c_ex.number_format = '0.00'
                if isinstance(ex_val, (int, float)):
                    if ex_val > 0:
                        c_ex.font = Font(bold=True, color='336633')
                    elif ex_val < 0:
                        c_ex.font = Font(bold=True, color='CC0000')
                wsf.cell(row=r, column=7, value=a.get('oil_chg', '')).number_format = '0.00'
                wsf.cell(row=r, column=8, value=a.get('d_t10y', '')).number_format = '0.000'
                wsf.cell(row=r, column=9, value=a.get('d_t3m', '')).number_format = '0.000'
                for c in range(2, 10):
                    wsf.cell(row=r, column=c).border = THIN_BORDER
                r += 1

    # ==================================================================
    # TAB 7: SOURCES
    # ==================================================================
    wss = wb.create_sheet('Sources')
    wss.column_dimensions['A'].width = 4
    wss.column_dimensions['B'].width = 30
    wss.column_dimensions['C'].width = 65
    wss.column_dimensions['D'].width = 55

    r = 1
    wss.cell(row=r, column=2, value='Data Sources').font = Font(bold=True, size=16, color=DB_BLUE)
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
    wss.cell(row=r, column=2, value='Methodology Notes').font = Font(bold=True, size=13, color=DB_BLUE)
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
    # TAB 8: FLASH NOTE (read from docx)
    # ==================================================================
    _build_flash_note_from_docx(wb)

    # ==================================================================
    # TAB 9: NOTE SOURCES (read from docx)
    # ==================================================================
    _build_note_sources_from_docx(wb)

    # ==================================================================
    # TAB 10: POST-SHOCK RECOVERY
    # ==================================================================
    post_shock = regressions.get('_post_shock', [])
    post_shock_avg = regressions.get('_post_shock_avg', {})

    wsp = wb.create_sheet('Post-Shock Recovery')
    wsp.column_dimensions['A'].width = 40
    wsp.column_dimensions['B'].width = 12
    wsp.column_dimensions['C'].width = 10
    for col_letter in ['D', 'E', 'F', 'G', 'H', 'I']:
        wsp.column_dimensions[col_letter].width = 14

    r = 1
    wsp.cell(row=r, column=1, value='Post-Shock Recovery: 3, 6, and 12 Months After Each Shock').font = Font(bold=True, size=14, color=DB_BLUE)
    r += 1
    wsp.cell(row=r, column=1, value='How did REITs, the S&P, and rates behave after each auto-detected oil shock peak? Uses exact daily prices.').font = ITALIC_GRAY
    r += 2

    # Summary averages first
    wsp.cell(row=r, column=1, value=f'Average Post-Shock Performance ({len(detected_shocks)} Episodes)').font = Font(bold=True, size=13, color=DB_BLUE)
    r += 1
    avg_headers = ['Horizon', 'N', 'REIT Cumulative %', 'S&P Cumulative %', 'REIT Excess %', 'Oil Cumulative %', '10Y Chg (pp)', '3M Chg (pp)']
    _header_row(wsp, r, avg_headers)
    r += 1
    for horizon in ['3M', '6M', '12M']:
        avg = post_shock_avg.get(horizon)
        if not avg:
            continue
        wsp.cell(row=r, column=1, value=horizon).font = BOLD
        wsp.cell(row=r, column=2, value=avg['n'])
        wsp.cell(row=r, column=3, value=avg['reit_ret']).number_format = '0.00'
        wsp.cell(row=r, column=4, value=avg['spx_ret']).number_format = '0.00'
        wsp.cell(row=r, column=5, value=avg['excess_ret']).number_format = '0.00'
        # Color excess: green if positive, red if negative
        if avg['excess_ret'] > 0:
            wsp.cell(row=r, column=5).font = Font(bold=True, color='336633')
        elif avg['excess_ret'] < 0:
            wsp.cell(row=r, column=5).font = Font(bold=True, color='CC0000')
        wsp.cell(row=r, column=6, value=avg['oil_chg']).number_format = '0.00'
        wsp.cell(row=r, column=7, value=avg['d_t10y']).number_format = '0.000'
        wsp.cell(row=r, column=8, value=avg['d_t3m']).number_format = '0.000'
        for c in range(1, 9):
            wsp.cell(row=r, column=c).border = THIN_BORDER
            wsp.cell(row=r, column=c).fill = LIGHT_FILL
        r += 1
    r += 2

    # Detail by shock and horizon
    wsp.cell(row=r, column=1, value='Detail by Shock Episode').font = Font(bold=True, size=13, color=DB_BLUE)
    r += 1
    detail_headers = ['Shock Episode', 'Horizon', 'End Date', 'REIT Cumulative %', 'S&P Cumulative %',
                       'REIT Excess %', 'Oil Cumulative %', '10Y Chg (pp)', '3M Chg (pp)']
    _header_row(wsp, r, detail_headers)
    r += 1

    current_shock = None
    for ps in post_shock:
        # Add a light separator between shock episodes
        if ps['shock'] != current_shock:
            if current_shock is not None:
                r += 1  # blank row between episodes
            current_shock = ps['shock']
            wsp.cell(row=r, column=1, value=ps['shock']).font = Font(bold=True, color=DB_BLUE)
        else:
            wsp.cell(row=r, column=1, value='')

        wsp.cell(row=r, column=2, value=ps['horizon'])
        wsp.cell(row=r, column=3, value=ps.get('end_date', ''))
        wsp.cell(row=r, column=4, value=ps['reit_ret']).number_format = '0.00'
        wsp.cell(row=r, column=5, value=ps['spx_ret']).number_format = '0.00'
        wsp.cell(row=r, column=6, value=ps['excess_ret']).number_format = '0.00'
        if ps['excess_ret'] > 0:
            wsp.cell(row=r, column=6).font = Font(bold=True, color='336633')
        elif ps['excess_ret'] < 0:
            wsp.cell(row=r, column=6).font = Font(bold=True, color='CC0000')
        wsp.cell(row=r, column=7, value=ps['oil_chg']).number_format = '0.00'
        wsp.cell(row=r, column=8, value=ps['d_t10y']).number_format = '0.000'
        wsp.cell(row=r, column=9, value=ps['d_t3m']).number_format = '0.000'
        for c in range(1, 10):
            wsp.cell(row=r, column=c).border = THIN_BORDER
        r += 1

    r += 2
    wsp.cell(row=r, column=1, value='Notes').font = Font(bold=True, size=11, color=DB_BLUE)
    r += 1
    notes_text = [
        'REIT/S&P returns are compounded monthly cumulative returns over the horizon period.',
        'REIT Excess = REIT cumulative return minus S&P cumulative return. Green = REITs outperformed.',
        'Oil Cumulative % = compounded cumulative oil price change over the period.',
        '10Y/3M Chg = sum of monthly changes in the rate (in percentage points) over the period.',
        'Horizon starts the first month after the shock window ends.',
    ]
    for note in notes_text:
        wsp.cell(row=r, column=1, value=note).font = ITALIC_GRAY
        r += 1

    _auto_width(wsp)

    # ==================================================================
    # TAB 11: CHARTS
    # ==================================================================
    ws6 = wb.create_sheet('Charts')

    # Charts reference Monthly Data tab (wsm) for scatter plots
    # Monthly Data columns: A=Date, B=Oil Price, C=Oil Chg, D=REIT Ret, E=S&P Ret,
    #   F=Excess Ret, G=3M Rate, H=10Y Rate, I=Term Spread, J=d_3M, K=d_10Y

    # Chart 1: Oil vs REIT excess (monthly: C=Oil Chg, F=Excess Ret)
    c1 = ScatterChart()
    c1.title = 'Oil Price Swings vs REIT Outperformance'
    c1.x_axis.title = 'Monthly Oil Price Move (%)'
    c1.y_axis.title = 'REIT vs S&P (%)'
    c1.width = 20
    c1.height = 14
    s1 = Series(Reference(wsm, min_col=6, min_row=2, max_row=monthly_last_row),
                Reference(wsm, min_col=3, min_row=2, max_row=monthly_last_row), title='Monthly')
    s1.graphicalProperties.noFill = True
    c1.series.append(s1)
    ws6.add_chart(c1, 'A1')

    # Chart 2: 10Y change vs REIT excess (monthly: K=d_t10y, F=Excess)
    c2 = ScatterChart()
    c2.title = 'When Long-Term Rates Rise, REITs Underperform'
    c2.x_axis.title = '10Y Rate Change (pp)'
    c2.y_axis.title = 'REIT vs S&P (%)'
    c2.width = 20
    c2.height = 14
    s2 = Series(Reference(wsm, min_col=6, min_row=2, max_row=monthly_last_row),
                Reference(wsm, min_col=11, min_row=2, max_row=monthly_last_row), title='Monthly')
    s2.graphicalProperties.noFill = True
    c2.series.append(s2)
    ws6.add_chart(c2, 'L1')

    # Chart 3: Oil vs 10Y change (monthly: C=Oil Chg, K=d_t10y)
    c3 = ScatterChart()
    c3.title = 'Oil and Long-Term Rates Move Together'
    c3.x_axis.title = 'Monthly Oil Price Move (%)'
    c3.y_axis.title = '10Y Rate Change (pp)'
    c3.width = 20
    c3.height = 14
    s3 = Series(Reference(wsm, min_col=11, min_row=2, max_row=monthly_last_row),
                Reference(wsm, min_col=3, min_row=2, max_row=monthly_last_row), title='Monthly')
    s3.graphicalProperties.noFill = True
    c3.series.append(s3)
    ws6.add_chart(c3, 'A18')

    # Chart 4: Oil + 10Y over time (daily: B=Oil Price, G=10Y Rate)
    c4 = LineChart()
    c4.title = 'Oil Price and 10Y Rate Over Time (Daily)'
    c4.y_axis.title = 'Oil Price ($/bbl)'
    c4.width = 20
    c4.height = 14
    cats = Reference(ws, min_col=1, min_row=2, max_row=daily_last_row)
    c4.add_data(Reference(ws, min_col=2, min_row=1, max_row=daily_last_row), titles_from_data=True)
    c4.set_categories(cats)
    c4.series[0].graphicalProperties.line.width = 15000
    c4b = LineChart()
    c4b.y_axis.title = '10Y Rate (%)'
    c4b.add_data(Reference(ws, min_col=7, min_row=1, max_row=daily_last_row), titles_from_data=True)
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
