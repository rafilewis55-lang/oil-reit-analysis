"""
Flask app — Oil, REITs & Interest Rates analysis.
Displays summary + interactive charts, lets users download Excel workbook.
"""

import warnings
warnings.filterwarnings('ignore')

import json
import numpy as np
import io
from flask import Flask, render_template, send_file
from data_fetch import fetch_all, run_regressions
from excel_builder import build_excel

app = Flask(__name__)
_cache = {}


def get_results():
    if 'data' not in _cache:
        _cache['data'] = fetch_all()
        _cache['regs'] = run_regressions(_cache['data'])
    return _cache['data'], _cache['regs']


@app.route('/')
def index():
    data, regs = get_results()
    df = data['df']

    summary = {
        'n_obs': len(df),
        'start': df.index[0].strftime('%B %Y'),
        'end': df.index[-1].strftime('%B %Y'),
    }

    def reg_summary(model):
        rows = []
        for var in model.params.index:
            pval = float(model.pvalues[var])
            sig = '***' if pval < 0.01 else '**' if pval < 0.05 else '*' if pval < 0.1 else ''
            rows.append({
                'var': var,
                'coef': f"{model.params[var]:.4f}",
                'se': f"{model.bse[var]:.4f}",
                'tstat': f"{model.tvalues[var]:.3f}",
                'pval': f"{pval:.4f}",
                'sig': sig,
            })
        return {
            'r2': f"{model.rsquared:.3f}",
            'r2_adj': f"{model.rsquared_adj:.3f}",
            'f_pval': f"{model.f_pvalue:.4f}",
            'n': int(model.nobs),
            'rows': rows,
        }

    reit_models = [
        ('Oil + Rate Changes (Main Model)', regs['reit_m2']),
        ('Oil + Rate Levels', regs['reit_m1']),
        ('Asymmetric Oil + Rate Changes', regs['reit_m3']),
    ]
    oil_models = [
        ('Do Rate Changes Explain Oil?', regs['oil_rates_changes']),
        ('Does Oil Move the 10Y Rate?', regs['t10y_on_oil']),
        ('Does Oil Move the 3M Rate?', regs['t3m_on_oil']),
        ('Does Lagged Oil Predict 10Y?', regs['t10y_oil_lagged']),
        ('Does Lagged Oil Predict 3M?', regs['t3m_oil_lagged']),
    ]

    reit_regs = [(name, reg_summary(m)) for name, m in reit_models]
    oil_regs = [(name, reg_summary(m)) for name, m in oil_models]

    desc = df[['oil_chg', 'reit_ret', 'spx_ret', 'excess_ret', 't3m', 't10y']].describe().round(3)

    # Chart data as JSON for Chart.js
    chart_data = {
        'dates': [d.strftime('%Y-%m') for d in df.index],
        'oil_chg': [round(float(v), 2) for v in df['oil_chg']],
        'excess_ret': [round(float(v), 2) for v in df['excess_ret']],
        'd_t10y': [round(float(v), 3) for v in df['d_t10y']],
        'd_t3m': [round(float(v), 3) for v in df['d_t3m']],
        'oil_price': [round(float(v), 2) for v in df['oil_price']],
        't10y': [round(float(v), 2) for v in df['t10y']],
        't3m': [round(float(v), 2) for v in df['t3m']],
        'rolling_excess': [round(float(v), 2) if not np.isnan(v) else None
                           for v in df['excess_ret'].rolling(12).mean()],
    }

    return render_template('index.html',
                           summary=summary,
                           reit_regs=reit_regs,
                           oil_regs=oil_regs,
                           desc=desc,
                           chart_data=json.dumps(chart_data))


@app.route('/download')
def download():
    data, regs = get_results()
    excel_bytes = build_excel(data, regs)
    return send_file(
        io.BytesIO(excel_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Oil_REITs_Rates_Analysis.xlsx'
    )


if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    print("Fetching data (first run may take 30s)...")
    get_results()
    print(f"Ready! Open http://127.0.0.1:{port}")
    app.run(debug=False, host='0.0.0.0', port=port)
