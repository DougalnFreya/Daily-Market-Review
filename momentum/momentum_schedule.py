"""Monthly momentum ETF schedule.

Reads prices.csv (month-end closes) and unemployment.csv (US unemployment rate, FRED UNRATE),
works out the fund picks for the coming month under two ranking rules and the protection signal,
and writes "ETF schedule.xlsx".

Rules (fixed before testing, see "Momentum ETF process.docx"):
  Baskets: 8 factor funds and 7 sector funds (UK listings of the 15 funds tested).
  Ranking A (blend): average of 1, 3, 6 and 12-month returns. Ranking B: 6-month return.
  Hold the top fund in each basket, half each, for the next month.
  Protection: if the US market (SPY price) is below its 10-month average AND US unemployment
  (3-month average) is at least 0.5 points above its 12-month low, hold half cash (CSH2) and
  half gold (SGLN) instead.
Usage: python momentum_schedule.py [prices.csv] [unemployment.csv] [output.xlsx]
"""
import sys, os
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FACTOR = [('IUMF', 'USA Momentum', 'iShares Edge MSCI USA Momentum Factor', 'IE00BD1F4N50'),
          ('IUQF', 'USA Quality', 'iShares Edge MSCI USA Quality Factor', 'IE00BD1F4L37'),
          ('IUVF', 'USA Value', 'iShares Edge MSCI USA Value Factor', 'IE00BD1F4M44'),
          ('CUS1', 'USA Small Cap', 'iShares MSCI USA Small Cap ESG Enhanced CTB', 'IE00B3VWM098'),
          ('IEFM', 'Europe Momentum', 'iShares Edge MSCI Europe Momentum Factor', 'IE00BQN1K786'),
          ('IEFQ', 'Europe Quality', 'iShares Edge MSCI Europe Quality Factor', 'IE00BQN1K562'),
          ('IEFV', 'Europe Value', 'iShares Edge MSCI Europe Value Factor', 'IE00BQN1K901'),
          ('XXSC', 'Europe Small Cap', 'Xtrackers MSCI Europe Small Cap', 'LU0322253906')]
SECTOR = [('ICDU', 'Consumer Discretionary', 'iShares S&P 500 Consumer Discretionary Sector', 'IE00B4MCHD36'),
          ('ICSU', 'Consumer Staples', 'iShares S&P 500 Consumer Staples Sector', 'IE00B40B8R38'),
          ('IESU', 'Energy', 'iShares S&P 500 Energy Sector', 'IE00B42NKQ00'),
          ('IHCU', 'Health Care', 'iShares S&P 500 Health Care Sector', 'IE00B43HR379'),
          ('IISU', 'Industrials', 'iShares S&P 500 Industrials Sector', 'IE00B4LN9N13'),
          ('IITU', 'Information Technology', 'iShares S&P 500 Information Technology Sector', 'IE00B3WJKG14'),
          ('IMSU', 'Materials', 'iShares S&P 500 Materials Sector', 'IE00B4MKCJ84')]
SAFE = [('CSH2', 'Cash', 'Lyxor Smart Overnight Return (GBP)'), ('SGLN', 'Gold', 'iShares Physical Gold')]
NAME = {t: n for t, n, *_ in FACTOR + SECTOR}; NAME.update({'CSH2': 'Cash', 'SGLN': 'Gold'})


def load(prices_path, unemp_path):
    p = pd.read_csv(prices_path, dtype={'month': str}).set_index('month').sort_index()
    p.index = pd.PeriodIndex(p.index, freq='M')
    # Europe Quality: London prices are unreliable, so use the Xetra listing converted to pounds
    p['IEFQ'] = p['CEMQ'] * p['IEFV'] / p['CEMS']
    u = pd.read_csv(unemp_path, dtype={'month': str}).set_index('month').sort_index()
    u.index = pd.PeriodIndex(u.index, freq='M')
    return p, u['UNRATE']


def scores(p, at):
    lb = {k: p.loc[at] / p.loc[at - k] - 1 for k in (1, 3, 6, 12)}
    df = pd.DataFrame(lb)
    df['blend'] = df[[1, 3, 6, 12]].mean(axis=1)
    return df


def picks(p, at):
    s = scores(p, at)
    out = {}
    for rule, col in [('Blend 1-3-6-12', 'blend'), ('6-month', 6)]:
        out[rule] = {'factor': s.loc[[t for t, *_ in FACTOR], col].idxmax(),
                     'sector': s.loc[[t for t, *_ in SECTOR], col].idxmax()}
    return s, out


def protection(p, u, at):
    spy = p['SPY']
    avg10 = spy.loc[at - 9:at].mean()
    below = spy.loc[at] < avg10
    # unemployment published with a lag: use the month before the month end
    ur = u.dropna()
    last = ur.index[ur.index <= at - 1].max()
    u3 = ur.loc[last - 2:last].mean()
    low12 = ur.loc[last - 11:last].min()
    rising = round(u3 - low12, 6) >= 0.5
    return dict(spy=spy.loc[at], avg10=avg10, below=bool(below), u_month=str(last), u3=u3, low12=low12,
                rising=bool(rising), on=bool(below and rising))


def history(p, u, first, last):
    rows = []
    m = first
    while m <= last:
        s, pk = picks(p, m)
        pr = protection(p, u, m)
        nxt = m + 1
        row = {'decided at month end': str(m), 'for month': str(nxt),
               'blend factor': pk['Blend 1-3-6-12']['factor'], 'blend sector': pk['Blend 1-3-6-12']['sector'],
               '6m factor': pk['6-month']['factor'], '6m sector': pk['6-month']['sector'], 'protection on': pr['on']}
        if nxt in p.index:
            for rule, a, b in [('blend', row['blend factor'], row['blend sector']), ('6m', row['6m factor'], row['6m sector'])]:
                hold = ['CSH2', 'SGLN'] if pr['on'] else [a, b]
                row[f'{rule} return'] = float(np.mean([p.loc[nxt, t] / p.loc[m, t] - 1 for t in hold]))
        rows.append(row)
        m += 1
    return pd.DataFrame(rows)


def write(p, u, out_path):
    at = p.dropna(subset=['IUMF', 'SPY']).index.max()
    s, pk = picks(p, at)
    pr = protection(p, u, at)
    hist = history(p, u, max(p.index.min() + 12, at - 23), at)
    prev = hist.iloc[-2] if len(hist) > 1 else None

    choices = {}
    if os.path.exists(out_path):
        try:
            from openpyxl import load_workbook
            old = load_workbook(out_path)
            if 'History' in old.sheetnames:
                for rw in old['History'].iter_rows(min_row=4, values_only=True):
                    if rw and rw[0] and len(rw) > 8 and rw[8] not in (None, ''):
                        choices[str(rw[0])] = rw[8]
        except Exception as e:
            print('Could not read previous choices:', e)
    A = 'Arial'
    def f(**k): return Font(name=A, **k)
    HDR = PatternFill('solid', fgColor='1F3864'); GREEN = PatternFill('solid', fgColor='E2EFDA'); AMBER = PatternFill('solid', fgColor='FCE4D6')
    thin = Side(style='thin', color='BFBFBF')
    wb = Workbook()
    ws = wb.active; ws.title = 'This month'
    ws['A1'] = f'ETF schedule for {(at + 1).strftime("%B %Y")}'; ws['A1'].font = f(bold=True, size=15)
    ws['A2'] = f'Based on closing prices at the end of {at.strftime("%B %Y")}. Act on the first trading day of the month.'; ws['A2'].font = f(italic=True)
    r = 4
    ws.cell(r, 1, 'Protection').font = f(bold=True, size=12); r += 1
    status = 'ON: hold half cash (CSH2) and half gold (SGLN) instead of the funds below' if pr['on'] else 'OFF: hold the funds below'
    c = ws.cell(r, 1, status); c.font = f(bold=True); c.fill = AMBER if pr['on'] else GREEN; r += 1
    ws.cell(r, 1, f"US market (SPY) {pr['spy']:.2f} against 10-month average {pr['avg10']:.2f}: {'below' if pr['below'] else 'above'}").font = f(); r += 1
    ws.cell(r, 1, f"US unemployment 3-month average to {pr['u_month']}: {pr['u3']:.2f}% against 12-month low {pr['low12']:.1f}%: {'rising' if pr['rising'] else 'not rising'} (needs +0.5 points)").font = f(); r += 2
    ws.cell(r, 1, 'Funds to hold (half each)').font = f(bold=True, size=12); r += 1
    heads = ['Ranking rule', 'Factor fund', 'Sector fund', 'Change from last month']
    for j, h in enumerate(heads):
        c = ws.cell(r, j + 1, h); c.font = f(bold=True, color='FFFFFF'); c.fill = HDR
    r += 1
    for rule, key in [('Blend 1-3-6-12', 'blend'), ('6-month', '6m')]:
        fa, se = pk[rule]['factor'], pk[rule]['sector']
        chg = 'First month' if prev is None else ('No change' if (prev[f'{key} factor'], prev[f'{key} sector']) == (fa, se) else
              f"Sell {', '.join(x for x in (prev[f'{key} factor'], prev[f'{key} sector']) if x not in (fa, se))}; buy {', '.join(x for x in (fa, se) if x not in (prev[f'{key} factor'], prev[f'{key} sector']))}")
        vals = [rule, f'{fa} - {NAME[fa]}', f'{se} - {NAME[se]}', chg]
        for j, v in enumerate(vals):
            c = ws.cell(r, j + 1, v); c.font = f(); c.border = Border(bottom=thin)
        r += 1
    agree = pk['Blend 1-3-6-12'] == pk['6-month']
    ws.cell(r, 1, 'The two rules agree this month.' if agree else 'The two rules disagree this month. Record which you follow on the History sheet.').font = f(bold=not agree)
    r += 2
    ws.cell(r, 1, 'Protection changes apply to both rules. Protection was tested on the US market as a whole, not on this rotation (see the process document).').font = f(italic=True, size=9)
    for col, w in zip('ABCD', [22, 40, 42, 40]): ws.column_dimensions[col].width = w

    # Rankings
    rk = wb.create_sheet('Rankings')
    rk['A1'] = f'Look-back returns at the end of {at.strftime("%B %Y")} (in pounds)'; rk['A1'].font = f(bold=True, size=12)
    heads = ['Basket', 'Ticker', 'Fund', '1 month', '3 months', '6 months', '12 months', 'Blend', 'Rank (blend)', 'Rank (6 months)']
    for j, h in enumerate(heads):
        c = rk.cell(3, j + 1, h); c.font = f(bold=True, color='FFFFFF'); c.fill = HDR
    rr = 4
    for basket, lst in [('Factor', FACTOR), ('Sector', SECTOR)]:
        tick = [t for t, *_ in lst]
        rb = s.loc[tick, 'blend'].rank(ascending=False, method='first'); r6 = s.loc[tick, 6].rank(ascending=False, method='first')
        for t, n, full, isin in sorted(lst, key=lambda x: rb[x[0]]):
            vals = [basket, t, f'{n} ({full})', s.loc[t, 1], s.loc[t, 3], s.loc[t, 6], s.loc[t, 12], s.loc[t, 'blend'], int(rb[t]), int(r6[t])]
            for j, v in enumerate(vals):
                c = rk.cell(rr, j + 1, v); c.font = f(bold=(rb[t] == 1 or r6[t] == 1)); c.border = Border(bottom=thin)
                if 3 <= j <= 7: c.number_format = '0.0%'
            rr += 1
        rr += 1
    for col, w in zip('ABCDEFGHIJ', [8, 7, 60, 9, 9, 9, 9, 9, 11, 13]): rk.column_dimensions[col].width = w

    # History
    hs = wb.create_sheet('History')
    hs['A1'] = 'Picks and returns month by month (returns are what each rule would have made, including protection)'; hs['A1'].font = f(bold=True, size=12)
    cols = ['for month', 'blend factor', 'blend sector', '6m factor', '6m sector', 'protection on', 'blend return', '6m return']
    for j, h in enumerate(cols + ['Rule you followed (fill in)']):
        c = hs.cell(3, j + 1, h); c.font = f(bold=True, color='FFFFFF'); c.fill = HDR
    for i, row in hist.iterrows():
        for j, k in enumerate(cols):
            v = row.get(k, None)
            if isinstance(v, float) and np.isnan(v): v = None
            c = hs.cell(4 + i, j + 1, v if not isinstance(v, (np.bool_,)) else bool(v)); c.font = f()
            if 'return' in k: c.number_format = '0.0%'
        if str(row['for month']) in choices: hs.cell(4 + i, 9, choices[str(row['for month'])]).font = f()
    for col in 'ABCDEFGHI': hs.column_dimensions[col].width = 14

    # Protection history
    ph = wb.create_sheet('Prices used')
    ph['A1'] = 'Month-end closing prices used (London listings in pence; CEMQ and CEMS in euros; SPY in dollars)'; ph['A1'].font = f(bold=True)
    tab = p.copy(); tab.index = tab.index.astype(str)
    for j, h in enumerate(['month'] + list(tab.columns)):
        c = ph.cell(3, j + 1, h); c.font = f(bold=True, color='FFFFFF'); c.fill = HDR
    for i, (m, vals) in enumerate(tab.iterrows()):
        ph.cell(4 + i, 1, m)
        for j, v in enumerate(vals):
            if pd.notna(v): ph.cell(4 + i, 2 + j, float(v)).number_format = '0.00'
    wb.save(out_path)
    return at, pk, pr, hist


if __name__ == '__main__':
    pp = sys.argv[1] if len(sys.argv) > 1 else 'prices.csv'
    up = sys.argv[2] if len(sys.argv) > 2 else 'unemployment.csv'
    op = sys.argv[3] if len(sys.argv) > 3 else 'ETF schedule.xlsx'
    p, u = load(pp, up)
    at, pk, pr, hist = write(p, u, op)
    print('Schedule for', at + 1)
    print('Picks:', pk)
    print('Protection:', pr)
    print(hist.tail(13).to_string())
