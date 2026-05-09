#!/usr/bin/env python3
"""
Unified MFS HTML Report Generator
Supports: Nagad · bKash · Rocket (DBBL)
Auto-detects provider from file.
No transaction details list — clean summary format.
"""

import pandas as pd, re, io, sys
from datetime import datetime
from pathlib import Path

try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ══════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════
def _f(v):
    try: return float(str(v).replace(',', '').strip())
    except: return 0.0

def size_cat(a):
    if 2000 <= a <= 8000:  return 'Small (2K–8K)'
    if 8001 <= a <= 20000: return 'Medium (8K–20K)'
    if a > 20000:          return 'High (20K+)'
    return '<2K'

# ══════════════════════════════════════════════════════════
#  PROVIDER DETECTION
# ══════════════════════════════════════════════════════════
def detect(path: str):
    """Return (provider, format) from file path."""
    p = path.lower()
    ext = Path(path).suffix.lower()

    if ext == '.pdf':
        return 'rocket', 'rocket_pdf'

    # read raw Excel/CSV
    try:
        if ext == '.csv':
            raw = pd.read_csv(path, header=None, dtype=str).fillna('').values.tolist()
        else:
            raw = pd.read_excel(path, header=None, dtype=str).fillna('').values.tolist()
    except:
        return 'unknown', 'unknown'

    flat  = ' '.join(str(c) for r in raw[:12] for c in r).lower()
    hdrs  = ' '.join(str(c) for c in raw[0]).lower()

    if any(k in hdrs for k in ['txn_date_time', 'txn_type_dr_cr', 'available_blc_after_txn']):
        return 'nagad', 'nagad_standard'

    for i, row in enumerate(raw[:12]):
        rlow = ' '.join(str(c) for c in row).lower()
        if 'transaction id' in rlow and ('sender' in rlow or 'receiver' in rlow):
            for dr in raw[i+1:i+6]:
                txn = str(dr[1]).strip() if len(dr) > 1 else ''
                if re.match(r'^DC', txn, re.I):       return 'bkash', 'bkash_business'
                if re.match(r'^[0-9A-Z]{8,12}$', txn):
                    if 'nagad' in flat:                return 'nagad', 'nagad_business'
                    return 'bkash', 'bkash_business'
            if 'nagad' in flat:                        return 'nagad', 'nagad_business'
            return 'bkash', 'bkash_business'

    if 'nagad'  in flat or 'nagad'  in p: return 'nagad', 'nagad_standard'
    if 'bkash'  in flat or 'bkash'  in p: return 'bkash', 'bkash_business'
    if 'rocket' in flat or 'dbbl'   in flat: return 'rocket', 'rocket_pdf'
    return 'unknown', 'unknown'


# ══════════════════════════════════════════════════════════
#  PARSERS
# ══════════════════════════════════════════════════════════
def parse_nagad_standard(path):
    df  = pd.read_excel(path, header=0)
    df['TXN_TYPE'] = df['TXN_TYPE'].astype(str)
    main = df[~df['TXN_TYPE'].str.contains('COMMISSION|REWARD', na=False)].copy()
    comm = df[df['TXN_TYPE'].str.contains('COMMISSION|REWARD', na=False)]

    df['STATEMENT_FOR_ACC'] = df['STATEMENT_FOR_ACC'].astype(str)
    acc_s = df[df['STATEMENT_FOR_ACC'].str.match(r'^01[0-9]{9}$', na=False)]['STATEMENT_FOR_ACC']
    acc   = acc_s.iloc[0] if len(acc_s) else ''

    main['dt']          = pd.to_datetime(df.loc[main.index, 'TXN_DATE_TIME'], errors='coerce')
    main['day']         = main['dt'].dt.strftime('%Y-%m-%d')
    main['month']       = main['dt'].dt.strftime('%Y-%m')
    main['hour']        = main['dt'].dt.hour.fillna(-1).astype(int)
    main['size']        = main['TXN_AMT'].apply(size_cat)
    main['direction']   = main['TXN_TYPE_DR_CR'].astype(str).map(lambda x: 'credit' if x.upper() == 'CREDIT' else 'debit')
    main['counterpart'] = main['TXN_WITH_ACC'].astype(str).replace('nan', '')
    main['amount']      = main['TXN_AMT']
    main['type']        = main['TXN_TYPE']
    main['channel']     = main['CHANNEL'].astype(str).replace('nan', '')
    main['commission']  = 0

    comm_amt = comm['TXN_AMT'].sum()
    comm_cnt = len(comm)
    return acc, main.reset_index(drop=True), comm_amt, comm_cnt


def parse_bkash_nagad_business(path):
    raw = pd.read_excel(path, header=None, dtype=str).fillna('').values.tolist()

    # Find header row
    hi = 4
    for i, row in enumerate(raw[:12]):
        rlow = ' '.join(str(c) for c in row).lower()
        if 'transaction id' in rlow and ('sender' in rlow or 'receiver' in rlow):
            hi = i; break

    # Extract account number
    acc = ''
    for row in raw[:12]:
        for c in row:
            if re.match(r'^01[0-9]{9}$', str(c).strip()): acc = str(c).strip(); break
        if acc: break

    headers = [str(h).replace('\n', ' ').strip().lower() for h in raw[hi]]
    def col(kw): return next((i for i, h in enumerate(headers) if kw in h), -1)

    iD = col('transaction date');  iT = col('transaction id')
    iTy = col('transaction type'); iSn = col('sender')
    iRv = col('receiver');         iAm = col('transacted amount')
    iCm = col('commission');       iBl = col('balance')

    rows = []
    for row in raw[hi+1:]:
        if not row or all(str(c).strip() in ('', 'nan') for c in row): continue
        amt = _f(row[iAm]) if iAm >= 0 else 0
        if amt == 0: continue
        ty  = str(row[iTy]).strip() if iTy >= 0 else ''
        sn  = str(row[iSn]).strip() if iSn >= 0 else ''
        rv  = str(row[iRv]).strip() if iRv >= 0 else ''
        cm  = _f(row[iCm]) if iCm >= 0 else 0
        if 'cash out' in ty.lower():  dr, cp = 'credit', sn
        elif 'cash in' in ty.lower(): dr, cp = 'debit',  rv
        else:                         dr, cp = 'credit', sn or rv

        dt_str = str(row[iD]).strip() if iD >= 0 else ''
        try:
            dt = pd.to_datetime(dt_str, errors='coerce')
        except:
            dt = pd.NaT
        rows.append({
            'day':    dt.strftime('%Y-%m-%d') if pd.notna(dt) else '',
            'month':  dt.strftime('%Y-%m')    if pd.notna(dt) else '',
            'hour':   int(dt.hour)            if pd.notna(dt) else -1,
            'amount': amt, 'direction': dr, 'counterpart': cp,
            'type': ty,
            'channel': 'USSD' if 'ussd' in ty.lower() else 'API',
            'size': size_cat(amt), 'commission': cm,
        })
    main = pd.DataFrame(rows)
    comm_amt = main['commission'].sum()
    comm_cnt = len(main[main['commission'] > 0])
    return acc, main, comm_amt, comm_cnt


def parse_rocket_pdf(path):
    if not PDF_OK:
        raise RuntimeError('pdfplumber not installed. Run: pip install pdfplumber')

    def pdt(v, t=''):
        for s in [v, t]:
            s = str(s or '').strip()
            if not s: continue
            for fmt in ['%d/%m/%Y %I:%M:%S %p', '%d/%m/%Y %I:%M %p', '%d-%b-%Y']:
                try: return datetime.strptime(s, fmt)
                except: pass
        return None

    rows = []; seen = set(); acc = ''
    with pdfplumber.open(path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            if page_num == 0:
                hdr = page.extract_text() or ''
                m = re.search(r'Account No\.?\s*[:\.]?\s*(\d{9,15})', hdr, re.I)
                if m: acc = m.group(1).strip()

            tbls = page.extract_tables({'vertical_strategy': 'lines',
                                        'horizontal_strategy': 'lines',
                                        'snap_tolerance': 4})
            for tbl in (tbls or []):
                for row in tbl:
                    if not row: continue
                    row = [str(c or '').replace('\n', ' ').strip() for c in row]
                    if row[0].upper() in ('SL', 'S/L', '') or 'TXN DATE' in row[0].upper(): continue
                    try: sl = int(row[0])
                    except: continue
                    if sl in seen: continue
                    seen.add(sl)
                    if len(row) >= 10:
                        txd, vd, ty, ref, dc, am = row[1], row[2], row[5], row[6], row[7].upper(), row[8]
                    elif len(row) == 9:
                        txd, vd, ty, ref, dc, am = row[1], row[1], row[4], row[5], row[6].upper(), row[7]
                    else: continue
                    amt = _f(am)
                    if amt == 0 or 'OPENING BALANCE' in ty.upper(): continue
                    is_c = 'COMMISSION' in ty.upper()
                    dt   = pdt(vd, txd)
                    rows.append({
                        'day':    dt.strftime('%Y-%m-%d') if dt else '',
                        'month':  dt.strftime('%Y-%m')    if dt else '',
                        'hour':   dt.hour                 if dt else -1,
                        'amount': amt,
                        'direction': 'credit' if dc == 'C' else 'debit',
                        'counterpart': ref, 'type': ty,
                        'channel': 'USSD' if 'ussd' in ty.lower() else 'ROCKET',
                        'size': size_cat(amt), 'commission': 0, 'is_comm': is_c,
                    })

    df_all = pd.DataFrame(rows)
    if df_all.empty: return acc, pd.DataFrame(), 0.0, 0
    main     = df_all[~df_all['is_comm']].reset_index(drop=True)
    comm     = df_all[df_all['is_comm']]
    comm_amt = comm['amount'].sum()
    comm_cnt = len(comm)
    return acc, main, comm_amt, comm_cnt


# ══════════════════════════════════════════════════════════
#  HTML REPORT GENERATOR  (unified format, no txn list)
# ══════════════════════════════════════════════════════════
def build_report(path: str) -> str:
    provider, fmt = detect(path)
    if provider == 'unknown':
        raise ValueError('Unrecognized file format')

    if fmt == 'nagad_standard':
        acc, main, comm_amt, comm_cnt = parse_nagad_standard(path)
    elif fmt == 'rocket_pdf':
        acc, main, comm_amt, comm_cnt = parse_rocket_pdf(path)
    else:
        acc, main, comm_amt, comm_cnt = parse_bkash_nagad_business(path)

    if main.empty:
        raise ValueError('No transactions found')

    # ── Colour scheme ──
    cfg = {
        'nagad':  {'name': 'Nagad',         'icon': '🟠', 'ac': '#e85d04', 'bg': '#fff8f4', 'border': '#f5cdb0'},
        'bkash':  {'name': 'bKash',          'icon': '🔴', 'ac': '#d01060', 'bg': '#fff5f9', 'border': '#f5b3d5'},
        'rocket': {'name': 'Rocket (DBBL)',  'icon': '🟣', 'ac': '#7c3aed', 'bg': '#fdf8ff', 'border': '#c4b5fd'},
    }[provider]

    ac  = cfg['ac']
    now = datetime.now().strftime('%d %B %Y, %I:%M %p')

    cr = main[main['direction'] == 'credit']
    db = main[main['direction'] == 'debit']

    date_start = main['day'].min()
    date_end   = main['day'].max()
    daily      = main.groupby('day')['amount'].agg(['sum', 'count'])
    active_days = len(daily)
    daily_avg   = daily['sum'].mean()

    # Night activity
    night_amt = main[main['hour'].isin([22, 23, 0, 1])]['amount'].sum()
    tot_amt   = main['amount'].sum()
    night_pct = round(night_amt / tot_amt * 100, 1) if tot_amt else 0

    # ── Top 10 counterparts (credit & debit) ──
    valid_cp = main[~main['counterpart'].isin(['', 'nan', 'None', 'SYSTEM'])]
    top_recv = (valid_cp[valid_cp['direction'] == 'credit']
                .groupby('counterpart')['amount']
                .agg(['sum', 'count'])
                .sort_values('sum', ascending=False).head(10))
    top_sent = (valid_cp[valid_cp['direction'] == 'debit']
                .groupby('counterpart')['amount']
                .agg(['sum', 'count'])
                .sort_values('sum', ascending=False).head(10))
    top_all  = (valid_cp
                .groupby('counterpart')['amount']
                .agg(['sum', 'count'])
                .sort_values('sum', ascending=False).head(10))
    top_days = daily.sort_values('sum', ascending=False).head(10)

    # ── Size breakdown ──
    sizes = {
        'Small (2K–8K)':   main[main['size'] == 'Small (2K–8K)'],
        'Medium (8K–20K)': main[main['size'] == 'Medium (8K–20K)'],
        'High (20K+)':     main[main['size'] == 'High (20K+)'],
        '<2K':             main[main['size'] == '<2K'],
    }
    sz_max = max((v['amount'].sum() for v in sizes.values()), default=1)

    # ── Type breakdown ──
    tp = (main.groupby('type')['amount']
              .agg(['sum', 'count'])
              .sort_values('sum', ascending=False).head(6))
    tp_max = tp['sum'].max() if not tp.empty else 1

    # ── Channel breakdown ──
    ch = (main.groupby('channel')['amount']
              .agg(['sum', 'count'])
              .sort_values('sum', ascending=False))
    ch_max = ch['sum'].max() if not ch.empty else 1

    # ── Hourly ──
    hr_data = [0.0] * 24
    for _, r in main.iterrows():
        if 0 <= r['hour'] <= 23:
            hr_data[r['hour']] += r['amount']

    # ── Monthly ──
    mo = main.groupby('month')['amount'].agg(['sum', 'count']).reset_index()

    # ── JS data helpers ──
    def jl(s): return '[' + ','.join(f'"{v}"' for v in s) + ']'
    def jn(s): return '[' + ','.join(f'{float(v):.2f}' for v in s) + ']'
    def ji(s): return '[' + ','.join(str(int(v)) for v in s) + ']'

    mo_labels  = jl(mo['month'])
    mo_amts    = jn(mo['sum'])
    mo_cnts    = ji(mo['count'])
    hr_js      = jn(hr_data)
    hr_colors  = '[' + ','.join(f'"{"#c0392b" if h in [22,23,0,1] else ac}"' for h in range(24)) + ']'
    recv_labels = jl(top_recv.index);  recv_amts = jn(top_recv['sum'])
    sent_labels = jl(top_sent.index);  sent_amts = jn(top_sent['sum'])

    # ── HTML Table helpers ──
    def rk_tbl(df, cols):
        rows = ''
        for i, (k, v) in enumerate(df.iterrows()):
            rows += f'<tr><td>{i+1}</td><td>{k}</td>'
            for c in cols[2:]:
                val = v[c]
                rows += f'<td>{"BDT " + f"{val:,.2f}" if "amt" in c.lower() or c == "sum" else int(val)}</td>'
            rows += '</tr>'
        return rows

    cp_all_rows  = ''.join(f'<tr><td>{i+1}</td><td>{k}</td><td>{int(v["count"])}</td><td>BDT {v["sum"]:,.2f}</td></tr>' for i,(k,v) in enumerate(top_all.iterrows()))
    recv_rows    = ''.join(f'<tr><td>{i+1}</td><td>{k}</td><td>{int(v["count"])}</td><td>BDT {v["sum"]:,.2f}</td></tr>' for i,(k,v) in enumerate(top_recv.iterrows()))
    sent_rows    = ''.join(f'<tr><td>{i+1}</td><td>{k}</td><td>{int(v["count"])}</td><td>BDT {v["sum"]:,.2f}</td></tr>' for i,(k,v) in enumerate(top_sent.iterrows()))
    top_day_rows = ''.join(f'<tr><td>{i+1}</td><td>{k}</td><td>{int(v["count"])}</td><td>BDT {v["sum"]:,.2f}</td></tr>' for i,(k,v) in enumerate(top_days.iterrows()))
    mo_rows      = ''.join(f'<tr><td>{r["month"]}</td><td>{int(r["count"])}</td><td>BDT {r["sum"]:,.2f}</td></tr>' for _,r in mo.iterrows())

    def sz_bar(label, grp, color):
        amt = grp['amount'].sum(); cnt = len(grp)
        pct = round(amt / sz_max * 100, 1) if sz_max else 0
        return f'<div class="br"><div class="bl">{label}</div><div class="bt"><div class="bf" style="width:{pct}%;background:{color};"></div></div><div class="bv">BDT {amt:,.2f}</div><div class="bc">{cnt}</div></div>'

    def tp_bar(label, amt, cnt, color, max_v):
        pct = round(amt / max_v * 100, 1) if max_v else 0
        return f'<div class="br"><div class="bl sm">{label[:24]}</div><div class="bt"><div class="bf" style="width:{pct}%;background:{color};"></div></div><div class="bv">BDT {amt:,.2f}</div><div class="bc">{cnt}</div></div>'

    sz_html = (sz_bar('Small (2K–8K)',   sizes['Small (2K–8K)'],   '#1557a0') +
               sz_bar('Medium (8K–20K)', sizes['Medium (8K–20K)'], '#b86e00') +
               sz_bar('High (20K+)',     sizes['High (20K+)'],     '#c0392b') +
               sz_bar('<2K',             sizes['<2K'],             '#b0b8c8'))

    tp_colors = ['#00875a','#c0392b','#1557a0','#7c3aed','#b86e00',ac]
    tp_html   = ''.join(tp_bar(k, v['sum'], int(v['count']), tp_colors[i%6], tp_max)
                        for i,(k,v) in enumerate(tp.iterrows()))

    ch_colors = [ac, '#1557a0', '#00875a', '#b86e00']
    ch_html   = ''.join(tp_bar(k, v['sum'], int(v['count']), ch_colors[i%4], ch_max)
                        for i,(k,v) in enumerate(ch.iterrows()))

    warn_recv = (f'<div class="warn">⚠️ <strong>{top_recv.index[0]}</strong> — {int(top_recv.iloc[0]["count"])} transactions. Unusually high frequency.</div>'
                 if not top_recv.empty and int(top_recv.iloc[0]['count']) > 30 else '')
    warn_sent = (f'<div class="warn">⚠️ {int(top_sent.iloc[0]["count"])} transactions to <strong>{top_sent.index[0]}</strong>. Review recommended.</div>'
                 if not top_sent.empty and int(top_sent.iloc[0]['count']) > 30 else '')

    fname = Path(path).name

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>{cfg['name']} Report — {acc}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
:root{{--bg:#f4f6f9;--bg2:#eef1f6;--card:#fff;--border:#dde3ed;--green:#00875a;--red:#c0392b;--amber:#b86e00;--blue:#1557a0;--text:#1a2332;--muted:#6b7f99;--ac:{ac};}}
*{{margin:0;padding:0;box-sizing:border-box;}}
body{{background:var(--bg);color:var(--text);font-family:'Inter',sans-serif;font-size:14px;}}
/* TOP BAR */
.topbar{{background:#fff;border-bottom:3px solid var(--ac);padding:18px 36px;display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;box-shadow:0 2px 8px rgba(0,0,0,.06);}}
.tb-logo{{width:46px;height:46px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:22px;background:linear-gradient(135deg,{ac}bb,{ac});}}
.tb-name{{font-size:22px;font-weight:800;color:var(--ac);}}
.tb-sub{{font-size:11px;color:var(--muted);letter-spacing:1.5px;text-transform:uppercase;margin-top:1px;}}
.tb-acc{{font-family:'JetBrains Mono',monospace;font-size:17px;font-weight:700;letter-spacing:2px;}}
.tb-period{{font-size:12px;color:var(--muted);margin-top:3px;font-family:'JetBrains Mono',monospace;}}
/* PAGE */
.page{{max-width:1260px;margin:0 auto;padding:24px 22px 60px;}}
/* INFOBOX */
.infobox{{background:{cfg['bg']};border:1px solid {cfg['border']};border-left:4px solid var(--ac);border-radius:8px;padding:13px 18px;font-size:13px;line-height:1.9;margin-bottom:20px;}}
.infobox .title{{font-size:10px;text-transform:uppercase;letter-spacing:1.5px;color:var(--ac);font-weight:700;margin-bottom:6px;}}
/* METRIC CARDS */
.strip{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:22px;}}
@media(max-width:860px){{.strip{{grid-template-columns:repeat(2,1fr);}}}}
.sc{{background:#fff;border:1px solid var(--border);border-radius:10px;padding:17px;position:relative;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.05);}}
.sc::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;}}
.sc.bl::before{{background:var(--blue);}}.sc.gr::before{{background:var(--green);}}.sc.rd::before{{background:var(--red);}}.sc.am::before{{background:var(--amber);}}
.sc .ic{{font-size:19px;margin-bottom:7px;}}
.sc .lb{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:1.2px;margin-bottom:4px;}}
.sc .vl{{font-size:21px;font-weight:800;font-family:'JetBrains Mono',monospace;line-height:1;}}
.sc .vl.bl{{color:var(--blue);}}.sc .vl.gr{{color:var(--green);}}.sc .vl.rd{{color:var(--red);}}.sc .vl.am{{color:var(--amber);}}
.sc .sb{{font-size:11px;color:var(--muted);margin-top:4px;font-family:'JetBrains Mono',monospace;}}
/* LAYOUT */
.two{{display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;}}
.three{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:15px;margin-bottom:15px;}}
@media(max-width:780px){{.two,.three{{grid-template-columns:1fr;}}}}
.panel{{background:#fff;border:1px solid var(--border);border-radius:10px;padding:19px 20px;box-shadow:0 1px 3px rgba(0,0,0,.04);}}
.pt{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:var(--muted);margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid var(--bg);display:flex;align-items:center;gap:7px;}}
.pt::before{{content:'';width:3px;height:13px;background:var(--ac);border-radius:2px;display:inline-block;}}
/* BAR ROWS */
.br{{display:flex;align-items:center;gap:9px;margin-bottom:10px;}}
.bl{{min-width:155px;font-size:12px;color:var(--text);}}.bl.sm{{min-width:120px;}}
.bt{{flex:1;height:6px;background:var(--bg2);border-radius:3px;overflow:hidden;}}
.bf{{height:100%;border-radius:3px;}}
.bv{{min-width:90px;text-align:right;font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--muted);}}
.bc{{min-width:38px;text-align:right;font-size:10px;color:var(--muted);}}
/* TABLE */
.tbl{{width:100%;border-collapse:collapse;font-size:12px;}}
.tbl thead tr{{background:var(--bg2);}}
.tbl th{{padding:8px 10px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);font-weight:700;border-bottom:1px solid var(--border);white-space:nowrap;}}
.tbl td{{padding:8px 10px;border-bottom:1px solid var(--bg);font-family:'JetBrains Mono',monospace;font-size:11px;}}
.tbl tr:hover td{{background:var(--bg2);}}
.rk{{display:inline-flex;width:18px;height:18px;background:var(--bg2);border:1px solid var(--border);border-radius:3px;align-items:center;justify-content:center;font-size:9px;color:var(--muted);margin-right:4px;}}
/* TOP RECV/SENT BADGES */
.sec-badge{{display:inline-block;padding:3px 10px;border-radius:5px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:12px;}}
.badge-recv{{background:#e6f6f0;border:1px solid #b3dcc8;color:#00875a;}}
.badge-sent{{background:#fde8ec;border:1px solid #f5b8c4;color:#c0392b;}}
/* WARN */
.warn{{background:#fffbec;border:1px solid #edd890;border-radius:6px;padding:9px 13px;font-size:12px;color:#7a5200;margin-top:8px;}}
/* STAT MINI GRID */
.kv{{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:12px;}}
.kv-i{{background:var(--bg);border-radius:6px;padding:9px 11px;border:1px solid var(--border);}}
.kv-k{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:1px;margin-bottom:2px;}}
.kv-v{{font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:700;}}
/* FOOTER */
.footer{{text-align:center;padding:15px;font-size:11px;color:var(--muted);border-top:1px solid var(--border);margin-top:24px;background:#fff;}}
@media print{{body{{background:#fff;}}.topbar{{position:static;}}canvas{{max-width:100%;}}}}
</style>
</head>
<body>

<!-- TOP BAR -->
<div class="topbar">
  <div style="display:flex;align-items:center;gap:13px;">
    <div class="tb-logo">{cfg['icon']}</div>
    <div>
      <div class="tb-name">{cfg['name']}</div>
      <div class="tb-sub">Mobile Financial Service · Statement Analysis</div>
    </div>
  </div>
  <div style="text-align:right;">
    <div class="tb-acc">{acc or '—'}</div>
    <div class="tb-period">{date_start} → {date_end} &nbsp;·&nbsp; Active {active_days} days</div>
  </div>
</div>

<div class="page">

<!-- INFOBOX -->
<div class="infobox">
  <div class="title">📋 Analysis Summary</div>
  <strong>{cfg['name']}</strong> account <strong>{acc}</strong> recorded
  <strong>{len(main):,} transactions</strong> over <strong>{active_days} active days</strong>
  ({date_start} → {date_end}).
  Total Credit (In): <strong style="color:var(--green)">BDT {cr['amount'].sum():,.2f}</strong> ({len(cr):,} txns) &nbsp;|&nbsp;
  Total Debit (Out): <strong style="color:var(--red)">BDT {db['amount'].sum():,.2f}</strong> ({len(db):,} txns).
  Commission Earned: <strong style="color:var(--amber)">BDT {comm_amt:,.2f}</strong>.
  Daily Average: <strong>BDT {daily_avg:,.2f}</strong>.
  Night-time activity (10PM–2AM): <strong style="color:{"var(--red)" if night_pct>35 else "var(--amber)"}">
  {night_pct}%</strong>.{'&nbsp; ⚠️ <strong>High night activity — suspicious.</strong>' if night_pct>35 else ''}
</div>

<!-- METRIC CARDS -->
<div class="strip">
  <div class="sc bl"><div class="ic">🔢</div><div class="lb">Total Transactions</div><div class="vl bl">{len(main):,}</div><div class="sb">Excl. commission</div></div>
  <div class="sc gr"><div class="ic">⬇️</div><div class="lb">Total Credit (In)</div><div class="vl gr">BDT {cr['amount'].sum():,.2f}</div><div class="sb">{len(cr):,} transactions</div></div>
  <div class="sc rd"><div class="ic">⬆️</div><div class="lb">Total Debit (Out)</div><div class="vl rd">BDT {db['amount'].sum():,.2f}</div><div class="sb">{len(db):,} transactions</div></div>
  <div class="sc am"><div class="ic">💰</div><div class="lb">Commission Earned</div><div class="vl am">BDT {comm_amt:,.2f}</div><div class="sb">{comm_cnt:,} entries</div></div>
</div>

<!-- TREND CHARTS -->
<div class="two">
  <div class="panel"><div class="pt">Monthly Transaction Volume</div><div style="height:240px;"><canvas id="cMonthly"></canvas></div></div>
  <div class="panel"><div class="pt">Hourly Activity Distribution (Red = Night 10PM–2AM)</div><div style="height:240px;"><canvas id="cHourly"></canvas></div></div>
</div>

<!-- SIZE / TYPE / CHANNEL -->
<div class="three">
  <div class="panel">
    <div class="pt">Transaction Size Breakdown</div>
    {sz_html}
  </div>
  <div class="panel">
    <div class="pt">Transaction Type Breakdown</div>
    {tp_html}
  </div>
  <div class="panel">
    <div class="pt">Channel & Daily Stats</div>
    {ch_html}
    <div class="kv">
      <div class="kv-i"><div class="kv-k">Daily Average</div><div class="kv-v">BDT {daily_avg:,.0f}</div></div>
      <div class="kv-i"><div class="kv-k">Avg Txns/Day</div><div class="kv-v">{daily['count'].mean():.0f}</div></div>
      <div class="kv-i"><div class="kv-k">Peak Day</div><div class="kv-v" style="color:var(--green)">BDT {daily['sum'].max():,.0f}</div></div>
      <div class="kv-i"><div class="kv-k">Night Activity</div><div class="kv-v" style="color:var(--amber)">{night_pct}%</div></div>
    </div>
  </div>
</div>

<!-- TOP 10 RECEIVED FROM -->
<div class="panel" style="margin-bottom:15px;">
  <div class="sec-badge badge-recv">💚 Top 10 — Money Received From</div>
  <div class="pt">Top 10 Account Numbers — Highest Amount Received From (Credit)</div>
  <div class="two" style="margin-bottom:0;">
    <div style="height:310px;"><canvas id="cRecv"></canvas></div>
    <div style="overflow-x:auto;">
      <table class="tbl">
        <thead><tr><th>#</th><th>Account Number</th><th>Transactions</th><th>Total Received</th></tr></thead>
        <tbody>{recv_rows}</tbody>
      </table>
      {warn_recv}
    </div>
  </div>
</div>

<!-- TOP 10 SENT TO -->
<div class="panel" style="margin-bottom:15px;">
  <div class="sec-badge badge-sent">🔴 Top 10 — Money Sent To</div>
  <div class="pt">Top 10 Account Numbers — Highest Amount Sent To (Debit)</div>
  <div class="two" style="margin-bottom:0;">
    <div style="height:310px;"><canvas id="cSent"></canvas></div>
    <div style="overflow-x:auto;">
      <table class="tbl">
        <thead><tr><th>#</th><th>Account Number</th><th>Transactions</th><th>Total Sent</th></tr></thead>
        <tbody>{sent_rows}</tbody>
      </table>
      {warn_sent}
    </div>
  </div>
</div>

<!-- TOP COUNTERPARTS + TOP DAYS -->
<div class="two">
  <div class="panel">
    <div class="pt">Top 10 Counterpart Accounts (All)</div>
    <table class="tbl">
      <thead><tr><th>#</th><th>Account</th><th>Transactions</th><th>Total Amount</th></tr></thead>
      <tbody>{cp_all_rows}</tbody>
    </table>
  </div>
  <div class="panel">
    <div class="pt">Monthly Breakdown + Top 10 Busiest Days</div>
    <table class="tbl" style="margin-bottom:14px;">
      <thead><tr><th>Month</th><th>Txns</th><th>Total Amount</th></tr></thead>
      <tbody>{mo_rows}</tbody>
    </table>
    <div class="pt" style="margin-top:12px;">Top 10 Busiest Days</div>
    <table class="tbl">
      <thead><tr><th>#</th><th>Date</th><th>Txns</th><th>Total Amount</th></tr></thead>
      <tbody>{top_day_rows}</tbody>
    </table>
  </div>
</div>

</div><!-- /page -->

<div class="footer">
  {cfg['name']} Statement Analysis &nbsp;·&nbsp; Account: {acc} &nbsp;·&nbsp; File: {fname} &nbsp;·&nbsp; Generated: {now}
</div>

<script>
Chart.defaults.font.family="'JetBrains Mono',sans-serif";
Chart.defaults.font.size=10;
Chart.defaults.color='#6b7f99';
const AC='{ac}';

// Monthly
new Chart(document.getElementById('cMonthly'),{{type:'bar',data:{{labels:{mo_labels},datasets:[
  {{label:'Amount',data:{mo_amts},backgroundColor:AC+'aa',borderColor:AC,borderWidth:1,borderRadius:4,yAxisID:'y'}},
  {{label:'Txns',data:{mo_cnts},type:'line',borderColor:'#c0392b',backgroundColor:'#c0392b18',borderWidth:2,tension:.4,pointRadius:3,pointBackgroundColor:'#c0392b',yAxisID:'y2'}}
]}},options:{{responsive:true,maintainAspectRatio:false,
  plugins:{{legend:{{position:'top',labels:{{font:{{size:10}},padding:8}}}}}},
  scales:{{y:{{ticks:{{callback:v=>'BDT '+(v>=100000?(v/100000).toFixed(1)+'L':(v/1000).toFixed(0)+'K')}},grid:{{color:'#eef1f6'}}}},
           y2:{{ticks:{{callback:v=>v+'txn'}},grid:{{display:false}},position:'right'}},
           x:{{grid:{{display:false}},ticks:{{maxTicksLimit:12}}}}}}}}}});

// Hourly
const hh={hr_js};
new Chart(document.getElementById('cHourly'),{{type:'bar',data:{{labels:Array.from({{length:24}},(_,i)=>i+':00'),datasets:[{{data:hh,backgroundColor:{hr_colors},borderWidth:0,borderRadius:3}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>'BDT '+parseFloat(ctx.raw).toLocaleString()}}}}}},scales:{{y:{{ticks:{{callback:v=>'BDT '+(v>=100000?(v/100000).toFixed(1)+'L':(v/1000).toFixed(0)+'K')}},grid:{{color:'#eef1f6'}}}},x:{{grid:{{display:false}}}}}}}}}});

// Recv
new Chart(document.getElementById('cRecv'),{{type:'bar',data:{{labels:{recv_labels},datasets:[{{label:'Received (BDT)',data:{recv_amts},backgroundColor:'#00875a99',borderColor:'#00875a',borderWidth:1,borderRadius:4}}]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' BDT '+parseFloat(ctx.raw).toLocaleString('en-BD',{{minimumFractionDigits:2}})}}}}}},scales:{{x:{{ticks:{{callback:v=>'BDT '+(v/1000).toFixed(0)+'K'}},grid:{{color:'#eef1f6'}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10,family:'monospace'}}}}}}}}}}}});

// Sent
new Chart(document.getElementById('cSent'),{{type:'bar',data:{{labels:{sent_labels},datasets:[{{label:'Sent (BDT)',data:{sent_amts},backgroundColor:'#c0392b99',borderColor:'#c0392b',borderWidth:1,borderRadius:4}}]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>' BDT '+parseFloat(ctx.raw).toLocaleString('en-BD',{{minimumFractionDigits:2}})}}}}}},scales:{{x:{{ticks:{{callback:v=>'BDT '+(v/1000).toFixed(0)+'K'}},grid:{{color:'#eef1f6'}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10,family:'monospace'}}}}}}}}}}}});
</script>
</body></html>"""


# ══════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════
if __name__ == '__main__':
    files = sys.argv[1:] or []
    if not files:
        print("Usage: python3 generate_report.py <file1> [file2] ...")
        sys.exit(1)

    for path in files:
        try:
            html = build_report(path)
            stem = Path(path).stem
            out  = f"{stem}_mfs_report.html"
            with open(out, 'w', encoding='utf-8') as f:
                f.write(html)
            print(f"✅ {Path(path).name} → {out}")
        except Exception as e:
            print(f"❌ {path}: {e}")
