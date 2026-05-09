"""
Mobile Financial System Analyzer — Bangladesh
Supports: bKash · Nagad · Rocket
Output: HTML Report + Word (docx) Report
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re, io, json
from datetime import datetime

# ─── optional word export ───────────────────────────────────────────────────
try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# pdf parser
try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ════════════════════════════════════════════════════════════════════════════
#  PAGE CONFIG
# ════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Mobile Financial System Analyzer",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ════════════════════════════════════════════════════════════════════════════
#  GLOBAL CSS
# ════════════════════════════════════════════════════════════════════════════
# Font import (separate from CSS to avoid conflicts)
st.markdown(
    '<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Bengali:wght@400;600;700&family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">',
    unsafe_allow_html=True
)

# CSS injected cleanly without special unicode characters in comments
_CSS = """
<style>
html,body,[class*="css"]{font-family:'Inter','Noto Sans Bengali',sans-serif!important;background:#fff!important;}
.stApp{background:#fff!important;}
.main .block-container{background:#fff;padding:0 0 40px 0;max-width:100%;}
.hero{background:linear-gradient(135deg,#1a0533 0%,#2d0a5e 40%,#8b1a6b 75%,#e2136e 100%);padding:48px 48px 40px;border-radius:0;position:relative;overflow:hidden;}
.hero::before{content:'';position:absolute;inset:0;background:radial-gradient(circle at 80% 30%,rgba(255,255,255,0.06) 0,transparent 40%);pointer-events:none;}
.hero h1{color:#fff;font-size:36px;font-weight:800;margin:0 0 10px 0;line-height:1.2;}
.hero p{color:rgba(255,255,255,.82);font-size:14px;margin:0 0 22px 0;max-width:540px;line-height:1.6;}
.hero-badges{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:22px;}
.hero-badge{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);color:#fff;padding:6px 16px;border-radius:8px;font-size:13px;font-weight:600;display:flex;align-items:center;gap:7px;}
.feature-pills{display:flex;gap:10px;flex-wrap:wrap;}
.fpill{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:rgba(255,255,255,.85);padding:6px 13px;border-radius:7px;font-size:12px;display:flex;align-items:center;gap:6px;}
.sec-wrap{padding:24px 32px 10px;}
.sec-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:2px;color:#6b7f99;border-bottom:2px solid #f0f4f8;padding-bottom:8px;margin-bottom:14px;}
.prov-nagad{background:linear-gradient(135deg,#fff8f4,#fff3ec);border:1px solid #f5cdb0;border-left:5px solid #e85d04;border-radius:10px;padding:16px 20px;margin:0 32px 18px;}
.prov-bkash{background:linear-gradient(135deg,#fff5f9,#fff0f6);border:1px solid #f5b3d5;border-left:5px solid #e2136e;border-radius:10px;padding:16px 20px;margin:0 32px 18px;}
.prov-rocket{background:linear-gradient(135deg,#f5f0ff,#ede8ff);border:1px solid #c4b3f0;border-left:5px solid #7c3aed;border-radius:10px;padding:16px 20px;margin:0 32px 18px;}
[data-testid="metric-container"]{background:#fff!important;border:1px solid #e8ecf4!important;border-radius:12px!important;padding:16px 18px!important;box-shadow:0 2px 8px rgba(0,0,0,.05)!important;}
[data-testid="metric-container"] label{font-size:11px!important;color:#6b7f99!important;text-transform:uppercase!important;letter-spacing:1px!important;}
[data-testid="metric-container"] [data-testid="stMetricValue"]{font-size:22px!important;font-weight:800!important;color:#1a2332!important;}
.sum-box{background:#f0faf6;border:1px solid #b3dcc8;border-radius:10px;padding:14px 20px;font-size:13px;line-height:1.9;margin:0 32px 14px;}
.warn-box{background:#fffbec;border:1px solid #edd890;border-radius:7px;padding:10px 14px;font-size:12px;color:#7a5200;margin-top:8px;}
.upload-card{background:#fff;border:1px solid #e8ecf4;border-radius:14px;padding:22px;box-shadow:0 2px 10px rgba(0,0,0,.05);}
.upload-card h3{font-size:15px;font-weight:700;color:#1a2332;margin:0 0 5px;}
.upload-card p{font-size:13px;color:#6b7f99;margin:0 0 14px;}
.wyg-item{display:flex;align-items:center;gap:8px;padding:8px 0;border-bottom:1px solid #f0f4f8;font-size:13px;color:#2d3748;}
.wyg-icon{width:20px;text-align:center;}
.preview-card{background:#fff;border:1px solid #e8ecf4;border-radius:10px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.04);}
.pc-label{font-size:11px;color:#6b7f99;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;}
.pc-value{font-size:22px;font-weight:800;}
.pc-green{color:#00875a;}.pc-red{color:#c0392b;}.pc-blue{color:#1557a0;}.pc-purple{color:#7c3aed;}
.stDataFrame{border:1px solid #e8ecf4;border-radius:8px;overflow:hidden;}
.stDownloadButton button{background:linear-gradient(135deg,#e2136e,#b8005a)!important;color:#fff!important;border:none!important;border-radius:8px!important;font-weight:700!important;padding:10px 20px!important;width:100%!important;}
#MainMenu,footer,header{visibility:hidden;}
.stDeployButton{display:none;}
[data-testid="stFileUploadDropzone"]{background:#fafbff!important;border:2px dashed #c8d0e0!important;border-radius:12px!important;}
</style>
"""
st.markdown(_CSS, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════════════════
def _f(v):
    try: return float(str(v).replace(',','').strip())
    except: return 0.0

def _date(v):
    if hasattr(v,'isoformat'): return str(v)[:19]
    s = str(v).strip()
    return s[:19] if s not in ('nan','None','','NaT') else ''

def size_cat(a):
    if 2000<=a<=8000:  return 'ছোট (২-৮হা)'
    if 8001<=a<=20000: return 'মধ্যম (৮-২০হা)'
    if a>20000:        return 'উচ্চ (২০হা+)'
    return '<২হা'

def bdt(n):
    try: return f"৳{float(n):,.2f}"
    except: return "৳0.00"

def fmtn(n):
    try: return f"{int(n):,}"
    except: return "0"

ACCENT = {'nagad':'#e85d04','bkash':'#e2136e','rocket':'#7c3aed','unknown':'#6b7f99'}


# ════════════════════════════════════════════════════════════════════════════
#  DETECTION
# ════════════════════════════════════════════════════════════════════════════
def detect_provider(raw, fname=''):
    fn   = fname.lower()
    flat = ' '.join(str(c) for r in raw[:12] for c in r).lower()
    hdrs = ' '.join(str(c) for c in raw[0]).lower() if raw else ''

    if any(k in hdrs for k in ['txn_date_time','txn_type_dr_cr','available_blc_after_txn']):
        return 'nagad','nagad_standard'

    for i,row in enumerate(raw[:12]):
        rlow = ' '.join(str(c) for c in row).lower()
        if 'transaction id' in rlow and ('sender' in rlow or 'receiver' in rlow):
            for dr in raw[i+1:i+6]:
                txn = str(dr[1]).strip() if len(dr)>1 else ''
                if re.match(r'^DC',txn,re.I): return 'bkash','bkash_business'
                if re.match(r'^[0-9A-Z]{8,12}$',txn):
                    if 'nagad' in flat or 'nagad' in fn: return 'nagad','nagad_business'
                    return 'bkash','bkash_business'
            if 'nagad' in flat or 'nagad' in fn: return 'nagad','nagad_business'
            return 'bkash','bkash_business'

    if 'rocket' in flat or 'rocket' in fn or 'dbbl' in flat: return 'rocket','rocket'
    if 'nagad'  in flat or 'nagad'  in fn: return 'nagad','nagad_standard'
    if 'bkash'  in flat or 'bkash'  in fn: return 'bkash','bkash_business'
    return 'unknown','unknown'


# ════════════════════════════════════════════════════════════════════════════
#  PARSERS
# ════════════════════════════════════════════════════════════════════════════
def parse_nagad_standard(raw):
    hi=0
    for i,row in enumerate(raw[:6]):
        if any(str(c).upper() in ('TXN_DATE_TIME','TXN_ID','TXN_TYPE_DR_CR') for c in row):
            hi=i; break
    H=[str(h).strip().upper() for h in raw[hi]]
    def c(n): return H.index(n) if n in H else -1
    iD=c('TXN_DATE_TIME');iT=c('TXN_ID');iTy=c('TXN_TYPE')
    iAF=c('STATEMENT_FOR_ACC');iAT=c('TXN_WITH_ACC')
    iDC=c('TXN_TYPE_DR_CR');iA=c('TXN_AMT');iB=c('AVAILABLE_BLC_AFTER_TXN')
    iCh=c('CHANNEL');iSt=c('STATUS')

    acc,rows='',[]
    for row in raw[hi+1:]:
        if not row or all(str(x).strip() in ('','nan') for x in row): continue
        amt=_f(row[iA]) if iA>=0 else 0
        if amt==0: continue
        ty=str(row[iTy]).strip() if iTy>=0 else ''
        dc=str(row[iDC]).strip().upper() if iDC>=0 else ''
        ic=any(k in ty.upper() for k in ['COMMISSION','REWARD'])
        av=str(row[iAF]).strip() if iAF>=0 else ''
        if not acc and re.match(r'^01[0-9]{9}$',av): acc=av
        dt=_date(row[iD] if iD>=0 else '')
        rows.append({
            'তারিখ':dt[:16],'TXN_ID':str(row[iT]).strip() if iT>=0 else '',
            'ধরন':ty,'direction':'credit' if dc=='CREDIT' else 'debit',
            'দিক':'গ্রহণ' if dc=='CREDIT' else 'পাঠানো',
            'পরিমাণ':amt,'ব্যালেন্স':_f(row[iB]) if iB>=0 else 0,
            'চ্যানেল':str(row[iCh]).strip() if iCh>=0 else '',
            'কাউন্টারপার্ট':str(row[iAT]).strip() if iAT>=0 else '',
            'স্ট্যাটাস':str(row[iSt]).strip() if iSt>=0 else '',
            'is_comm':ic,'আকার':size_cat(amt),
            'দিন':dt[:10],'মাস':dt[:7],
            'কমিশন':0,
        })
    return acc,rows

def parse_bkash_nagad_business(raw):
    """
    bKash Statement Parser (Business Agent OR Customer).
    Direction logic:
      - Sender == this account  → DEBIT  (account sent money out)
      - Receiver == this account → CREDIT (account received money)
      - Neither matches          → fallback heuristic
    """
    hi, acc = -1, ''
    for i, row in enumerate(raw[:12]):
        for c in row:
            m = re.match(r'^01[0-9]{9}$', str(c).strip())
            if m: acc = str(c).strip()
        rlow = ' '.join(str(c) for c in row).lower()
        if 'transaction id' in rlow and ('sender' in rlow or 'transaction type' in rlow):
            hi = i; break
    if hi < 0: hi = 4

    H = [str(h).replace('\n', ' ').strip().lower() for h in raw[hi]]
    def col(kw): return next((i for i, h in enumerate(H) if kw in h), -1)

    iDt = col('transaction date'); iT  = col('transaction id')
    iTy = col('transaction type'); iSn = col('sender')
    iRv = col('receiver');         iA  = col('transacted amount')
    iCm = col('commission');       iFe = col('fee')
    iBl = col('balance')

    acc_clean = re.sub(r'\s', '', acc)  # strip spaces for comparison

    rows = []
    for row in raw[hi+1:]:
        if not row or all(str(x).strip() in ('', 'nan', 'None') for x in row): continue
        amt = _f(row[iA]) if iA >= 0 else 0
        if amt == 0: continue

        ty = str(row[iTy]).strip() if iTy >= 0 else ''

        # Clean sender / receiver (remove spaces, strip)
        sn = re.sub(r'\s', '', str(row[iSn]).strip()) if iSn >= 0 else ''
        rv = re.sub(r'\s', '', str(row[iRv]).strip()) if iRv >= 0 else ''

        # ── Direction: Sender/Receiver vs account ──────────────────────
        if sn == acc_clean:
            dr = 'debit'    # this account sent money
            cp = rv if rv not in ('', 'nan', 'NaN', 'None') else ''
        elif rv == acc_clean:
            dr = 'credit'   # this account received money
            cp = sn if sn not in ('', 'nan', 'NaN', 'None') else ''
        else:
            # Neither matches — last resort: check type keywords
            ty_low = ty.lower()
            if any(k in ty_low for k in ['cash out', 'payment', 'topup', 'bill pay',
                                          'loan repay', 'subscribe', 'direct pay']):
                dr, cp = 'debit', rv or sn
            else:
                dr, cp = 'credit', sn or rv

        cm = _f(row[iCm]) if iCm >= 0 else 0
        # Customer statements have Fee instead of Commission
        if cm == 0 and iFe >= 0:
            cm = _f(row[iFe])

        # ── Date: parse properly to handle AM/PM ──────────────────────
        raw_dt = str(row[iDt]).strip() if iDt >= 0 else ''
        try:
            dt_obj = pd.to_datetime(raw_dt, errors='coerce', dayfirst=True)
            if pd.isna(dt_obj):
                date_s = raw_dt[:16] if raw_dt else ''
                day_s  = ''
                mon_s  = ''
                hour   = -1
            else:
                date_s = dt_obj.strftime('%Y-%m-%d %H:%M')
                day_s  = dt_obj.strftime('%Y-%m-%d')
                mon_s  = dt_obj.strftime('%Y-%m')
                hour   = dt_obj.hour
        except:
            date_s = raw_dt[:16]; day_s = ''; mon_s = ''; hour = -1

        rows.append({
            'তারিখ':       date_s,
            'TXN_ID':       str(row[iT]).strip() if iT >= 0 else '',
            'ধরন':          ty,
            'direction':    dr,
            'দিক':          'Credit (In)' if dr == 'credit' else 'Debit (Out)',
            'পরিমাণ':       amt,
            'ব্যালেন্স':    _f(row[iBl]) if iBl >= 0 else 0,
            'চ্যানেল':      'USSD' if 'ussd' in ty.lower() else 'API',
            'কাউন্টারপার্ট': cp,
            'স্ট্যাটাস':   'APPROVED',
            'is_comm':      False,
            'আকার':         size_cat(amt),
            'দিন':          day_s,
            'মাস':          mon_s,
            'hour':         hour,
            'কমিশন':        cm,
        })
    return acc, rows


# ════════════════════════════════════════════════════════════════════════════
#  ANALYSIS
# ════════════════════════════════════════════════════════════════════════════

# ════════════════════════════════════════════════════════════════════════════
#  ROCKET (DBBL) PDF PARSER
# ════════════════════════════════════════════════════════════════════════════
def detect_rocket_pdf(file_bytes):
    """Quick check: is this a DBBL Rocket PDF?"""
    if not PDF_OK:
        return False
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            text = (pdf.pages[0].extract_text() or "").lower()
            return any(k in text for k in ["dutch-bangla","dbbl","mobile banking","rocket"])
    except:
        return False


def parse_rocket_pdf(file_bytes):
    """
    Parse DBBL Rocket Mobile Banking PDF statement.
    
    Columns: SL | Txn Date | Value Date | Txn ID | Br.Code | Txn Type | Ref A/C | Dr/Cr | Txn Amount | Running Balance
    Dr/Cr: C = Credit (received), D = Debit (sent)
    """
    if not PDF_OK:
        return "", []

    import pdfplumber

    account = ""
    title   = ""
    rows    = []
    seen_sl = set()

    def parse_dt(val_date, txn_date):
        """Convert Value Date (DD/MM/YYYY HH:MM:SS AM/PM) to ISO string."""
        for v in [val_date, txn_date]:
            v = str(v or "").strip()
            if not v:
                continue
            # DD/MM/YYYY HH:MM:SS AM/PM
            for fmt in ["%d/%m/%Y %I:%M:%S %p", "%d/%m/%Y %I:%M %p",
                        "%d-%b-%Y %I:%M:%S %p", "%d-%b-%Y"]:
                try:
                    return datetime.strptime(v, fmt).strftime("%Y-%m-%d %H:%M:%S")
                except:
                    pass
            # fallback: return as-is
            return v
        return ""

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages):

            # --- Extract account info from first page header text ---
            if page_num == 0:
                hdr_text = page.extract_text() or ""
                m = re.search(r"Account No\.?\s*[:\.]?\s*(\d{9,15})", hdr_text, re.I)
                if m:
                    account = m.group(1).strip()
                mt = re.search(r"Account Title\s*[:\.]?\s*(.+)", hdr_text, re.I)
                if mt:
                    title = mt.group(1).strip().split("\n")[0]

            # --- Extract tables ---
            tables = page.extract_tables({
                "vertical_strategy":   "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance":       4,
            })

            for table in (tables or []):
                for row in table:
                    if not row:
                        continue
                    # Clean cells
                    row = [str(c or "").replace("\n", " ").strip() for c in row]

                    # Skip header rows
                    first = row[0].upper()
                    if first in ("SL", "S/L", "") or "TXN DATE" in first.upper():
                        continue

                    # Must have numeric SL
                    try:
                        sl = int(row[0])
                    except:
                        continue

                    # Skip duplicates (same SL appears on repeated page headers)
                    if sl in seen_sl:
                        continue
                    seen_sl.add(sl)

                    # Flexible column mapping (10 or 9 cols)
                    if len(row) >= 10:
                        txn_date, val_date = row[1], row[2]
                        txn_id   = row[3]
                        txn_type = row[5]
                        ref_ac   = row[6]
                        dr_cr    = row[7].upper()
                        amt_str  = row[8]
                        bal_str  = row[9]
                    elif len(row) == 9:
                        txn_date, val_date = row[1], row[1]
                        txn_id   = row[2]
                        txn_type = row[4]
                        ref_ac   = row[5]
                        dr_cr    = row[6].upper()
                        amt_str  = row[7]
                        bal_str  = row[8]
                    else:
                        continue

                    amt = _f(amt_str)
                    if amt == 0:
                        continue

                    # Skip "Opening Balance" row
                    if "OPENING BALANCE" in txn_type.upper():
                        continue

                    direction = "credit" if dr_cr == "C" else "debit"
                    is_comm   = "COMMISSION" in txn_type.upper()
                    dt        = parse_dt(val_date, txn_date)

                    # Ref A/C = counterpart mobile number
                    cp = ref_ac if re.match(r"^0[0-9]{10,12}$", ref_ac) else ref_ac

                    rows.append({
                        "তারিখ":      dt[:16],
                        "TXN_ID":                               txn_id,
                        "ধরন":       txn_type,
                        "direction":                            direction,
                        "দিক":      "গ্রহণ" if direction=="credit" else "পাঠানো",
                        "পরিমাণ":  amt,
                        "ব্যালেন্স": _f(bal_str),
                        "চ্যানেল": "ROCKET",
                        "কাউন্টারপার্ট": cp,
                        "স্ট্যাটাস":  "APPROVED",
                        "is_comm":                              is_comm,
                        "আকার":       size_cat(amt),
                        "দিন":         dt[:10],
                        "মাস":         dt[:7],
                        "কমিশন":    0,
                    })

    return account, rows

def analyse(rows):
    df=pd.DataFrame(rows)
    if df.empty: return None,None
    main=df[~df['is_comm']]
    comm=df[df['is_comm']]
    cr=main[main['direction']=='credit']
    db=main[main['direction']=='debit']
    daily=main.groupby('দিন')['পরিমাণ'].agg(['sum','count'])

    try:
        main2=main.copy()
        main2['ঘণ্টা']=pd.to_datetime(main2['তারিখ'],errors='coerce').dt.hour.fillna(0).astype(int)
        night=main2[main2['ঘণ্টা'].isin([22,23,0,1])]['পরিমাণ'].sum()
        tot=main2['পরিমাণ'].sum()
        npct=round(night/tot*100,1) if tot else 0
    except: npct=0

    s={
        'total':len(main),'cr_amt':cr['পরিমাণ'].sum(),'cr_cnt':len(cr),
        'db_amt':db['পরিমাণ'].sum(),'db_cnt':len(db),
        'comm_amt':comm['পরিমাণ'].sum()+main['কমিশন'].sum(),
        'comm_cnt':len(comm),
        'start':main['দিন'].min() if not main.empty else '',
        'end':main['দিন'].max() if not main.empty else '',
        'active':len(daily),
        'avg_amt':round(daily['sum'].mean(),2) if not daily.empty else 0,
        'avg_cnt':round(daily['count'].mean(),1) if not daily.empty else 0,
        'night_pct':npct,
        'max_day':daily['sum'].idxmax() if not daily.empty else '',
        'max_day_amt':daily['sum'].max() if not daily.empty else 0,
    }
    return s,main


# ════════════════════════════════════════════════════════════════════════════
#  HTML REPORT GENERATOR
# ════════════════════════════════════════════════════════════════════════════
def build_html(s, main, acc, provider, fmt, fname, focus_num=''):
    ac   = ACCENT.get(provider, '#6b7f99')
    pname = {'nagad':'Nagad','bkash':'bKash','rocket':'Rocket'}.get(provider,'Unknown')
    icon  = {'nagad':'🟠','bkash':'🔴','rocket':'🟣'}.get(provider,'❓')
    now   = datetime.now().strftime('%d %B %Y, %I:%M %p')

    # ── Optional Phone Focus Section ──
    focus_html = ''
    if focus_num:
        f_df = main[main['কাউন্টারপার্ট'].astype(str).str.contains(focus_num, na=False, regex=False)]
        if not f_df.empty:
            f_cr = f_df[f_df['direction']=='credit']
            f_db = f_df[f_df['direction']=='debit']
            f_rows = ''
            for i,(_,r) in enumerate(f_df.iterrows()):
                d_color = '#00875a' if r.get('direction')=='credit' else '#c0392b'
                d_label = 'Credit' if r.get('direction')=='credit' else 'Debit'
                f_rows += f'<tr><td>{i+1}</td><td>{r.get("তারিখ","")}</td><td>{r.get("TXN_ID","")}</td><td>{str(r.get("ধরন",""))[:30]}</td><td style="color:{d_color};font-weight:600">{d_label}</td><td>BDT {r.get("পরিমাণ",0):,.2f}</td><td>{r.get("কাউন্টারপার্ট","")}</td></tr>'
            focus_html = f'''
  <div class="panel" style="margin-bottom:16px;border:2px solid #00875a;background:#f0faf6;">
    <div class="sec-badge" style="background:#00875a;color:#fff;border:none;">🔎 Focus on Number — {focus_num}</div>
    <div class="pt">All transactions with this number ({len(f_df):,} found)</div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px;">
      <div style="background:#fff;border:1px solid #b3dcc8;border-radius:8px;padding:12px;text-align:center;">
        <div style="font-size:10px;color:#6b7f99;text-transform:uppercase;letter-spacing:1px;">Total Txns</div>
        <div style="font-size:20px;font-weight:800;color:#1557a0;">{len(f_df):,}</div>
      </div>
      <div style="background:#fff;border:1px solid #b3dcc8;border-radius:8px;padding:12px;text-align:center;">
        <div style="font-size:10px;color:#6b7f99;text-transform:uppercase;letter-spacing:1px;">Received From</div>
        <div style="font-size:18px;font-weight:800;color:#00875a;font-family:monospace;">BDT {f_cr['পরিমাণ'].sum():,.2f}</div>
        <div style="font-size:11px;color:#6b7f99;">{len(f_cr)} txns</div>
      </div>
      <div style="background:#fff;border:1px solid #b3dcc8;border-radius:8px;padding:12px;text-align:center;">
        <div style="font-size:10px;color:#6b7f99;text-transform:uppercase;letter-spacing:1px;">Sent To</div>
        <div style="font-size:18px;font-weight:800;color:#c0392b;font-family:monospace;">BDT {f_db['পরিমাণ'].sum():,.2f}</div>
        <div style="font-size:11px;color:#6b7f99;">{len(f_db)} txns</div>
      </div>
      <div style="background:#fff;border:1px solid #b3dcc8;border-radius:8px;padding:12px;text-align:center;">
        <div style="font-size:10px;color:#6b7f99;text-transform:uppercase;letter-spacing:1px;">Net Flow</div>
        <div style="font-size:18px;font-weight:800;color:{'#00875a' if (f_cr['পরিমাণ'].sum()-f_db['পরিমাণ'].sum())>=0 else '#c0392b'};font-family:monospace;">BDT {f_cr['পরিমাণ'].sum()-f_db['পরিমাণ'].sum():,.2f}</div>
      </div>
    </div>
    <div style="overflow-x:auto;max-height:500px;">
      <table class="tbl">
        <thead><tr><th>#</th><th>Date</th><th>TXN ID</th><th>Type</th><th>Direction</th><th>Amount</th><th>Counterpart</th></tr></thead>
        <tbody>{f_rows}</tbody>
      </table>
    </div>
  </div>'''
        else:
            focus_html = f'''
  <div class="panel" style="margin-bottom:16px;border:2px solid #c0392b;background:#fde8ec;">
    <div class="sec-badge" style="background:#c0392b;color:#fff;border:none;">🔎 Focus on Number — {focus_num}</div>
    <div style="padding:14px;color:#7a5200;">⚠️ No transactions found involving this number.</div>
  </div>'''

    # ── filter valid counterparts ──
    cp_df = main[~main['কাউন্টারপার্ট'].isin(['SYSTEM','','nan','None','NaN'])]

    # Top 10 overall counterparts
    top_cp = cp_df.groupby('কাউন্টারপার্ট').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
                  .sort_values('amt', ascending=False).head(10).reset_index()

    # Top 10 RECEIVED FROM  (credit = agent received, counterpart = sender)
    cr_df  = cp_df[cp_df['direction'] == 'credit']
    top_recv = cr_df.groupby('কাউন্টারপার্ট').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
                    .sort_values('amt', ascending=False).head(10).reset_index()

    # Top 10 SENT TO  (debit = agent sent, counterpart = receiver)
    db_df  = cp_df[cp_df['direction'] == 'debit']
    top_sent = db_df.groupby('কাউন্টারপার্ট').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
                    .sort_values('amt', ascending=False).head(10).reset_index()

    # Top days
    td = main.groupby('দিন').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
             .sort_values('amt', ascending=False).head(10).reset_index()

    # Size / Type / Monthly
    sz = main.groupby('আকার').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum')).reset_index()
    tp = main.groupby('ধরন').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
             .sort_values('amt', ascending=False).head(8).reset_index()
    mo = main.groupby('মাস').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum')).reset_index()

    # ── Chart.js data helpers ──
    def js_list(series):   return '[' + ','.join(f'"{v}"' for v in series) + ']'
    def js_nums(series):   return '[' + ','.join(f'{v:.2f}' for v in series) + ']'
    def js_ints(series):   return '[' + ','.join(str(int(v)) for v in series) + ']'

    recv_labels = js_list(top_recv['কাউন্টারপার্ট'])
    recv_amts   = js_nums(top_recv['amt'])
    sent_labels = js_list(top_sent['কাউন্টারপার্ট'])
    sent_amts   = js_nums(top_sent['amt'])
    mo_labels   = js_list(mo['মাস'])
    mo_amts     = js_nums(mo['amt'])
    mo_cnts     = js_ints(mo['cnt'])

    try:
        m2 = main.copy()
        m2['hr'] = pd.to_datetime(m2['তারিখ'], errors='coerce').dt.hour.fillna(0).astype(int)
        hr_ser = m2.groupby('hr')['পরিমাণ'].sum().reindex(range(24), fill_value=0)
        hourly_amts   = js_nums(hr_ser)
        hourly_labels = js_list([f"{h}:00" for h in range(24)])
        hourly_colors = '[' + ','.join(f'"{"#c0392b" if h in [22,23,0,1] else ac}"' for h in range(24)) + ']'
    except:
        hourly_amts   = js_nums([0]*24)
        hourly_labels = js_list([f"{h}:00" for h in range(24)])
        hourly_colors = f'["{ac}"]'

    # ── HTML table row helpers ──
    def tbl_rows_cp(df):
        rows = ''
        for i,(_,r) in enumerate(df.iterrows()):
            rows += (f'<tr><td>{i+1}</td><td>{r["কাউন্টারপার্ট"]}</td>'
                     f'<td>{int(r["cnt"])}</td><td>BDT {r["amt"]:,.2f}</td></tr>')
        return rows

    cp_rows   = tbl_rows_cp(top_cp)
    recv_rows = tbl_rows_cp(top_recv)
    sent_rows = tbl_rows_cp(top_sent)

    td_rows = ''.join(
        f'<tr><td>{i+1}</td><td>{r["দিন"]}</td><td>{int(r["cnt"])}</td><td>BDT {r["amt"]:,.2f}</td></tr>'
        for i,(_,r) in enumerate(td.iterrows()))
    sz_rows = ''.join(
        f'<tr><td>{r["আকার"]}</td><td>{int(r["cnt"])}</td><td>BDT {r["amt"]:,.2f}</td></tr>'
        for _,r in sz.iterrows())
    tp_rows = ''.join(
        f'<tr><td>{r["ধরন"][:35]}</td><td>{int(r["cnt"])}</td><td>BDT {r["amt"]:,.2f}</td></tr>'
        for _,r in tp.iterrows())
    mo_rows = ''.join(
        f'<tr><td>{r["মাস"]}</td><td>{int(r["cnt"])}</td><td>BDT {r["amt"]:,.2f}</td></tr>'
        for _,r in mo.iterrows())

    warn_top = (f'<div class="warn">⚠️ <strong>{top_cp.iloc[0]["কাউন্টারপার্ট"]}</strong> — '
                f'{int(top_cp.iloc[0]["cnt"])} transactions. Unusual activity — review recommended.</div>'
                if not top_cp.empty and int(top_cp.iloc[0]['cnt']) > 30 else '')

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>MFS Forensic Report — {acc}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
*{{margin:0;padding:0;box-sizing:border-box;}}
body{{background:#f0f4f8;font-family:'Inter',sans-serif;font-size:14px;color:#1a2332;}}
.header{{background:linear-gradient(135deg,#1a0533 0%,#2d0a5e 40%,#8b1a6b 75%,{ac} 100%);padding:36px 48px;}}
.header h1{{color:#fff;font-size:28px;font-weight:800;margin-bottom:6px;}}
.header p{{color:rgba(255,255,255,.78);font-size:13px;}}
.hm{{display:flex;gap:30px;margin-top:20px;flex-wrap:wrap;}}
.hm-item .lbl{{font-size:10px;color:rgba(255,255,255,.6);text-transform:uppercase;letter-spacing:1px;}}
.hm-item .val{{font-size:15px;font-weight:700;color:#fff;margin-top:2px;font-family:monospace;}}
.body{{max-width:1240px;margin:0 auto;padding:28px 24px 60px;}}
.sg{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:22px;}}
@media(max-width:800px){{.sg{{grid-template-columns:repeat(2,1fr);}}}}
.sc{{background:#fff;border:1px solid #dde3ed;border-radius:10px;padding:17px;position:relative;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.05);}}
.sc::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;}}
.sc.bl::before{{background:#1557a0;}}.sc.gr::before{{background:#00875a;}}
.sc.rd::before{{background:#c0392b;}}.sc.am::before{{background:#b86e00;}}
.sc .ic{{font-size:18px;margin-bottom:7px;}}
.sc .lb{{font-size:10px;color:#6b7f99;text-transform:uppercase;letter-spacing:1.2px;margin-bottom:3px;}}
.sc .vl{{font-size:21px;font-weight:800;font-family:monospace;line-height:1;}}
.sc .vl.bl{{color:#1557a0;}}.sc .vl.gr{{color:#00875a;}}.sc .vl.rd{{color:#c0392b;}}.sc .vl.am{{color:#b86e00;}}
.sc .sb{{font-size:11px;color:#6b7f99;margin-top:4px;}}
.panel{{background:#fff;border:1px solid #dde3ed;border-radius:10px;padding:20px 22px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,.04);}}
.pt{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:#6b7f99;
     margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid #f0f4f8;
     display:flex;align-items:center;gap:7px;}}
.pt::before{{content:'';width:3px;height:13px;background:{ac};border-radius:2px;display:inline-block;}}
.two{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;}}
@media(max-width:760px){{.two{{grid-template-columns:1fr;}}}}
.tbl{{width:100%;border-collapse:collapse;font-size:12px;}}
.tbl thead tr{{background:#f4f6f9;}}
.tbl th{{padding:9px 11px;text-align:left;font-size:10px;text-transform:uppercase;
         letter-spacing:1px;color:#6b7f99;font-weight:700;border-bottom:1px solid #dde3ed;white-space:nowrap;}}
.tbl td{{padding:9px 11px;border-bottom:1px solid #f0f4f8;font-family:monospace;font-size:11px;}}
.tbl tr:hover td{{background:#fafbff;}}
.tbl td:first-child{{color:#9baab8;text-align:center;width:28px;}}
.infobox{{background:#f0faf6;border:1px solid #b3dcc8;border-left:4px solid #00875a;
          border-radius:8px;padding:14px 18px;font-size:13px;line-height:1.9;margin-bottom:20px;}}
.warn{{background:#fffbec;border:1px solid #edd890;border-radius:6px;
       padding:10px 14px;font-size:12px;color:#7a5200;margin-top:8px;}}
.sec-badge{{display:inline-block;background:{ac}18;border:1px solid {ac}44;
            color:{ac};padding:3px 10px;border-radius:5px;font-size:11px;
            font-weight:700;margin-bottom:14px;text-transform:uppercase;letter-spacing:.5px;}}
.chart-wrap{{position:relative;}}
.footer{{text-align:center;padding:20px;font-size:11px;color:#9baab8;
         border-top:1px solid #dde3ed;margin-top:30px;}}
@media print{{
  body{{background:#fff;}}
  .header{{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
  .chart-wrap canvas{{max-width:100%;}}
}}
</style>
</head>
<body>

<!-- ═══════ HEADER ═══════ -->
<div class="header">
  <h1>{icon} Mobile Financial System — Forensic Report</h1>
  <p>Provider: <strong>{pname}</strong> &nbsp;·&nbsp; Format: {fmt} &nbsp;·&nbsp;
     File: {fname} &nbsp;·&nbsp; Generated: {now}</p>
  <div class="hm">
    <div class="hm-item"><div class="lbl">Account</div><div class="val">{acc or '—'}</div></div>
    <div class="hm-item"><div class="lbl">Start Date</div><div class="val">{s['start']}</div></div>
    <div class="hm-item"><div class="lbl">End Date</div><div class="val">{s['end']}</div></div>
    <div class="hm-item"><div class="lbl">Active Days</div><div class="val">{s['active']}</div></div>
    <div class="hm-item"><div class="lbl">Night Activity</div><div class="val">{s['night_pct']}%</div></div>
    <div class="hm-item"><div class="lbl">Daily Avg</div><div class="val">BDT {s['avg_amt']:,.0f}</div></div>
  </div>
</div>

<div class="body">

  <!-- ═══════ SUMMARY BOX ═══════ -->
  <div class="infobox">
    📋 <strong>{pname}</strong> account recorded <strong>{s['total']:,} transactions</strong>
    over <strong>{s['active']} active days</strong>.
    Total Credit (In): <strong style="color:#00875a">BDT {s['cr_amt']:,.2f}</strong> ({s['cr_cnt']:,} txns) &nbsp;|&nbsp;
    Total Debit (Out): <strong style="color:#c0392b">BDT {s['db_amt']:,.2f}</strong> ({s['db_cnt']:,} txns) &nbsp;|&nbsp;
    Commission Earned: <strong style="color:#b86e00">BDT {s['comm_amt']:,.2f}</strong>.
    Night-time activity (10PM–2AM): <strong style="color:#b86e00">{s['night_pct']}%</strong>.
    {'<br>⚠️ <strong>High night-time activity detected — suspicious, requires investigation.</strong>' if s['night_pct']>35 else ''}
  </div>

  <!-- ═══════ METRIC CARDS ═══════ -->
  <div class="sg">
    <div class="sc bl">
      <div class="ic">🔢</div><div class="lb">Total Transactions</div>
      <div class="vl bl">{s['total']:,}</div><div class="sb">Excluding commission</div>
    </div>
    <div class="sc gr">
      <div class="ic">⬇️</div><div class="lb">Total Credit (In)</div>
      <div class="vl gr">BDT {s['cr_amt']:,.2f}</div><div class="sb">{s['cr_cnt']:,} transactions</div>
    </div>
    <div class="sc rd">
      <div class="ic">⬆️</div><div class="lb">Total Debit (Out)</div>
      <div class="vl rd">BDT {s['db_amt']:,.2f}</div><div class="sb">{s['db_cnt']:,} transactions</div>
    </div>
    <div class="sc am">
      <div class="ic">💰</div><div class="lb">Commission Earned</div>
      <div class="vl am">BDT {s['comm_amt']:,.2f}</div><div class="sb">{s['comm_cnt']:,} entries</div>
    </div>
  </div>

  <!-- ═══════ TREND CHARTS ═══════ -->
  <div class="two">
    <div class="panel">
      <div class="pt">Monthly Transaction Volume</div>
      <div class="chart-wrap" style="height:240px;">
        <canvas id="cMonthly"></canvas>
      </div>
    </div>
    <div class="panel">
      <div class="pt">Hourly Activity Distribution (Red = Night 10PM–2AM)</div>
      <div class="chart-wrap" style="height:240px;">
        <canvas id="cHourly"></canvas>
      </div>
    </div>
  </div>

  <!-- ═══════ TOP 10 MONEY RECEIVED FROM ═══════ -->
  <div class="panel">
    <div class="sec-badge">💚 Top 10 — Money Received From</div>
    <div class="pt">Top 10 Numbers — Highest Amount Received From (Credit)</div>
    <div class="two" style="margin-bottom:0;">
      <div>
        <div class="chart-wrap" style="height:320px;">
          <canvas id="cRecv"></canvas>
        </div>
      </div>
      <div style="overflow-x:auto;">
        <table class="tbl">
          <thead>
            <tr><th>#</th><th>Account Number</th><th>Transactions</th><th>Total Received</th></tr>
          </thead>
          <tbody>{recv_rows}</tbody>
        </table>
        {'<div class="warn">⚠️ ' + str(top_recv.iloc[0]["কাউন্টারপার্ট"]) + ' — ' + str(int(top_recv.iloc[0]["cnt"])) + ' transactions. Unusually high frequency.</div>' if not top_recv.empty and int(top_recv.iloc[0]['cnt'])>30 else ''}
      </div>
    </div>
  </div>

  <!-- ═══════ TOP 10 MONEY SENT TO ═══════ -->
  <div class="panel">
    <div class="sec-badge" style="background:#fde8ec;border-color:#f5b8c4;color:#c0392b;">🔴 Top 10 — Money Sent To</div>
    <div class="pt">Top 10 Numbers — Highest Amount Sent To (Debit)</div>
    <div class="two" style="margin-bottom:0;">
      <div>
        <div class="chart-wrap" style="height:320px;">
          <canvas id="cSent"></canvas>
        </div>
      </div>
      <div style="overflow-x:auto;">
        <table class="tbl">
          <thead>
            <tr><th>#</th><th>Account Number</th><th>Transactions</th><th>Total Sent</th></tr>
          </thead>
          <tbody>{sent_rows}</tbody>
        </table>
        {'<div class="warn">⚠️ ' + str(top_sent.iloc[0]["কাউন্টারপার্ট"]) + ' — ' + str(int(top_sent.iloc[0]["cnt"])) + ' transactions. Unusually high frequency.</div>' if not top_sent.empty and int(top_sent.iloc[0]['cnt'])>30 else ''}
      </div>
    </div>
  </div>

  <!-- BREAKDOWN TABLES -->
  <div class="two">
    <div class="panel">
      <div class="pt">Monthly Breakdown</div>
      <table class="tbl">
        <thead><tr><th>Month</th><th>Transactions</th><th>Total Amount</th></tr></thead>
        <tbody>{mo_rows}</tbody>
      </table>
    </div>
    <div class="panel">
      <div class="pt">Transaction Size Breakdown</div>
      <table class="tbl">
        <thead><tr><th>Category</th><th>Count</th><th>Total Amount</th></tr></thead>
        <tbody>{sz_rows}</tbody>
      </table>
      <br>
      <div class="pt">Transaction Type Breakdown</div>
      <table class="tbl">
        <thead><tr><th>Type</th><th>Count</th><th>Total Amount</th></tr></thead>
        <tbody>{tp_rows}</tbody>
      </table>
    </div>
  </div>

  <!-- TOP COUNTERPARTS + TOP DAYS -->
  <div class="two">
    <div class="panel">
      <div class="pt">Top 10 Counterpart Accounts (All)</div>
      <table class="tbl">
        <thead><tr><th>#</th><th>Account</th><th>Transactions</th><th>Total Amount</th></tr></thead>
        <tbody>{cp_rows}</tbody>
      </table>
      {warn_top}
    </div>
    <div class="panel">
      <div class="pt">Top 10 Busiest Days</div>
      <table class="tbl">
        <thead><tr><th>#</th><th>Date</th><th>Transactions</th><th>Total Amount</th></tr></thead>
        <tbody>{td_rows}</tbody>
      </table>
    </div>
  </div>

  {focus_html}

</div>

<div class="footer">
  MFS Forensic Report &nbsp;·&nbsp; {pname} &nbsp;·&nbsp; Account: {acc} &nbsp;·&nbsp; Generated: {now}
</div>

<!-- CHART.JS SCRIPTS -->
<script>
Chart.defaults.font.family = "'Inter', sans-serif";
Chart.defaults.font.size   = 11;
Chart.defaults.color       = '#6b7f99';

// ── Monthly Chart ──
new Chart(document.getElementById('cMonthly'), {{
  type: 'bar',
  data: {{
    labels: {mo_labels},
    datasets: [
      {{
        label: 'Amount (BDT)',
        data: {mo_amts},
        backgroundColor: '{ac}bb',
        borderColor: '{ac}',
        borderWidth: 1,
        borderRadius: 4,
        yAxisID: 'y',
      }},
      {{
        label: 'Transactions',
        data: {mo_cnts},
        type: 'line',
        borderColor: '#c0392b',
        backgroundColor: '#c0392b22',
        borderWidth: 2,
        pointRadius: 4,
        pointBackgroundColor: '#c0392b',
        tension: 0.4,
        yAxisID: 'y2',
      }}
    ]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{ legend: {{ position: 'top' }} }},
    scales: {{
      y:  {{ position: 'left',  ticks: {{ callback: v => 'BDT ' + (v/1000).toFixed(0) + 'K' }},
             grid: {{ color: '#f0f4f8' }} }},
      y2: {{ position: 'right', ticks: {{ callback: v => v + ' txn' }},
             grid: {{ display: false }} }},
      x:  {{ grid: {{ display: false }} }}
    }}
  }}
}});

// ── Hourly Chart ──
new Chart(document.getElementById('cHourly'), {{
  type: 'bar',
  data: {{
    labels: {hourly_labels},
    datasets: [{{
      label: 'Amount (BDT)',
      data: {hourly_amts},
      backgroundColor: {hourly_colors},
      borderWidth: 0,
      borderRadius: 3,
    }}]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ display: false }},
      tooltip: {{ callbacks: {{ label: ctx => 'BDT ' + ctx.raw.toLocaleString() }} }}
    }},
    scales: {{
      y: {{ ticks: {{ callback: v => 'BDT ' + (v/1000).toFixed(0) + 'K' }},
            grid: {{ color: '#f0f4f8' }} }},
      x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 9 }} }} }}
    }}
  }}
}});

// ── Top 10 Money Received From ──
new Chart(document.getElementById('cRecv'), {{
  type: 'bar',
  data: {{
    labels: {recv_labels},
    datasets: [{{
      label: 'Total Received (BDT)',
      data: {recv_amts},
      backgroundColor: '#00875a99',
      borderColor: '#00875a',
      borderWidth: 1,
      borderRadius: 4,
    }}]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ display: false }},
      tooltip: {{
        callbacks: {{
          label: ctx => ' BDT ' + parseFloat(ctx.raw).toLocaleString('en-BD', {{minimumFractionDigits:2}})
        }}
      }}
    }},
    scales: {{
      x: {{ ticks: {{ callback: v => 'BDT ' + (v/1000).toFixed(0) + 'K' }},
            grid: {{ color: '#f0f4f8' }} }},
      y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10, family: 'monospace' }} }} }}
    }}
  }}
}});

// ── Top 10 Money Sent To ──
new Chart(document.getElementById('cSent'), {{
  type: 'bar',
  data: {{
    labels: {sent_labels},
    datasets: [{{
      label: 'Total Sent (BDT)',
      data: {sent_amts},
      backgroundColor: '#c0392b99',
      borderColor: '#c0392b',
      borderWidth: 1,
      borderRadius: 4,
    }}]
  }},
  options: {{
    indexAxis: 'y',
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{ display: false }},
      tooltip: {{
        callbacks: {{
          label: ctx => ' BDT ' + parseFloat(ctx.raw).toLocaleString('en-BD', {{minimumFractionDigits:2}})
        }}
      }}
    }},
    scales: {{
      x: {{ ticks: {{ callback: v => 'BDT ' + (v/1000).toFixed(0) + 'K' }},
            grid: {{ color: '#f0f4f8' }} }},
      y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10, family: 'monospace' }} }} }}
    }}
  }}
}});
</script>
</body>
</html>"""


# ════════════════════════════════════════════════════════════════════════════
#  WORD REPORT GENERATOR
# ════════════════════════════════════════════════════════════════════════════
def build_docx(s, main, acc, provider, fmt, fname, focus_num=''):
    if not DOCX_OK:
        return None

    from docx.oxml.ns import qn
    from docx.oxml   import OxmlElement

    doc = Document()

    # ── Accent color ──────────────────────────────────────────────────────
    ac_hex = {
        'nagad':  'E85D04',
        'bkash':  'E2136E',
        'rocket': '7C3AED',
    }.get(provider, '6B7F99')

    def hex_to_rgb(h):
        return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

    ac_rgb = hex_to_rgb(ac_hex)

    green  = RGBColor(0x00, 0x87, 0x5a)
    red    = RGBColor(0xc0, 0x39, 0x2b)
    amber  = RGBColor(0xb8, 0x6e, 0x00)
    blue   = RGBColor(0x15, 0x57, 0xa0)
    muted  = RGBColor(0x6b, 0x7f, 0x99)
    dark   = RGBColor(0x1a, 0x23, 0x32)
    white  = RGBColor(0xff, 0xff, 0xff)

    pname = {'nagad':'Nagad','bkash':'bKash','rocket':'Rocket (DBBL)'}.get(provider,'Unknown')
    now   = datetime.now().strftime('%d %B %Y, %I:%M %p')

    # ── Page setup ───────────────────────────────────────────────────────
    for sec in doc.sections:
        sec.top_margin    = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin   = Cm(2.0)
        sec.right_margin  = Cm(2.0)

    # ── Helper: shade a table cell ────────────────────────────────────────
    def shade_cell(cell, hex_color):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement('w:shd')
        shd.set(qn('w:val'),   'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'),  hex_color)
        tcPr.append(shd)

    # ── Helper: set cell text with formatting ─────────────────────────────
    def cell_text(cell, text, bold=False, size=9, color=None, align='LEFT'):
        cell.text = ''
        p   = cell.paragraphs[0]
        p.alignment = {'LEFT': WD_ALIGN_PARAGRAPH.LEFT,
                       'CENTER': WD_ALIGN_PARAGRAPH.CENTER,
                       'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
        run = p.add_run(str(text))
        run.font.size = Pt(size)
        run.font.bold = bold
        if color: run.font.color.rgb = color

    # ── Helper: add styled table ──────────────────────────────────────────
    def add_table(headers, rows, col_w=None, hdr_color='1A2332'):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = 'Table Grid'
        # Header row
        hrow = t.rows[0]
        for i, h in enumerate(headers):
            shade_cell(hrow.cells[i], hdr_color)
            cell_text(hrow.cells[i], h, bold=True, size=9, color=white, align='CENTER')
        # Data rows
        for ri, row_data in enumerate(rows):
            row = t.add_row()
            bg  = 'F8F9FB' if ri % 2 == 0 else 'FFFFFF'
            for i, val in enumerate(row_data):
                shade_cell(row.cells[i], bg)
                cell_text(row.cells[i], val, size=9, color=dark)
        # Column widths
        if col_w:
            for row in t.rows:
                for i, cell in enumerate(row.cells):
                    if i < len(col_w):
                        cell.width = Cm(col_w[i])
        return t

    # ── Helper: section heading ───────────────────────────────────────────
    def section_heading(text):
        p   = doc.add_paragraph()
        run = p.add_run(text)
        run.font.size  = Pt(12)
        run.font.bold  = True
        run.font.color.rgb = ac_rgb
        # bottom border
        pPr   = p._p.get_or_add_pPr()
        pBdr  = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'),   'single')
        bottom.set(qn('w:sz'),    '4')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), ac_hex)
        pBdr.append(bottom)
        pPr.append(pBdr)
        return p

    def add_spacer():
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ── Helper: key-value info line ───────────────────────────────────────
    def info_line(label, value, val_color=None):
        p    = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        run1 = p.add_run(f"{label}: ")
        run1.font.size  = Pt(10)
        run1.font.color.rgb = muted
        run2 = p.add_run(str(value))
        run2.font.size  = Pt(10)
        run2.font.bold  = True
        run2.font.color.rgb = val_color or dark

    # ══════════════════════════════════════════════════════════════════════
    #  COVER HEADER (shaded box)
    # ══════════════════════════════════════════════════════════════════════
    # Title
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(6)
    run = p.add_run('Mobile Financial System — Forensic Report')
    run.font.size  = Pt(20)
    run.font.bold  = True
    run.font.color.rgb = ac_rgb

    # Subtitle line
    p2  = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(10)
    sub = p2.add_run(f'Provider: {pname}   |   Account: {acc or "—"}   |   Generated: {now}')
    sub.font.size  = Pt(10)
    sub.font.color.rgb = muted

    # Info grid using a borderless table
    info_t = doc.add_table(rows=2, cols=4)
    info_t.style = 'Table Grid'
    info_data = [
        ('Start Date',    s['start'],       dark),
        ('End Date',      s['end'],         dark),
        ('Active Days',   str(s['active']), dark),
        ('Night Activity',f"{s['night_pct']}%",
         red if s['night_pct'] > 35 else amber),
    ]
    for i, (lbl, val, col) in enumerate(info_data):
        shade_cell(info_t.rows[0].cells[i], 'F0F4F8')
        shade_cell(info_t.rows[1].cells[i], 'FFFFFF')
        cell_text(info_t.rows[0].cells[i], lbl,  bold=True, size=8, color=muted, align='CENTER')
        cell_text(info_t.rows[1].cells[i], val,  bold=True, size=11, color=col,  align='CENTER')

    add_spacer()

    # ══════════════════════════════════════════════════════════════════════
    #  EXECUTIVE SUMMARY
    # ══════════════════════════════════════════════════════════════════════
    section_heading('Executive Summary')
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    p.add_run(
        f'{pname} account {acc} recorded '
    ).font.size = Pt(11)
    run_bold = p.add_run(f'{s["total"]:,} transactions')
    run_bold.font.bold = True
    run_bold.font.size = Pt(11)
    p.add_run(
        f' over {s["active"]} active days ({s["start"]} to {s["end"]}). '
        f'Daily average: BDT {s["avg_amt"]:,.2f}. '
        f'Night-time activity (10PM–2AM): {s["night_pct"]}%.'
        + (' WARNING: High night-time activity — requires investigation.' if s['night_pct'] > 35 else '')
    ).font.size = Pt(11)
    add_spacer()

    # ══════════════════════════════════════════════════════════════════════
    #  KEY METRICS TABLE
    # ══════════════════════════════════════════════════════════════════════
    section_heading('Key Financial Metrics')
    add_table(
        ['Metric', 'Amount (BDT)', 'Count'],
        [
            ['Total Transactions',      '—',                         f"{s['total']:,}"],
            ['Total Credit (In)',       f"BDT {s['cr_amt']:,.2f}",  f"{s['cr_cnt']:,} txns"],
            ['Total Debit (Out)',       f"BDT {s['db_amt']:,.2f}",  f"{s['db_cnt']:,} txns"],
            ['Commission Earned',       f"BDT {s['comm_amt']:,.2f}",f"{s['comm_cnt']:,} entries"],
            ['Daily Average Amount',    f"BDT {s['avg_amt']:,.2f}", f"≈{s['avg_cnt']} txns/day"],
            ['Night Activity (10PM-2AM)',f"{s['night_pct']}%",      'Review if > 35%'],
        ],
        col_w=[5.5, 4.0, 3.5]
    )
    add_spacer()

    # ══════════════════════════════════════════════════════════════════════
    #  MONTHLY BREAKDOWN
    # ══════════════════════════════════════════════════════════════════════
    section_heading('Monthly Transaction Breakdown')
    mo = main.groupby('মাস').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum')).reset_index()
    add_table(
        ['Month', 'Transactions', 'Total Amount (BDT)'],
        [[r['মাস'], f"{int(r['cnt']):,}", f"BDT {r['amt']:,.2f}"] for _,r in mo.iterrows()],
        col_w=[4.5, 3.5, 5.0]
    )
    add_spacer()

    # ══════════════════════════════════════════════════════════════════════
    #  SIZE BREAKDOWN
    # ══════════════════════════════════════════════════════════════════════
    section_heading('Transaction Size Breakdown')
    sz = main.groupby('আকার').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum')).reset_index()
    add_table(
        ['Size Category', 'Count', 'Total Amount (BDT)'],
        [[r['আকার'], f"{int(r['cnt']):,}", f"BDT {r['amt']:,.2f}"] for _,r in sz.iterrows()],
        col_w=[5.5, 3.0, 4.5]
    )
    add_spacer()

    # ══════════════════════════════════════════════════════════════════════
    #  TRANSACTION TYPE BREAKDOWN
    # ══════════════════════════════════════════════════════════════════════
    section_heading('Transaction Type Breakdown')
    tp = main.groupby('ধরন').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
             .sort_values('amt', ascending=False).head(8).reset_index()
    add_table(
        ['Transaction Type', 'Count', 'Total Amount (BDT)'],
        [[r['ধরন'][:40], f"{int(r['cnt']):,}", f"BDT {r['amt']:,.2f}"] for _,r in tp.iterrows()],
        col_w=[7.0, 2.5, 3.5]
    )
    add_spacer()

    # ══════════════════════════════════════════════════════════════════════
    #  TOP 10 RECEIVED FROM (Credit)
    # ══════════════════════════════════════════════════════════════════════
    section_heading('Top 10 Numbers — Money Received From (Credit)')
    cp_df = main[~main['কাউন্টারপার্ট'].isin(['SYSTEM','','nan','None','NaN'])]
    top_recv = cp_df[cp_df['direction']=='credit']\
                 .groupby('কাউন্টারপার্ট').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
                 .sort_values('amt', ascending=False).head(10).reset_index()
    add_table(
        ['#', 'Account Number', 'Transactions', 'Total Received (BDT)'],
        [[i+1, r['কাউন্টারপার্ট'], f"{int(r['cnt']):,}", f"BDT {r['amt']:,.2f}"]
         for i,(_,r) in enumerate(top_recv.iterrows())],
        col_w=[1.0, 5.0, 3.0, 4.0]
    )
    add_spacer()

    # ══════════════════════════════════════════════════════════════════════
    #  TOP 10 SENT TO (Debit)
    # ══════════════════════════════════════════════════════════════════════
    section_heading('Top 10 Numbers — Money Sent To (Debit)')
    top_sent = cp_df[cp_df['direction']=='debit']\
                 .groupby('কাউন্টারপার্ট').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
                 .sort_values('amt', ascending=False).head(10).reset_index()
    add_table(
        ['#', 'Account Number', 'Transactions', 'Total Sent (BDT)'],
        [[i+1, r['কাউন্টারপার্ট'], f"{int(r['cnt']):,}", f"BDT {r['amt']:,.2f}"]
         for i,(_,r) in enumerate(top_sent.iterrows())],
        col_w=[1.0, 5.0, 3.0, 4.0]
    )
    add_spacer()

    # ══════════════════════════════════════════════════════════════════════
    #  TOP 10 OVERALL COUNTERPARTS
    # ══════════════════════════════════════════════════════════════════════
    section_heading('Top 10 Counterpart Accounts (All Directions)')
    top_all = cp_df.groupby('কাউন্টারপার্ট').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
                .sort_values('amt', ascending=False).head(10).reset_index()
    add_table(
        ['#', 'Account Number', 'Transactions', 'Total Amount (BDT)'],
        [[i+1, r['কাউন্টারপার্ট'], f"{int(r['cnt']):,}", f"BDT {r['amt']:,.2f}"]
         for i,(_,r) in enumerate(top_all.iterrows())],
        col_w=[1.0, 5.0, 3.0, 4.0]
    )
    add_spacer()

    # ══════════════════════════════════════════════════════════════════════
    #  TOP 10 BUSIEST DAYS
    # ══════════════════════════════════════════════════════════════════════
    section_heading('Top 10 Busiest Days')
    td = main.groupby('দিন').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
             .sort_values('amt', ascending=False).head(10).reset_index()
    add_table(
        ['#', 'Date', 'Transactions', 'Total Amount (BDT)'],
        [[i+1, r['দিন'], f"{int(r['cnt']):,}", f"BDT {r['amt']:,.2f}"]
         for i,(_,r) in enumerate(td.iterrows())],
        col_w=[1.0, 4.0, 3.0, 5.0]
    )

    # ══════════════════════════════════════════════════════════════════════
    #  PHONE NUMBER FOCUS (only if provided)
    # ══════════════════════════════════════════════════════════════════════
    if focus_num:
        add_spacer()
        f_df = main[main['কাউন্টারপার্ট'].astype(str).str.contains(focus_num, na=False, regex=False)]
        section_heading(f'Number Focus Analysis — {focus_num}')
        if f_df.empty:
            p = doc.add_paragraph()
            p.add_run(f'No transactions found involving {focus_num}.').font.color.rgb = red
        else:
            fcr = f_df[f_df['direction']=='credit']
            fdb = f_df[f_df['direction']=='debit']
            net = fcr['পরিমাণ'].sum() - fdb['পরিমাণ'].sum()

            # Summary metrics
            add_table(
                ['Metric', 'Value'],
                [
                    ['Target Number',         focus_num],
                    ['Total Transactions',     f"{len(f_df):,}"],
                    ['Received From (Credit)', f"BDT {fcr['পরিমাণ'].sum():,.2f}  ({len(fcr)} txns)"],
                    ['Sent To (Debit)',        f"BDT {fdb['পরিমাণ'].sum():,.2f}  ({len(fdb)} txns)"],
                    ['Net Flow',               f"BDT {net:,.2f}"],
                    ['First Transaction',      f_df['তারিখ'].min()],
                    ['Last Transaction',       f_df['তারিখ'].max()],
                ],
                col_w=[5.5, 7.5]
            )
            add_spacer()

            # All transactions with this number
            doc.add_paragraph().add_run(
                f'All {len(f_df):,} transactions involving {focus_num}:'
            ).font.bold = True

            cols_show = [c for c in ['তারিখ','TXN_ID','ধরন','direction','পরিমাণ','চ্যানেল'] if c in f_df.columns]
            add_table(
                ['#', 'Date', 'TXN ID', 'Type', 'Direction', 'Amount (BDT)', 'Channel'],
                [[i+1,
                  r.get('তারিখ',''),
                  r.get('TXN_ID',''),
                  str(r.get('ধরন',''))[:28],
                  'Credit' if r.get('direction')=='credit' else 'Debit',
                  f"BDT {r.get('পরিমাণ',0):,.2f}",
                  r.get('চ্যানেল','')]
                 for i,(_,r) in enumerate(f_df.iterrows())],
                col_w=[0.8, 3.2, 3.0, 3.5, 1.8, 2.8, 2.0]
            )

    # ── Footer ──────────────────────────────────────────────────────────
    add_spacer()
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    run = p.add_run(f'Mobile Financial System Analyzer  |  {pname}  |  Account: {acc}  |  {now}')
    run.font.size  = Pt(9)
    run.font.color.rgb = muted
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
#  CHARTS
# ════════════════════════════════════════════════════════════════════════════
def ch_layout(title='', h=300):
    return dict(height=h, paper_bgcolor='white', plot_bgcolor='white',
                margin=dict(l=8,r=8,t=38,b=8),
                font=dict(family='Inter,sans-serif',size=11,color='#6b7f99'),
                title=dict(text=title,font=dict(size=13,color='#1a2332'),x=0),
                xaxis=dict(gridcolor='#f0f4f8',linecolor='#dde3ed'),
                yaxis=dict(gridcolor='#f0f4f8',linecolor='#dde3ed'),
                legend=dict(orientation='h',y=-0.22))

def fig_monthly(main, acc):
    df = main.groupby('মাস').agg(Count=('পরিমাণ','count'), Amount=('পরিমাণ','sum')).reset_index()
    df.columns = ['Month', 'Count', 'Amount']
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=df['Month'], y=df['Amount'], name='Amount (BDT)',
                         marker_color=acc, marker_line_width=0, opacity=0.85), secondary_y=False)
    fig.add_trace(go.Scatter(x=df['Month'], y=df['Count'], name='Transactions',
                             mode='lines+markers', line=dict(color='#c0392b', width=2),
                             marker=dict(size=5, color='#c0392b')), secondary_y=True)
    fig.update_layout(**ch_layout('Monthly Transaction Volume'))
    fig.update_yaxes(title_text="Amount (BDT)", secondary_y=False)
    fig.update_yaxes(title_text="Transactions", secondary_y=True)
    return fig

def fig_daily(main, acc):
    df = main[main['দিন'] != ''].groupby('দিন')['পরিমাণ'].sum().reset_index()
    if df.empty:
        return None
    df.columns = ['Date', 'Amount']
    fig = px.bar(df, x='Date', y='Amount', color_discrete_sequence=[acc],
                 labels={'Amount': 'Amount (BDT)', 'Date': 'Date'})
    fig.update_traces(marker_line_width=0, marker_color=acc, opacity=0.85)
    fig.update_layout(**ch_layout('Daily Transaction Volume'))
    return fig

def fig_hourly(main, acc):
    try:
        m2 = main.copy()
        m2['hour'] = pd.to_datetime(m2['তারিখ'], errors='coerce').dt.hour.fillna(0).astype(int)
        df = m2.groupby('hour')['পরিমাণ'].sum().reindex(range(24), fill_value=0).reset_index()
        df.columns = ['Hour', 'Amount']
        df['color'] = df['Hour'].apply(lambda h: '#c0392b' if h in [22, 23, 0, 1] else acc)
        fig = go.Figure(go.Bar(x=df['Hour'], y=df['Amount'],
                               marker_color=df['color'], marker_line_width=0))
        fig.update_layout(**ch_layout('Hourly Activity (Red = Night 10PM–2AM)'))
        fig.update_xaxes(title_text='Hour of Day')
        fig.update_yaxes(title_text='Amount (BDT)')
        return fig
    except: return None

def fig_size_pie(main):
    df = main.groupby('আকার')['পরিমাণ'].sum().reset_index()
    df.columns = ['Size', 'Amount']
    size_map  = {'ছোট (২-৮হা)': 'Small (2K-8K)', 'মধ্যম (৮-২০হা)': 'Medium (8K-20K)',
                 'উচ্চ (২০হা+)': 'High (20K+)', '<২হা': '<2K',
                 'Small (2K-8K)': 'Small (2K-8K)', 'Medium (8K-20K)': 'Medium (8K-20K)',
                 'High (20K+)': 'High (20K+)', '<2K': '<2K', 'Small (2K-8K)': 'Small (2K-8K)'}
    df['Size'] = df['Size'].map(lambda x: size_map.get(x, x))
    colors = {'Small (2K-8K)': '#1557a0', 'Medium (8K-20K)': '#b86e00',
              'High (20K+)': '#c0392b', '<2K': '#b0b8c8'}
    fig = px.pie(df, values='Amount', names='Size',
                 color='Size', color_discrete_map=colors, hole=0.55)
    fig.update_layout(**ch_layout('Transaction Size Breakdown', h=280))
    fig.update_traces(textposition='outside', textinfo='percent+label')
    return fig

def fig_direction(main):
    df = main.groupby('direction')['পরিমাণ'].sum().reset_index()
    df.columns = ['Direction', 'Amount']
    df['Direction'] = df['Direction'].map({'credit': 'Credit (In)', 'debit': 'Debit (Out)'})
    fig = px.pie(df, values='Amount', names='Direction', hole=0.55,
                 color='Direction',
                 color_discrete_map={'Credit (In)': '#00875a', 'Debit (Out)': '#c0392b'})
    fig.update_layout(**ch_layout('Credit vs Debit', h=280))
    return fig

def fig_type_bar(main, acc):
    df = main.groupby('ধরন').agg(Count=('পরিমাণ','count'), Amount=('পরিমাণ','sum'))\
             .sort_values('Amount', ascending=False).head(6).reset_index()
    df.columns = ['Type', 'Count', 'Amount']
    fig = px.bar(df, x='Amount', y='Type', orientation='h',
                 color_discrete_sequence=[acc], text='Count',
                 labels={'Amount': 'Amount (BDT)', 'Type': 'Transaction Type'})
    fig.update_traces(marker_line_width=0, texttemplate='%{text}', textposition='inside')
    fig.update_layout(**ch_layout('Transaction Type Breakdown (Top 6)', h=260))
    return fig

try:
    from plotly.subplots import make_subplots
except: pass


# ════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ════════════════════════════════════════════════════════════════════════════
def main_app():
    # ── NAVBAR ──────────────────────────────────────────────────────────
    st.markdown(
        '<div style="background:#fff;border-bottom:1px solid #e8ecf4;padding:12px 28px;'
        'display:flex;align-items:center;justify-content:space-between;'
        'box-shadow:0 1px 4px rgba(0,0,0,.05);">'
        '<div style="display:flex;align-items:center;gap:10px;">'
        '<div style="width:36px;height:36px;background:linear-gradient(135deg,#e2136e,#7c3aed);'
        'border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:17px;">💳</div>'
        '<div><div style="font-size:15px;font-weight:800;color:#1a2332;">Mobile Financial System Analyzer</div>'
        '<div style="font-size:11px;color:#9baab8;">Smart Analysis • Secure • Accurate • Fast</div></div></div>'
        '<div style="display:flex;gap:24px;font-size:13px;font-weight:600;color:#6b7f99;">'
        '<span>ⓘ How It Works</span><span>✓ Security</span><span>? Help</span></div></div>',
        unsafe_allow_html=True
    )

    # ── HERO (split into small chunks) ───────────────────────────────────
    # Part 1: Hero wrapper + title + description
    st.markdown(
        '<div style="background:linear-gradient(135deg,#1a0533 0%,#2d0a5e 35%,#7c3aed 70%,#e2136e 100%);'
        'padding:48px 44px 36px;">'
        '<h1 style="color:#fff;font-size:36px;font-weight:800;margin:0 0 12px;line-height:1.2;">'
        'Mobile Financial System Analyzer</h1>'
        '<p style="color:rgba(255,255,255,.82);font-size:14px;margin:0 0 22px;max-width:560px;line-height:1.7;">'
        'Upload mobile financial statements (Excel) and get comprehensive insights, '
        'analytics, and professional forensic reports instantly.</p></div>',
        unsafe_allow_html=True
    )

    # Part 2: Supported Services badges
    st.markdown(
        '<div style="background:linear-gradient(135deg,#2d0a5e,#7c3aed);padding:0 44px 12px;">'
        '<div style="font-size:12px;color:rgba(255,255,255,.6);text-transform:uppercase;'
        'letter-spacing:1.5px;margin-bottom:10px;">Supported Services:</div>'
        '<div style="display:flex;gap:10px;flex-wrap:wrap;">'
        '<div style="background:#fff;border-radius:8px;padding:7px 18px;font-size:13px;'
        'font-weight:700;color:#e2136e;">🔴 bKash</div>'
        '<div style="background:#fff;border-radius:8px;padding:7px 18px;font-size:13px;'
        'font-weight:700;color:#e85d04;">🟠 Nagad</div>'
        '<div style="background:#fff;border-radius:8px;padding:7px 18px;font-size:13px;'
        'font-weight:700;color:#7c3aed;">🟣 Rocket</div>'
        '</div></div>',
        unsafe_allow_html=True
    )

    # Part 3: Feature pills
    st.markdown(
        '<div style="background:linear-gradient(135deg,#7c3aed,#e2136e);padding:12px 44px 32px;">'
        '<div style="display:flex;gap:8px;flex-wrap:wrap;">'
        '<span style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);'
        'color:rgba(255,255,255,.9);padding:6px 13px;border-radius:7px;font-size:12px;font-weight:600;">'
        '📊 Financial Analysis</span>'
        '<span style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);'
        'color:rgba(255,255,255,.9);padding:6px 13px;border-radius:7px;font-size:12px;font-weight:600;">'
        '🔄 Transaction Analytics</span>'
        '<span style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);'
        'color:rgba(255,255,255,.9);padding:6px 13px;border-radius:7px;font-size:12px;font-weight:600;">'
        '👥 Top Sender/Receiver Analysis</span>'
        '<span style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);'
        'color:rgba(255,255,255,.9);padding:6px 13px;border-radius:7px;font-size:12px;font-weight:600;">'
        '📈 Trend & Pattern Detection</span>'
        '<span style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);'
        'color:rgba(255,255,255,.9);padding:6px 13px;border-radius:7px;font-size:12px;font-weight:600;">'
        '📄 HTML Forensic Reports</span>'
        '</div></div>',
        unsafe_allow_html=True
    )

    # ── UPLOAD SECTION ────────────────────────────────────────────────────
    st.markdown('<div style="height:20px;"></div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([5, 4, 4], gap="large")

    with col1:
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">'
            '<div style="width:30px;height:30px;background:linear-gradient(135deg,#e2136e,#b8005a);'
            'border-radius:7px;display:flex;align-items:center;justify-content:center;">📤</div>'
            '<div><div style="font-size:14px;font-weight:700;color:#1a2332;">'
            '1. Upload Statement (Excel / PDF)</div>'
            '<div style="font-size:11px;color:#9baab8;">'
            'Upload bKash / Nagad (Excel) or Rocket (PDF)</div></div></div>',
            unsafe_allow_html=True
        )
        uploaded = st.file_uploader(
            "Drag & drop or click to browse",
            type=["xlsx", "xls", "csv", "pdf"],
            help="Max 50MB · XLSX/XLS/CSV for bKash/Nagad · PDF for Rocket"
        )
        st.markdown(
            '<div style="font-size:11px;color:#9baab8;margin-top:4px;">'
            '✔ Supported: .XLSX · .XLS · .CSV · .PDF &nbsp;|&nbsp; ⓘ Max: 50MB</div>',
            unsafe_allow_html=True
        )
        if uploaded:
            st.success(f"✅ **{uploaded.name}** ({uploaded.size/1024:.1f} KB) — Ready")

    with col2:
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">'
            '<div style="width:30px;height:30px;background:linear-gradient(135deg,#7c3aed,#5b21b6);'
            'border-radius:7px;display:flex;align-items:center;justify-content:center;">📱</div>'
            '<div><div style="font-size:14px;font-weight:700;color:#1a2332;">'
            '2. Account / Number (Optional)</div>'
            '<div style="font-size:11px;color:#9baab8;">'
            'Focus analysis on a specific number</div></div></div>',
            unsafe_allow_html=True
        )
        focus_num = st.text_input(
            "Phone / Account Number",
            value="",
            placeholder="e.g. 01806125243",
            label_visibility="collapsed"
        ).strip()
        st.markdown(
            '<div style="background:#f8f9fb;border:1px solid #e8ecf4;border-radius:8px;'
            'padding:10px 12px;font-size:12px;color:#6b7f99;margin-top:4px;">'
            'ⓘ This helps highlight transactions related to the specific number in the report.</div>',
            unsafe_allow_html=True
        )
        if focus_num:
            st.info(f"🔎 Will highlight: **{focus_num}**")

    with col3:
        st.markdown(
            '<div style="background:#fff;border:1px solid #e8ecf4;border-radius:12px;padding:16px;">'
            '<div style="font-size:13px;font-weight:700;color:#1a2332;margin-bottom:10px;">'
            'What You Will Get</div>', unsafe_allow_html=True
        )
        items = [
            ("📊", "Overall Financial Summary"),
            ("↕", "Credit (In) vs Debit (Out) Analysis"),
            ("≡", "Transaction Type Analysis"),
            ("👥", "Top Senders / Receivers"),
            ("🗂", "Category-wise Breakdown"),
            ("📅", "Daily / Monthly Trend Analysis"),
            ("⚠", "Suspicious Activity Detection"),
            ("📄", "Professional HTML Forensic Report"),
        ]
        for icon, label in items:
            st.markdown(
                f'<div style="display:flex;align-items:center;justify-content:space-between;'
                f'padding:6px 0;border-bottom:1px solid #f4f6f9;font-size:12px;color:#3a4a5c;">'
                f'<span>{icon} {label}</span><span style="color:#9baab8;">›</span></div>',
                unsafe_allow_html=True
            )
        st.markdown('</div>', unsafe_allow_html=True)

    # ── NO FILE — SAMPLE PREVIEW ─────────────────────────────────────────
    if not uploaded:
        st.markdown("---")
        st.markdown(
            '<div style="font-size:15px;font-weight:700;color:#1a2332;margin-bottom:14px;">'
            '📊 Report Preview (Sample Insights)</div>',
            unsafe_allow_html=True
        )
        p1, p2, p3, p4 = st.columns(4)
        p1.markdown(
            '<div style="background:#fff;border:1px solid #e8ecf4;border-radius:10px;padding:16px;">'
            '<div style="font-size:11px;color:#6b7f99;font-weight:600;text-transform:uppercase;">'
            'Total Credit (In)</div>'
            '<div style="font-size:22px;font-weight:800;color:#00875a;margin-top:4px;">৳ 1,64,273.00</div>'
            '<div style="font-size:11px;color:#00875a;margin-top:2px;">↑ Received</div></div>',
            unsafe_allow_html=True
        )
        p2.markdown(
            '<div style="background:#fff;border:1px solid #e8ecf4;border-radius:10px;padding:16px;">'
            '<div style="font-size:11px;color:#6b7f99;font-weight:600;text-transform:uppercase;">'
            'Total Debit (Out)</div>'
            '<div style="font-size:22px;font-weight:800;color:#c0392b;margin-top:4px;">৳ 95,010.00</div>'
            '<div style="font-size:11px;color:#c0392b;margin-top:2px;">↑ Sent</div></div>',
            unsafe_allow_html=True
        )
        p3.markdown(
            '<div style="background:#fff;border:1px solid #e8ecf4;border-radius:10px;padding:16px;">'
            '<div style="font-size:11px;color:#6b7f99;font-weight:600;text-transform:uppercase;">'
            'Net Balance</div>'
            '<div style="font-size:22px;font-weight:800;color:#1557a0;margin-top:4px;">৳ 69,263.00</div>'
            '<div style="font-size:11px;color:#1557a0;margin-top:2px;">⚖ Net flow</div></div>',
            unsafe_allow_html=True
        )
        p4.markdown(
            '<div style="background:#fff;border:1px solid #e8ecf4;border-radius:10px;padding:16px;">'
            '<div style="font-size:11px;color:#6b7f99;font-weight:600;text-transform:uppercase;">'
            'Total Transactions</div>'
            '<div style="font-size:28px;font-weight:800;color:#7c3aed;margin-top:4px;">156</div>'
            '<div style="font-size:11px;color:#7c3aed;margin-top:2px;">📋 Records</div></div>',
            unsafe_allow_html=True
        )

        st.markdown("---")
        st.markdown(
            '<div style="font-size:15px;font-weight:700;color:#1a2332;margin-bottom:14px;">'
            '⭐ Key Capabilities</div>',
            unsafe_allow_html=True
        )
        caps = [
            ("📷","Automatic Data Processing","Auto-detect bKash, Nagad, Rocket"),
            ("📈","Pattern & Trend Detection","Hourly, daily, monthly trends"),
            ("🏷","Smart Categorization","Transaction size & type grouping"),
            ("🔍","Search & Export Data","CSV and HTML report export"),
            ("⚠","Duplicate Detection","Identify repeated transactions"),
            ("👥","Multi-Account Analysis","Focus on specific numbers"),
            ("🔧","Advanced Filtering","Filter by size, type, direction"),
            ("📄","Detailed Forensic Report","Professional HTML & Word output"),
            ("🚨","High Value Transaction Alert","Flag large or suspicious txns"),
            ("🔒","Data Security & Privacy","Processed locally, no data stored"),
        ]
        ca, cb = st.columns(2)
        for i, (icon, title, sub) in enumerate(caps):
            col = ca if i % 2 == 0 else cb
            col.markdown(
                f'<div style="display:flex;align-items:center;gap:10px;background:#fff;'
                f'border:1px solid #e8ecf4;border-radius:8px;padding:12px 14px;'
                f'margin-bottom:8px;box-shadow:0 1px 3px rgba(0,0,0,.04);">'
                f'<span style="font-size:18px;">{icon}</span>'
                f'<div><div style="font-size:12px;font-weight:700;color:#1a2332;">{title}</div>'
                f'<div style="font-size:11px;color:#9baab8;">{sub}</div></div></div>',
                unsafe_allow_html=True
            )

        st.markdown("---")
        f1, f2, f3, f4 = st.columns(4)
        for col, icon, title, desc in [
            (f1,"🔒","Secure & Private","Your data is 100% secure and processed locally."),
            (f2,"⚡","Fast Processing","Get results and reports within seconds."),
            (f3,"🎯","Accurate Insights","Smart analysis with advanced algorithms."),
            (f4,"📄","HTML Report","Professional, printable and easy to share."),
        ]:
            col.markdown(
                f'<div style="text-align:center;padding:14px;">'
                f'<div style="font-size:24px;margin-bottom:8px;">{icon}</div>'
                f'<div style="font-size:13px;font-weight:700;color:#1a2332;">{title}</div>'
                f'<div style="font-size:11px;color:#9baab8;margin-top:4px;">{desc}</div></div>',
                unsafe_allow_html=True
            )
        return

    # =====================================================================
    #  PROCESS FILE
    # =====================================================================
    with st.spinner("Analyzing file, please wait..."):
        ext = uploaded.name.rsplit(".", 1)[-1].lower()
        file_bytes = uploaded.read()

        if ext == "pdf":
            if not PDF_OK:
                st.error("pdfplumber not installed. Run: pip install pdfplumber")
                return
            if not detect_rocket_pdf(file_bytes):
                st.error("This PDF does not appear to be a DBBL Rocket statement.")
                return
            with st.spinner("Extracting data from PDF..."):
                acc, rows = parse_rocket_pdf(file_bytes)
            provider, fmt = "rocket", "rocket"
        else:
            try:
                import io as _io
                buf = _io.BytesIO(file_bytes)
                df_raw = pd.read_csv(buf, header=None, dtype=str) if ext == "csv"                          else pd.read_excel(buf, header=None, dtype=str)
                raw = df_raw.fillna("").values.tolist()
            except Exception as e:
                st.error(f"Cannot read file: {e}")
                return
            provider, fmt = detect_provider(raw, uploaded.name)
            if provider == "unknown":
                st.error("Format not recognized. Upload a bKash, Nagad or Rocket statement.")
                return
            if fmt == "nagad_standard":
                acc, rows = parse_nagad_standard(raw)
            else:
                acc, rows = parse_bkash_nagad_business(raw)
                if "nagad" in uploaded.name.lower():
                    provider = "nagad"

        if not rows:
            st.error("No transactions found.")
            return
        s, main = analyse(rows)
        if s is None:
            st.error("Analysis failed.")
            return

    ac     = ACCENT.get(provider, "#6b7f99")
    pname  = {"nagad":"Nagad","bkash":"bKash","rocket":"Rocket (DBBL)"}.get(provider,"Unknown")
    icon   = {"nagad":"🟠","bkash":"🔴","rocket":"🟣"}.get(provider,"❓")
    fmt_lbl= {"nagad_standard":"Agent/Individual","nagad_business":"Business",
               "bkash_business":"Business Agent","rocket":"Rocket Agent"}.get(fmt, fmt)

    # ── PROVIDER BANNER ──
    st.markdown(
        f'<div class="prov-{provider}">'
        f'<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;">'
        f'<div style="display:flex;align-items:center;gap:14px;">'
        f'<span style="font-size:30px;">{icon}</span>'
        f'<div><div style="font-size:20px;font-weight:800;color:{ac};">{pname}</div>'
        f'<div style="font-size:11px;color:#9baab8;text-transform:uppercase;letter-spacing:1px;">'
        f'{fmt_lbl} · {uploaded.name}</div></div></div>'
        f'<div style="text-align:right;">'
        f'<div style="font-family:monospace;font-size:16px;font-weight:700;color:#1a2332;">{acc or "—"}</div>'
        f'<div style="font-size:12px;color:#9baab8;">{s["start"]} → {s["end"]} · {s["active"]} active days</div>'
        f'</div></div></div>',
        unsafe_allow_html=True
    )

    # ── SUMMARY ──
    night_warn = (f' <span style="color:#c0392b;">⚠ HIGH night activity ({s["night_pct"]}%) — investigate.</span>'
                  if s["night_pct"] > 35 else f' Night activity: <strong>{s["night_pct"]}%</strong>')
    st.markdown(
        f'<div class="sum-box"><strong>{pname}</strong> account <strong>{acc}</strong> — '
        f'<strong>{s["total"]:,} transactions</strong> over <strong>{s["active"]} active days</strong> '
        f'({s["start"]} to {s["end"]}).'
        f' | Credit (In): <strong style="color:#00875a">BDT {s["cr_amt"]:,.2f}</strong> ({s["cr_cnt"]:,})'
        f' | Debit (Out): <strong style="color:#c0392b">BDT {s["db_amt"]:,.2f}</strong> ({s["db_cnt"]:,})'
        f' | Commission: <strong style="color:#b86e00">BDT {s["comm_amt"]:,.2f}</strong>'
        f'{night_warn}</div>',
        unsafe_allow_html=True
    )

    # ── METRICS ──
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Transactions",   f"{s['total']:,}",              "Excl. commission")
    m2.metric("Total Credit (In)",    f"BDT {s['cr_amt']:,.2f}",      f"{s['cr_cnt']:,} txns")
    m3.metric("Total Debit (Out)",    f"BDT {s['db_amt']:,.2f}",      f"{s['db_cnt']:,} txns")
    m4.metric("Commission Earned",    f"BDT {s['comm_amt']:,.2f}",    f"{s['comm_cnt']:,} entries")

    # ── CHARTS ──
    st.markdown("---")
    st.markdown("**📈 Trend & Pattern Analysis**")
    chart = fig_monthly(main, ac) if s["active"] > 7 else fig_daily(main, ac)
    if chart: st.plotly_chart(chart, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        fh = fig_hourly(main, ac)
        if fh: st.plotly_chart(fh, use_container_width=True)
    with c2:
        ft = fig_type_bar(main, ac)
        if ft: st.plotly_chart(ft, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3: st.plotly_chart(fig_size_pie(main), use_container_width=True)
    with c4: st.plotly_chart(fig_direction(main), use_container_width=True)

    # ── CATEGORY BREAKDOWN ──
    st.markdown("---")
    st.markdown("**📊 Category-wise Breakdown**")
    tb1, tb2, tb3 = st.columns(3)
    with tb1:
        st.markdown("**Transaction Size**")
        sz = main.groupby("আকার").agg(Count=("পরিমাণ","count"), Amount=("পরিমাণ","sum")).reset_index()
        sz.columns = ["Size", "Count", "Total (BDT)"]
        sz["Total (BDT)"] = sz["Total (BDT)"].apply(lambda x: f"BDT {x:,.2f}")
        st.dataframe(sz, use_container_width=True, hide_index=True)
    with tb2:
        st.markdown("**Transaction Type**")
        tp = main.groupby("ধরন").agg(Count=("পরিমাণ","count"), Amount=("পরিমাণ","sum"))                 .sort_values("Amount", ascending=False).reset_index()
        tp.columns = ["Type", "Count", "Total (BDT)"]
        tp["Total (BDT)"] = tp["Total (BDT)"].apply(lambda x: f"BDT {x:,.2f}")
        st.dataframe(tp, use_container_width=True, hide_index=True)
    with tb3:
        st.markdown("**Channel**")
        ch = main.groupby("চ্যানেল").agg(Count=("পরিমাণ","count"), Amount=("পরিমাণ","sum")).reset_index()
        ch.columns = ["Channel", "Count", "Total (BDT)"]
        ch["Total (BDT)"] = ch["Total (BDT)"].apply(lambda x: f"BDT {x:,.2f}")
        st.dataframe(ch, use_container_width=True, hide_index=True)

    # ── TOP SENDERS / RECEIVERS ──
    st.markdown("---")
    st.markdown("**👥 Top Senders / Receivers**")
    cp_valid = main[~main["কাউন্টারপার্ট"].isin(["SYSTEM","","nan","None"])]
    tl1, tl2 = st.columns(2)
    with tl1:
        st.markdown("**Top 10 — Highest Amount Received From (Credit)**")
        tr = cp_valid[cp_valid["direction"]=="credit"]               .groupby("কাউন্টারপার্ট").agg(Txns=("পরিমাণ","count"), Amount=("পরিমাণ","sum"))               .sort_values("Amount", ascending=False).head(10).reset_index()
        tr.columns = ["Account Number","Transactions","Total Received (BDT)"]
        tr["Total Received (BDT)"] = tr["Total Received (BDT)"].apply(lambda x: f"BDT {x:,.2f}")
        st.dataframe(tr, use_container_width=True, hide_index=True)
    with tl2:
        st.markdown("**Top 10 — Highest Amount Sent To (Debit)**")
        ts = cp_valid[cp_valid["direction"]=="debit"]               .groupby("কাউন্টারপার্ট").agg(Txns=("পরিমাণ","count"), Amount=("পরিমাণ","sum"))               .sort_values("Amount", ascending=False).head(10).reset_index()
        ts.columns = ["Account Number","Transactions","Total Sent (BDT)"]
        ts["Total Sent (BDT)"] = ts["Total Sent (BDT)"].apply(lambda x: f"BDT {x:,.2f}")
        st.dataframe(ts, use_container_width=True, hide_index=True)

    # ── TOP DAYS ──
    st.markdown("---")
    st.markdown("**📅 Top 10 Busiest Days**")
    td = main.groupby("দিন").agg(Txns=("পরিমাণ","count"), Amount=("পরিমাণ","sum"))             .sort_values("Amount", ascending=False).head(10).reset_index()
    td.columns = ["Date","Transactions","Total Amount (BDT)"]
    td["Total Amount (BDT)"] = td["Total Amount (BDT)"].apply(lambda x: f"BDT {x:,.2f}")
    st.dataframe(td, use_container_width=True, hide_index=True)

    # ── PHONE NUMBER FOCUS ──
    if focus_num:
        st.markdown("---")
        st.markdown(f"**🔎 Number Focus: {focus_num}**")
        fdf = main[main["কাউন্টারপার্ট"].astype(str).str.contains(focus_num, na=False, regex=False)]
        if fdf.empty:
            st.warning(f"No transactions found involving **{focus_num}**.")
        else:
            fcr = fdf[fdf["direction"]=="credit"]
            fdb = fdf[fdf["direction"]=="debit"]
            st.markdown(
                f'<div style="background:#f0faf6;border:1px solid #b3dcc8;border-left:4px solid #00875a;'
                f'border-radius:8px;padding:12px 16px;font-size:13px;margin-bottom:12px;">'
                f'✅ <strong>{len(fdf):,} transactions</strong> found for <strong>{focus_num}</strong> | '
                f'Received: <strong style="color:#00875a">BDT {fcr["পরিমাণ"].sum():,.2f}</strong> ({len(fcr)}) | '
                f'Sent: <strong style="color:#c0392b">BDT {fdb["পরিমাণ"].sum():,.2f}</strong> ({len(fdb)})'
                f'</div>',
                unsafe_allow_html=True
            )
            fc1,fc2,fc3,fc4 = st.columns(4)
            fc1.metric("Total Transactions", f"{len(fdf):,}")
            fc2.metric("Received From", f"BDT {fcr['পরিমাণ'].sum():,.2f}", f"{len(fcr)} txns")
            fc3.metric("Sent To",       f"BDT {fdb['পরিমাণ'].sum():,.2f}", f"{len(fdb)} txns")
            fc4.metric("Net Flow",      f"BDT {fcr['পরিমাণ'].sum()-fdb['পরিমাণ'].sum():,.2f}")
            dcols = [c for c in ["তারিখ","TXN_ID","ধরন","দিক","পরিমাণ","আকার","চ্যানেল","কাউন্টারপার্ট","ব্যালেন্স"] if c in fdf.columns]
            rmap  = {"তারিখ":"Date","TXN_ID":"TXN ID","ধরন":"Type","দিক":"Direction",
                     "পরিমাণ":"Amount","আকার":"Size","চ্যানেল":"Channel","কাউন্টারপার্ট":"Counterpart","ব্যালেন্স":"Balance"}
            st.dataframe(fdf[dcols].rename(columns=rmap).reset_index(drop=True),
                         use_container_width=True, height=min(400,60+35*len(fdf)))

    # ── DOWNLOADS ──
    st.markdown("---")
    st.markdown("**⬇ Download Reports**")
    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        html_data = build_html(s, main, acc, provider, fmt_lbl, uploaded.name, focus_num=focus_num)
        st.download_button("📄 Download HTML Report",
            data=html_data.encode("utf-8"),
            file_name=f"mfs_report_{acc}_{provider}.html",
            mime="text/html", use_container_width=True)
    with dl2:
        if DOCX_OK:
            ddata = build_docx(s, main, acc, provider, fmt_lbl, uploaded.name, focus_num=focus_num)
            if ddata:
                st.download_button("📝 Download Word Report",
                    data=ddata,
                    file_name=f"mfs_report_{acc}_{provider}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True)
        else:
            st.info("Word export: `pip install python-docx`")
    with dl3:
        dcols2 = [c for c in ["তারিখ","TXN_ID","ধরন","দিক","পরিমাণ","আকার","চ্যানেল","কাউন্টারপার্ট","ব্যালেন্স"] if c in main.columns]
        rmap2  = {"তারিখ":"Date","TXN_ID":"TXN ID","ধরন":"Type","দিক":"Direction",
                  "পরিমাণ":"Amount","আকার":"Size","চ্যানেল":"Channel","কাউন্টারপার্ট":"Counterpart","ব্যালেন্স":"Balance"}
        st.download_button("📊 Download CSV",
            data=main[dcols2].rename(columns=rmap2).to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
            file_name=f"mfs_transactions_{acc}.csv",
            mime="text/csv", use_container_width=True)

    st.markdown(
        f'<div style="text-align:center;font-size:11px;color:#9baab8;padding:14px 0;'
        f'border-top:1px solid #e8ecf4;margin-top:12px;">'
        f'Mobile Financial System Analyzer · {pname} · Account: {acc} · Bangladesh</div>',
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main_app()
