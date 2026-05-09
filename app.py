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
    hi,acc=-1,''
    for i,row in enumerate(raw[:12]):
        for c in row:
            if re.match(r'^01[0-9]{9}$',str(c).strip()): acc=str(c).strip()
        rlow=' '.join(str(c) for c in row).lower()
        if 'transaction id' in rlow and ('sender' in rlow or 'transaction type' in rlow):
            hi=i; break
    if hi<0: hi=4
    H=[str(h).replace('\n',' ').strip().lower() for h in raw[hi]]
    def c(kw): return next((i for i,h in enumerate(H) if kw in h),-1)
    iDt=c('transaction date');iT=c('transaction id');iTy=c('transaction type')
    iSn=c('sender');iRv=c('receiver');iA=c('transacted amount')
    iCm=c('commission');iBl=c('balance')

    rows=[]
    for row in raw[hi+1:]:
        if not row or all(str(x).strip() in ('','nan','None') for x in row): continue
        amt=_f(row[iA]) if iA>=0 else 0
        if amt==0: continue
        ty=str(row[iTy]).strip() if iTy>=0 else ''
        sn=str(row[iSn]).strip() if iSn>=0 else ''
        rv=str(row[iRv]).strip() if iRv>=0 else ''
        cm=_f(row[iCm]) if iCm>=0 else 0
        if 'cash out' in ty.lower(): dr,cp='credit',sn
        elif 'cash in' in ty.lower(): dr,cp='debit',rv
        else: dr,cp='credit',sn or rv
        dt=_date(row[iDt] if iDt>=0 else '')
        rows.append({
            'তারিখ':dt[:16],'TXN_ID':str(row[iT]).strip() if iT>=0 else '',
            'ধরন':ty,'direction':dr,
            'দিক':'গ্রহণ' if dr=='credit' else 'পাঠানো',
            'পরিমাণ':amt,'ব্যালেন্স':_f(row[iBl]) if iBl>=0 else 0,
            'চ্যানেল':'USSD' if 'ussd' in ty.lower() else 'API',
            'কাউন্টারপার্ট':cp,'স্ট্যাটাস':'APPROVED',
            'is_comm':False,'আকার':size_cat(amt),
            'দিন':dt[:10],'মাস':dt[:7],'কমিশন':cm,
        })
    return acc,rows


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

    # ── Transaction table (first 1000) ──
    show   = min(len(main), 1000)
    txn_df = main[['তারিখ','TXN_ID','ধরন','দিক','পরিমাণ','আকার','চ্যানেল','কাউন্টারপার্ট']].head(show).copy()
    txn_df['পরিমাণ'] = txn_df['পরিমাণ'].apply(lambda x: f"BDT {x:,.2f}")

    # ── Chart.js data helpers ──
    def js_list(series):   return '[' + ','.join(f'"{v}"' for v in series) + ']'
    def js_nums(series):   return '[' + ','.join(f'{v:.2f}' for v in series) + ']'
    def js_ints(series):   return '[' + ','.join(str(int(v)) for v in series) + ']'

    # Data for Top-10 Received chart
    recv_labels = js_list(top_recv['কাউন্টারপার্ট'])
    recv_amts   = js_nums(top_recv['amt'])
    recv_cnts   = js_ints(top_recv['cnt'])

    # Data for Top-10 Sent chart
    sent_labels = js_list(top_sent['কাউন্টারপার্ট'])
    sent_amts   = js_nums(top_sent['amt'])
    sent_cnts   = js_ints(top_sent['cnt'])

    # Data for Monthly chart
    mo_labels = js_list(mo['মাস'])
    mo_amts   = js_nums(mo['amt'])
    mo_cnts   = js_ints(mo['cnt'])

    # Hourly data
    try:
        m2 = main.copy()
        m2['hr'] = pd.to_datetime(m2['তারিখ'], errors='coerce').dt.hour.fillna(0).astype(int)
        hr_ser = m2.groupby('hr')['পরিমাণ'].sum().reindex(range(24), fill_value=0)
        hourly_amts = js_nums(hr_ser)
        hourly_labels = js_list([f"{h}:00" for h in range(24)])
        hourly_colors = '[' + ','.join(
            f'"{"#c0392b" if h in [22,23,0,1] else ac}"' for h in range(24)) + ']'
    except:
        hourly_amts = js_nums([0]*24)
        hourly_labels = js_list([f"{h}:00" for h in range(24)])
        hourly_colors = f'["{ac}"]'

    # ── HTML table helpers ──
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
    txn_rows = ''.join(
        f'<tr><td>{i+1}</td>' + ''.join(f'<td>{row[c]}</td>' for c in txn_df.columns) + '</tr>'
        for i,(_,row) in enumerate(txn_df.iterrows()))

    warn_top = (f'<div class="warn">⚠️ <strong>{top_cp.iloc[0]["কাউন্টারপার্ট"]}</strong> — '
                f'{int(top_cp.iloc[0]["cnt"])} transactions. Unusual activity.</div>'
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

  {focus_html}

  <!-- ═══════ BREAKDOWN TABLES ═══════ -->
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

  <!-- ═══════ TOP COUNTERPARTS + TOP DAYS ═══════ -->
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

  <!-- ═══════ TRANSACTION TABLE ═══════ -->
  <div class="panel">
    <div class="pt">Transaction Detail List (First {show} Records)</div>
    <div style="overflow-x:auto;">
      <table class="tbl">
        <thead>
          <tr><th>SL</th><th>Date</th><th>TXN ID</th><th>Type</th><th>Direction</th>
              <th>Amount</th><th>Size</th><th>Channel</th><th>Counterpart</th></tr>
        </thead>
        <tbody>{txn_rows}</tbody>
      </table>
    </div>
  </div>

</div>

<div class="footer">
  MFS Forensic Report &nbsp;·&nbsp; {pname} &nbsp;·&nbsp; Account: {acc} &nbsp;·&nbsp; Generated: {now}
</div>

<!-- ═══════ CHART.JS SCRIPTS ═══════ -->
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
def build_docx(s, main, acc, provider, fmt, fname):
    if not DOCX_OK:
        return None

    doc = Document()
    ac_rgb = {
        'nagad':  RGBColor(0xe8, 0x5d, 0x04),
        'bkash':  RGBColor(0xe2, 0x13, 0x6e),
        'rocket': RGBColor(0x7c, 0x3a, 0xed),
    }.get(provider, RGBColor(0x6b, 0x7f, 0x99))

    pname = {'nagad':'নগদ (Nagad)','bkash':'বিকাশ (bKash)','rocket':'রকেট (Rocket)'}.get(provider,'অজানা')
    now   = datetime.now().strftime('%d %B %Y, %I:%M %p')

    # ── Page margins ──
    for sec in doc.sections:
        sec.top_margin    = Cm(1.8)
        sec.bottom_margin = Cm(1.8)
        sec.left_margin   = Cm(2.0)
        sec.right_margin  = Cm(2.0)

    def add_heading(text, level=1, color=None):
        p = doc.add_heading('', level=level)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        run.font.size  = Pt(18 if level==1 else 13 if level==2 else 11)
        run.font.bold  = True
        if color: run.font.color.rgb = color
        return p

    def add_para(text, bold=False, size=11, color=None):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        if color: run.font.color.rgb = color
        return p

    def add_table(headers, data_rows, col_widths=None):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        hrow = t.rows[0]
        for i, h in enumerate(headers):
            cell = hrow.cells[i]
            cell.text = h
            cell.paragraphs[0].runs[0].font.bold  = True
            cell.paragraphs[0].runs[0].font.size  = Pt(9)
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x6b, 0x7f, 0x99)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for row_data in data_rows:
            row = t.add_row()
            for i, val in enumerate(row_data):
                row.cells[i].text = str(val)
                row.cells[i].paragraphs[0].runs[0].font.size = Pt(9)
        return t

    # ── Title ──
    add_heading(f"Mobile Financial System — Forensic Report", 1, ac_rgb)
    add_para(f"প্রদানকারী: {pname}   |   অ্যাকাউন্ট: {acc or '—'}   |   তৈরি: {now}", size=10,
             color=RGBColor(0x6b, 0x7f, 0x99))
    add_para(f"ফাইল: {fname}   |   মেয়াদ: {s['start']} → {s['end']}   |   সক্রিয় দিন: {s['active']}", size=10,
             color=RGBColor(0x6b, 0x7f, 0x99))
    doc.add_paragraph()

    # ── Summary ──
    add_heading("সারসংক্ষেপ", 2, ac_rgb)
    summary_text = (
        f"এই {pname} অ্যাকাউন্টে {s['active']} দিনে মোট {s['total']:,}টি লেনদেন সম্পন্ন হয়েছে। "
        f"মোট গ্রহণ ৳{s['cr_amt']:,.2f} ({s['cr_cnt']:,}টি), "
        f"মোট পাঠানো ৳{s['db_amt']:,.2f} ({s['db_cnt']:,}টি)। "
        f"কমিশন উপার্জন ৳{s['comm_amt']:,.2f}। "
        f"গড় দৈনিক পরিমাণ ৳{s['avg_amt']:,.2f}। "
        f"রাত্রিকালীন কার্যক্রম (রাত ১০–ভোর ২): {s['night_pct']}%।"
    )
    add_para(summary_text, size=11)
    doc.add_paragraph()

    # ── Key Metrics Table ──
    add_heading("মূল পরিসংখ্যান", 2, ac_rgb)
    add_table(
        ['বিষয়', 'পরিমাণ', 'সংখ্যা'],
        [
            ['মোট লেনদেন', '—', f"{s['total']:,}টি"],
            ['মোট গ্রহণ (Credit)', f"৳{s['cr_amt']:,.2f}", f"{s['cr_cnt']:,}টি"],
            ['মোট পাঠানো (Debit)', f"৳{s['db_amt']:,.2f}", f"{s['db_cnt']:,}টি"],
            ['কমিশন উপার্জন', f"৳{s['comm_amt']:,.2f}", f"{s['comm_cnt']:,}টি"],
            ['গড় দৈনিক পরিমাণ', f"৳{s['avg_amt']:,.2f}", f"≈{s['avg_cnt']}টি/দিন"],
            ['রাত্রিকালীন কার্যক্রম', f"{s['night_pct']}%", '(রাত ১০–ভোর ২)'],
        ]
    )
    doc.add_paragraph()

    # ── Monthly ──
    add_heading("মাসিক বিশ্লেষণ", 2, ac_rgb)
    mo = main.groupby('মাস').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum')).reset_index()
    add_table(
        ['মাস', 'লেনদেন সংখ্যা', 'মোট পরিমাণ (৳)'],
        [[r['মাস'], f"{int(r['cnt']):,}", f"৳{r['amt']:,.2f}"] for _,r in mo.iterrows()]
    )
    doc.add_paragraph()

    # ── Size Breakdown ──
    add_heading("লেনদেনের আকার বিভাজন", 2, ac_rgb)
    sz = main.groupby('আকার').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum')).reset_index()
    add_table(
        ['আকার', 'সংখ্যা', 'মোট পরিমাণ (৳)'],
        [[r['আকার'], f"{int(r['cnt']):,}", f"৳{r['amt']:,.2f}"] for _,r in sz.iterrows()]
    )
    doc.add_paragraph()

    # ── Type Breakdown ──
    add_heading("লেনদেনের ধরন", 2, ac_rgb)
    tp = main.groupby('ধরন').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
             .sort_values('amt', ascending=False).head(8).reset_index()
    add_table(
        ['ধরন', 'সংখ্যা', 'মোট পরিমাণ (৳)'],
        [[r['ধরন'][:35], f"{int(r['cnt']):,}", f"৳{r['amt']:,.2f}"] for _,r in tp.iterrows()]
    )
    doc.add_paragraph()

    # ── Top Counterparts ──
    add_heading("শীর্ষ ১০ কাউন্টারপার্ট অ্যাকাউন্ট", 2, ac_rgb)
    cp = main[~main['কাউন্টারপার্ট'].isin(['SYSTEM','','nan','None'])]
    top_cp = cp.groupby('কাউন্টারপার্ট').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
               .sort_values('amt', ascending=False).head(10).reset_index()
    add_table(
        ['#', 'অ্যাকাউন্ট নম্বর', 'লেনদেন সংখ্যা', 'মোট পরিমাণ (৳)'],
        [[i+1, r['কাউন্টারপার্ট'], f"{int(r['cnt']):,}", f"৳{r['amt']:,.2f}"]
         for i,(_,r) in enumerate(top_cp.iterrows())]
    )
    doc.add_paragraph()

    # ── Top Days ──
    add_heading("শীর্ষ ১০ ব্যস্ত দিন", 2, ac_rgb)
    td = main.groupby('দিন').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
             .sort_values('amt', ascending=False).head(10).reset_index()
    add_table(
        ['#', 'তারিখ', 'লেনদেন সংখ্যা', 'মোট পরিমাণ (৳)'],
        [[i+1, r['দিন'], f"{int(r['cnt']):,}", f"৳{r['amt']:,.2f}"]
         for i,(_,r) in enumerate(td.iterrows())]
    )
    doc.add_paragraph()

    # ── Transaction List (first 500) ──
    add_heading("লেনদেনের বিস্তারিত তালিকা (প্রথম ৫০০টি)", 2, ac_rgb)
    show = main[['তারিখ','TXN_ID','ধরন','দিক','পরিমাণ','আকার','চ্যানেল','কাউন্টারপার্ট']].head(500)
    add_table(
        ['SL','তারিখ','TXN ID','ধরন','দিক','পরিমাণ','আকার','চ্যানেল','কাউন্টারপার্ট'],
        [[i+1, r['তারিখ'], r['TXN_ID'], r['ধরন'][:20], r['দিক'],
          f"৳{r['পরিমাণ']:,.2f}", r['আকার'], r['চ্যানেল'], r['কাউন্টারপার্ট']]
         for i,(_,r) in enumerate(show.iterrows())]
    )

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
    df = main.groupby('মাস').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum')).reset_index()
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=df['মাস'], y=df['amt'], name='পরিমাণ (৳)',
                         marker_color=acc, marker_line_width=0, opacity=0.85), secondary_y=False)
    fig.add_trace(go.Scatter(x=df['মাস'], y=df['cnt'], name='লেনদেন',
                             mode='lines+markers', line=dict(color='#c0392b',width=2),
                             marker=dict(size=5,color='#c0392b')), secondary_y=True)
    fig.update_layout(**ch_layout('মাসিক লেনদেন বিশ্লেষণ'))
    fig.update_yaxes(title_text="পরিমাণ (৳)", secondary_y=False)
    fig.update_yaxes(title_text="সংখ্যা", secondary_y=True)
    return fig

def fig_daily(main, acc):
    df = main[main['দিন'] != ''].groupby('দিন')['পরিমাণ'].sum().reset_index()
    if df.empty:
        return None
    fig = px.bar(df, x='দিন', y='পরিমাণ', color_discrete_sequence=[acc])
    fig.update_traces(marker_line_width=0, marker_color=acc, opacity=0.85)
    fig.update_layout(**ch_layout('Daily Transaction Volume'))
    return fig

def fig_hourly(main, acc):
    try:
        m2 = main.copy()
        m2['ঘণ্টা'] = pd.to_datetime(m2['তারিখ'], errors='coerce').dt.hour.fillna(0).astype(int)
        df = m2.groupby('ঘণ্টা')['পরিমাণ'].sum().reindex(range(24), fill_value=0).reset_index()
        df.columns = ['ঘণ্টা', 'পরিমাণ']
        df['রঙ'] = df['ঘণ্টা'].apply(lambda h: '#c0392b' if h in [22,23,0,1] else acc)
        fig = go.Figure(go.Bar(x=df['ঘণ্টা'], y=df['পরিমাণ'],
                               marker_color=df['রঙ'], marker_line_width=0))
        fig.update_layout(**ch_layout('ঘণ্টাভিত্তিক কার্যক্রম (লাল = রাত ১০–ভোর ২)'))
        return fig
    except: return None

def fig_size_pie(main):
    df = main.groupby('আকার')['পরিমাণ'].sum().reset_index()
    colors = {'ছোট (২-৮হা)':'#1557a0','মধ্যম (৮-২০হা)':'#b86e00','উচ্চ (২০হা+)':'#c0392b','<২হা':'#b0b8c8'}
    fig = px.pie(df, values='পরিমাণ', names='আকার',
                 color='আকার', color_discrete_map=colors, hole=0.55)
    fig.update_layout(**ch_layout('লেনদেনের আকার', h=280))
    fig.update_traces(textposition='outside', textinfo='percent+label')
    return fig

def fig_direction(main):
    df = main.groupby('দিক')['পরিমাণ'].sum().reset_index()
    fig = px.pie(df, values='পরিমাণ', names='দিক', hole=0.55,
                 color='দিক', color_discrete_map={'গ্রহণ':'#00875a','পাঠানো':'#c0392b'})
    fig.update_layout(**ch_layout('গ্রহণ vs পাঠানো', h=280))
    return fig

def fig_type_bar(main, acc):
    df = main.groupby('ধরন').agg(cnt=('পরিমাণ','count'), amt=('পরিমাণ','sum'))\
             .sort_values('amt', ascending=False).head(6).reset_index()
    fig = px.bar(df, x='amt', y='ধরন', orientation='h',
                 color_discrete_sequence=[acc], text='cnt')
    fig.update_traces(marker_line_width=0, texttemplate='%{text}টি', textposition='inside')
    fig.update_layout(**ch_layout('লেনদেনের ধরন (শীর্ষ ৬)', h=260))
    return fig

try:
    from plotly.subplots import make_subplots
except: pass


# ════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ════════════════════════════════════════════════════════════════════════════
def main_app():
    # ── HERO BANNER ──────────────────────────────────────────────────────
    st.markdown("""
    <div class="hero">
      <h1>Mobile Financial System Analyzer</h1>
      <p>Upload mobile financial statements (Excel) and get comprehensive insights,<br>
         analytics, and professional forensic reports instantly.</p>
      <div class="hero-badges">
        <div class="hero-badge">🔴 bKash</div>
        <div class="hero-badge">🟠 নগদ</div>
        <div class="hero-badge">🟣 Rocket (PDF)</div>
      </div>
      <div class="feature-pills">
        <div class="fpill">📊 Financial Analysis</div>
        <div class="fpill">🔄 Transaction Analytics</div>
        <div class="fpill">👥 Top Sender/Receiver</div>
        <div class="fpill">📈 Trend & Pattern Detection</div>
        <div class="fpill">📄 HTML &amp; Word Reports</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── UPLOAD + WHAT YOU GET ─────────────────────────────────────────────
    st.markdown('<div class="sec-wrap">', unsafe_allow_html=True)
    col_up, col_wyg = st.columns([2, 1])

    with col_up:
        st.markdown("""
        <div class="upload-card">
          <h3>📤 Upload Statement (Excel / PDF)</h3>
          <p>Upload bKash / Nagad (Excel) or Rocket (PDF) statement.</p>
        </div>
        """, unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "", type=['xlsx','xls','csv','pdf'],
            help="Max 50MB · XLSX / XLS / CSV (bKash/Nagad) · PDF (Rocket)"
        )
        if uploaded:
            st.success(f"✅ {uploaded.name} ({uploaded.size/1024:.1f} KB)")

    with col_wyg:
        st.markdown("""
        <div style="background:#fff;border:1px solid #e8ecf4;border-radius:14px;padding:20px;">
          <div style="font-size:14px;font-weight:700;color:#1a2332;margin-bottom:12px;">⭐ What You Will Get</div>
          <div class="wyg-item"><div class="wyg-icon">📊</div> Overall Financial Summary</div>
          <div class="wyg-item"><div class="wyg-icon">🔄</div> Credit (In) vs Debit (Out) Analysis</div>
          <div class="wyg-item"><div class="wyg-icon">📋</div> Transaction Type Analysis</div>
          <div class="wyg-item"><div class="wyg-icon">👥</div> Top Senders / Receivers</div>
          <div class="wyg-item"><div class="wyg-icon">🗂️</div> Category-wise Breakdown</div>
          <div class="wyg-item"><div class="wyg-icon">📅</div> Daily / Monthly Trend Analysis</div>
          <div class="wyg-item"><div class="wyg-icon">⚠️</div> Suspicious Activity Detection</div>
          <div class="wyg-item"><div class="wyg-icon">📄</div> HTML &amp; Word Forensic Report</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    if not uploaded:
        # ── SAMPLE PREVIEW ────────────────────────────────────────────────
        st.markdown('<div class="sec-wrap"><div class="sec-title">📊 Report Preview (Sample Insights)</div>', unsafe_allow_html=True)
        pc1, pc2, pc3, pc4 = st.columns(4)
        pc1.markdown('<div class="preview-card"><div class="pc-label">Total Credit (In)</div><div class="pc-value pc-green">৳1,64,273.00 ↑</div></div>', unsafe_allow_html=True)
        pc2.markdown('<div class="preview-card"><div class="pc-label">Total Debit (Out)</div><div class="pc-value pc-red">৳95,010.00 ↑</div></div>', unsafe_allow_html=True)
        pc3.markdown('<div class="preview-card"><div class="pc-label">Net Balance</div><div class="pc-value pc-blue">৳69,263.00 ⚖️</div></div>', unsafe_allow_html=True)
        pc4.markdown('<div class="preview-card"><div class="pc-label">Total Transactions</div><div class="pc-value pc-purple">156 📋</div></div>', unsafe_allow_html=True)

        st.markdown("""
        <div style="background:#f9fafb;border:1px solid #e8ecf4;border-radius:10px;padding:20px 28px;margin-top:16px;text-align:center;">
          <div style="font-size:28px;margin-bottom:8px;">📥</div>
          <div style="font-size:14px;font-weight:600;color:#1a2332;">উপরে XLSX ফাইল আপলোড করুন বিশ্লেষণ শুরু করতে</div>
          <div style="font-size:12px;color:#6b7f99;margin-top:4px;">Supported: bKash (XLSX) · Nagad (XLSX) · Rocket (PDF) · Max 50MB</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("""
        <div style="padding:0 36px;">
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:12px;">
          <div style="background:#fff;border:1px solid #e8ecf4;border-radius:10px;padding:16px;text-align:center;">
            <div style="font-size:24px;margin-bottom:6px;">🔒</div>
            <div style="font-size:13px;font-weight:700;">Secure & Private</div>
            <div style="font-size:11px;color:#6b7f99;margin-top:3px;">Your data is 100% secure and processed locally.</div>
          </div>
          <div style="background:#fff;border:1px solid #e8ecf4;border-radius:10px;padding:16px;text-align:center;">
            <div style="font-size:24px;margin-bottom:6px;">⚡</div>
            <div style="font-size:13px;font-weight:700;">Fast Processing</div>
            <div style="font-size:11px;color:#6b7f99;margin-top:3px;">Get results and reports within seconds.</div>
          </div>
          <div style="background:#fff;border:1px solid #e8ecf4;border-radius:10px;padding:16px;text-align:center;">
            <div style="font-size:24px;margin-bottom:6px;">🎯</div>
            <div style="font-size:13px;font-weight:700;">Accurate Insights</div>
            <div style="font-size:11px;color:#6b7f99;margin-top:3px;">Smart analysis with advanced algorithms.</div>
          </div>
          <div style="background:#fff;border:1px solid #e8ecf4;border-radius:10px;padding:16px;text-align:center;">
            <div style="font-size:24px;margin-bottom:6px;">📄</div>
            <div style="font-size:13px;font-weight:700;">HTML & Word Report</div>
            <div style="font-size:11px;color:#6b7f99;margin-top:3px;">Professional, printable and easy to share.</div>
          </div>
        </div></div>
        """, unsafe_allow_html=True)
        return

    # ════════════════════════════════════════════════════════════════════
    #  PROCESS FILE
    # ════════════════════════════════════════════════════════════════════
    with st.spinner("🔍 Analyzing file..."):
        ext = uploaded.name.rsplit('.',1)[-1].lower()
        file_bytes = uploaded.read()

        # ── Rocket PDF branch ──────────────────────────────────────────
        if ext == 'pdf':
            if not PDF_OK:
                st.error("❌ pdfplumber not installed. Run: pip install pdfplumber")
                return
            if not detect_rocket_pdf(file_bytes):
                st.error("❌ This PDF does not appear to be a DBBL Rocket statement.")
                return
            with st.spinner("📄 Extracting data from PDF (may take a moment for large files)..."):
                acc, rows = parse_rocket_pdf(file_bytes)
            provider, fmt = 'rocket', 'rocket'

        # ── Excel / CSV branch ─────────────────────────────────────────
        else:
            try:
                import io as _io
                buf = _io.BytesIO(file_bytes)
                if ext == 'csv':
                    df_raw = pd.read_csv(buf, header=None, dtype=str)
                else:
                    df_raw = pd.read_excel(buf, header=None, dtype=str)
                raw = df_raw.fillna('').values.tolist()
            except Exception as e:
                st.error(f"❌ Cannot read file: {e}"); return

            provider, fmt = detect_provider(raw, uploaded.name)
            if provider == 'unknown':
                st.error("❌ Format not recognized. Please upload a bKash, Nagad or Rocket statement."); return

            if fmt == 'nagad_standard':
                acc, rows = parse_nagad_standard(raw)
            else:
                acc, rows = parse_bkash_nagad_business(raw)
                if 'nagad' in uploaded.name.lower(): provider = 'nagad'

        if not rows:
            st.error("❌ No transactions found in this file."); return

        s, main = analyse(rows)
        if s is None:
            st.error("❌ Analysis failed."); return

    ac       = ACCENT.get(provider, '#6b7f99')
    pname    = {'nagad':'নগদ (Nagad)','bkash':'বিকাশ (bKash)','rocket':'রকেট (Rocket)'}.get(provider,'অজানা')
    icon     = {'nagad':'🟠','bkash':'🔴','rocket':'🟣'}.get(provider,'❓')
    fmt_lbl  = {'nagad_standard':'Agent/Individual','nagad_business':'Business','bkash_business':'Business Agent','rocket':'Rocket'}.get(fmt, fmt)

    # ── PROVIDER BANNER ──
    st.markdown(f"""
    <div class="prov-{provider}">
      <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;">
        <div style="display:flex;align-items:center;gap:14px;">
          <span style="font-size:34px;">{icon}</span>
          <div>
            <div style="font-size:20px;font-weight:800;color:{'#e85d04' if provider=='nagad' else '#e2136e' if provider=='bkash' else '#7c3aed'};">{pname}</div>
            <div style="font-size:11px;color:#6b7f99;text-transform:uppercase;letter-spacing:1px;">{fmt_lbl} · {uploaded.name}</div>
          </div>
        </div>
        <div style="font-family:monospace;font-size:17px;font-weight:700;color:#1a2332;">{acc or '—'}</div>
        <div style="font-size:12px;color:#6b7f99;">{s['start']} → {s['end']} &nbsp;·&nbsp; সক্রিয় {s['active']} দিন</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── SUMMARY ──
    st.markdown(f"""
    <div class="sum-box">
      📋 <strong>{pname}</strong> অ্যাকাউন্টে <strong>{s['active']} দিনে</strong>
      মোট <strong>{s['total']:,}টি</strong> লেনদেন।
      মোট গ্রহণ <strong style="color:#00875a">৳{s['cr_amt']:,.2f}</strong> ({s['cr_cnt']:,}টি) ·
      মোট পাঠানো <strong style="color:#c0392b">৳{s['db_amt']:,.2f}</strong> ({s['db_cnt']:,}টি) ·
      কমিশন <strong style="color:#b86e00">৳{s['comm_amt']:,.2f}</strong>।
      {'<span style="color:#c0392b"> ⚠️ রাত্রিকালীন কার্যক্রম ' + str(s["night_pct"]) + '% — তদন্তযোগ্য।</span>' if s['night_pct']>35 else f'রাত্রিকালীন: <strong>{s["night_pct"]}%</strong>।'}
    </div>
    """, unsafe_allow_html=True)

    # ── METRICS ──
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("🔢 মোট লেনদেন",    f"{s['total']:,}",           "কমিশন বাদে")
    m2.metric("⬇️ মোট গ্রহণ",    bdt(s['cr_amt']),            f"{s['cr_cnt']:,}টি")
    m3.metric("⬆️ মোট পাঠানো",   bdt(s['db_amt']),            f"{s['db_cnt']:,}টি")
    m4.metric("💰 কমিশন",         bdt(s['comm_amt']),          f"{s['comm_cnt']:,}টি")

    st.markdown("---")

    # ── CHARTS ──
    st.markdown('<div class="sec-wrap"><div class="sec-title">📈 Trend & Pattern Analysis</div>', unsafe_allow_html=True)

    if s['active'] > 7:
        fchart = fig_monthly(main, ac)
    else:
        fchart = fig_daily(main, ac)
    if fchart: st.plotly_chart(fchart, use_container_width=True)

    ch1, ch2 = st.columns(2)
    with ch1:
        fh = fig_hourly(main, ac)
        if fh: st.plotly_chart(fh, use_container_width=True)
    with ch2:
        st.plotly_chart(fig_type_bar(main, ac), use_container_width=True)

    ch3, ch4 = st.columns(2)
    with ch3:
        st.plotly_chart(fig_size_pie(main), use_container_width=True)
    with ch4:
        st.plotly_chart(fig_direction(main), use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("---")

    # ── BREAKDOWN TABLES ──
    st.markdown('<div class="sec-wrap"><div class="sec-title">📊 Category-wise Breakdown</div>', unsafe_allow_html=True)
    tb1, tb2, tb3 = st.columns(3)
    with tb1:
        st.markdown("**লেনদেনের আকার**")
        sz = main.groupby('আকার').agg(সংখ্যা=('পরিমাণ','count'),মোট=('পরিমাণ','sum')).reset_index()
        sz['মোট'] = sz['মোট'].apply(lambda x: f"৳{x:,.2f}")
        st.dataframe(sz, use_container_width=True, hide_index=True)
    with tb2:
        st.markdown("**লেনদেনের ধরন**")
        tp = main.groupby('ধরন').agg(সংখ্যা=('পরিমাণ','count'),মোট=('পরিমাণ','sum'))\
                 .sort_values('মোট',ascending=False).reset_index()
        tp['মোট'] = tp['মোট'].apply(lambda x: f"৳{x:,.2f}")
        st.dataframe(tp, use_container_width=True, hide_index=True)
    with tb3:
        st.markdown("**চ্যানেল**")
        ch = main.groupby('চ্যানেল').agg(সংখ্যা=('পরিমাণ','count'),মোট=('পরিমাণ','sum')).reset_index()
        ch['মোট'] = ch['মোট'].apply(lambda x: f"৳{x:,.2f}")
        st.dataframe(ch, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── TOP LISTS ──
    st.markdown("---")
    st.markdown('<div class="sec-wrap"><div class="sec-title">👥 Top Senders / Receivers</div>', unsafe_allow_html=True)
    tl1, tl2 = st.columns(2)
    with tl1:
        st.markdown("**শীর্ষ ১০ কাউন্টারপার্ট অ্যাকাউন্ট**")
        cp = main[~main['কাউন্টারপার্ট'].isin(['SYSTEM','','nan','None'])]
        top = cp.groupby('কাউন্টারপার্ট').agg(লেনদেন=('পরিমাণ','count'),মোট=('পরিমাণ','sum'))\
                .sort_values('মোট',ascending=False).head(10).reset_index()
        top['মোট'] = top['মোট'].apply(lambda x: f"৳{x:,.2f}")
        st.dataframe(top, use_container_width=True, hide_index=True)
        if not top.empty and int(top.iloc[0]['লেনদেন']) > 30:
            st.markdown(f'<div class="warn-box">⚠️ <strong>{top.iloc[0]["কাউন্টারপার্ট"]}</strong> — {int(top.iloc[0]["লেনদেন"])} বার লেনদেন। সন্দেহজনক।</div>', unsafe_allow_html=True)
    with tl2:
        st.markdown("**শীর্ষ ১০ ব্যস্ত দিন**")
        td2 = main.groupby('দিন').agg(লেনদেন=('পরিমাণ','count'),মোট=('পরিমাণ','sum'))\
                  .sort_values('মোট',ascending=False).head(10).reset_index()
        td2.columns = ['তারিখ','লেনদেন','মোট']
        td2['মোট'] = td2['মোট'].apply(lambda x: f"৳{x:,.2f}")
        st.dataframe(td2, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── PHONE NUMBER FOCUS (optional) ─────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="sec-wrap"><div class="sec-title">🔎 Focus on a Specific Number (Optional)</div>', unsafe_allow_html=True)
    st.caption("Enter any phone/account number to see ONLY transactions involving that number with this account.")
    focus_num = st.text_input("📱 Phone / Account Number", value="",
                              placeholder="e.g. 01897606001",
                              help="Leave blank to skip. Partial numbers also work.").strip()

    if focus_num:
        # Filter transactions where counterpart contains the entered number
        focus_df = main[main['কাউন্টারপার্ট'].astype(str).str.contains(focus_num, na=False, regex=False)]

        if focus_df.empty:
            st.warning(f"⚠️ No transactions found involving **{focus_num}** in this statement.")
        else:
            f_cr = focus_df[focus_df['direction'] == 'credit']
            f_db = focus_df[focus_df['direction'] == 'debit']
            f_first = focus_df['তারিখ'].min() if not focus_df.empty else '—'
            f_last  = focus_df['তারিখ'].max() if not focus_df.empty else '—'

            # Status banner
            st.markdown(f"""
            <div style="background:#f0faf6;border:1px solid #b3dcc8;border-left:4px solid #00875a;border-radius:8px;padding:14px 18px;font-size:13px;line-height:1.8;margin-bottom:16px;">
              ✅ Found <strong>{len(focus_df):,} transactions</strong> involving <strong style="font-family:monospace;color:#1a2332;">{focus_num}</strong>.<br>
              <strong style="color:#00875a">Received from this number:</strong> BDT {f_cr['পরিমাণ'].sum():,.2f} ({len(f_cr)} txns) &nbsp;|&nbsp;
              <strong style="color:#c0392b">Sent to this number:</strong> BDT {f_db['পরিমাণ'].sum():,.2f} ({len(f_db)} txns)<br>
              First seen: {f_first} &nbsp;·&nbsp; Last seen: {f_last}
            </div>
            """, unsafe_allow_html=True)

            # Metric cards
            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("🔢 Total Transactions", f"{len(focus_df):,}")
            mc2.metric("⬇️ Received From",      bdt(f_cr['পরিমাণ'].sum()), f"{len(f_cr)} txns")
            mc3.metric("⬆️ Sent To",            bdt(f_db['পরিমাণ'].sum()), f"{len(f_db)} txns")
            mc4.metric("⚖️ Net Flow",           bdt(f_cr['পরিমাণ'].sum() - f_db['পরিমাণ'].sum()))

            # Show full transaction list for this number
            st.markdown("**📋 All Transactions with this Number:**")
            show_cols = [c for c in ['তারিখ','TXN_ID','ধরন','দিক','পরিমাণ','আকার','চ্যানেল','কাউন্টারপার্ট','ব্যালেন্স']
                         if c in focus_df.columns]
            st.dataframe(focus_df[show_cols].reset_index(drop=True),
                         use_container_width=True, height=min(400, 50 + 35*len(focus_df)))
    else:
        st.info("💡 Tip: Enter a phone number above to filter and see only transactions related to that specific number.")

    st.markdown('</div>', unsafe_allow_html=True)

    # ── DOWNLOADS ──
    st.markdown("---")
    st.markdown('<div class="sec-wrap"><div class="sec-title">⬇️ Download Reports</div>', unsafe_allow_html=True)
    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        html_data = build_html(s, main, acc, provider, fmt_lbl, uploaded.name, focus_num=focus_num)
        st.download_button(
            "📄 HTML রিপোর্ট ডাউনলোড",
            data=html_data.encode('utf-8'),
            file_name=f"mfs_report_{acc}_{provider}.html",
            mime="text/html",
            use_container_width=True,
        )

    with dl2:
        if DOCX_OK:
            docx_data = build_docx(s, main, acc, provider, fmt_lbl, uploaded.name)
            if docx_data:
                st.download_button(
                    "📝 Word রিপোর্ট ডাউনলোড (.docx)",
                    data=docx_data,
                    file_name=f"mfs_report_{acc}_{provider}.docx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.info("Word export এর জন্য: `pip install python-docx`")

    with dl3:
        csv_cols = [c for c in ['তারিখ','TXN_ID','ধরন','দিক','পরিমাণ','আকার','চ্যানেল','কাউন্টারপার্ট','ব্যালেন্স'] if c in main.columns]
        csv = main[csv_cols].to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            "📊 Excel/CSV ডাউনলোড",
            data=csv.encode('utf-8-sig'),
            file_name=f"mfs_transactions_{acc}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div style="text-align:center;font-size:11px;color:#9baab8;padding:16px 0 8px;border-top:1px solid #e8ecf4;margin-top:10px;">
      Mobile Financial System Analyzer · {pname} · Account: {acc} · Bangladesh
    </div>""", unsafe_allow_html=True)


if __name__ == "__main__":
    main_app()
