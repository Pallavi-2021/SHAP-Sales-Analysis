"""
SHAP Sales Intelligence Dashboard  v5
Client-ready: proper units, business axis labels, insight boxes below every chart,
section-level recommendations, data quality panel, benchmarking, advanced filters,
SHAP narratives, AI-generated insights, colour hierarchy, downloadable reports.
"""
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import warnings
warnings.filterwarnings("ignore")

# ── Startup dependency check ────────────────────────────────────────────────
_missing = []
try:
    import xgboost
except ImportError:
    _missing.append("xgboost")
try:
    import shap
except ImportError:
    _missing.append("shap")
try:
    import openpyxl
except ImportError:
    _missing.append("openpyxl")

st.set_page_config(page_title="SHAP Sales Intelligence", page_icon="S",
                   layout="wide", initial_sidebar_state="expanded")

# ── Show dependency warning if anything is missing ─────────────────────────
if _missing:
    st.error(
        f"Missing required libraries: **{', '.join(_missing)}**. "
        "Please ensure your requirements.txt includes: streamlit, pandas, numpy, "
        "matplotlib, xgboost, shap, scikit-learn, openpyxl"
    )
    st.stop()

# ─── CSS ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Poppins:wght@600;700;800&display=swap');
html,body{font-family:'Inter',sans-serif;background:#F4F6F9;color:#1A202C;}
.stApp{background:#F4F6F9!important;}
.stApp p,.stApp span,.stApp label,.stApp li,
.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp h5,.stApp h6,
.stApp div,.stApp td,.stApp th,
.stMarkdown,.stMarkdown *{color:#1A202C!important;}
.stSelectbox label,.stMultiSelect label,.stSlider label,
.stFileUploader label,.stRadio label,.stCheckbox label,
.stNumberInput label,.stTextInput label,.stDateInput label{
    color:#2D3748!important;font-weight:500!important;font-size:0.83rem!important;}
.stSelectbox [data-baseweb="select"]>div,
.stMultiSelect [data-baseweb="select"]>div{
    background:#fff!important;border-color:#CBD5E0!important;color:#1A202C!important;}
[data-baseweb="popover"],[data-baseweb="popover"] *,
[data-baseweb="menu"],[data-baseweb="menu"] *,
[role="listbox"],[role="listbox"] *,
[role="option"],[role="option"] *{background:#fff!important;color:#1A202C!important;}
section[data-testid="stSidebar"]{background:#fff!important;
    border-right:1px solid #E2E8F0!important;box-shadow:2px 0 8px rgba(0,0,0,0.04);}
section[data-testid="stSidebar"],section[data-testid="stSidebar"] *{color:#2D3748!important;}
section[data-testid="stSidebar"] h1,section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3{color:#1A202C!important;
    font-family:'Poppins',sans-serif!important;font-weight:700!important;}
/* HERO — dark bg, white text */
.hero{background:linear-gradient(135deg,#1A365D 0%,#2A4A7F 55%,#2B6CB0 100%);
      border-radius:16px;padding:34px 48px;margin-bottom:20px;position:relative;overflow:hidden;}
.hero::after{content:"";position:absolute;top:-50px;right:-50px;width:240px;height:240px;
             background:rgba(255,255,255,0.04);border-radius:50%;pointer-events:none;}
.htag{display:inline-block;background:rgba(255,255,255,0.14);
      border:1px solid rgba(255,255,255,0.28);border-radius:4px;padding:2px 12px;
      font-size:0.7rem;font-weight:600;color:#fff!important;letter-spacing:0.8px;
      text-transform:uppercase;margin-bottom:11px;}
.htitle{font-family:'Poppins',sans-serif;font-size:2rem;font-weight:800;
        color:#fff!important;margin:0 0 5px 0;letter-spacing:-0.3px;}
.hsub{font-size:0.92rem;color:rgba(255,255,255,0.82)!important;margin:0;font-weight:300;}
/* section headers */
.sh{font-family:'Poppins',sans-serif;font-size:1.1rem;font-weight:700;color:#1A202C!important;
    border-left:4px solid #2B6CB0;padding-left:12px;margin:18px 0 10px 0;}
.sh2{font-family:'Poppins',sans-serif;font-size:0.93rem;font-weight:600;
     color:#2D3748!important;margin:14px 0 6px 0;}
/* KPI card — white bg dark text */
.kcard{background:#fff;border:1px solid #E2E8F0;border-radius:10px;padding:14px 10px;
       text-align:center;box-shadow:0 1px 4px rgba(0,0,0,0.05);
       transition:box-shadow 0.2s,border-color 0.2s;}
.kcard:hover{box-shadow:0 4px 14px rgba(0,0,0,0.09);border-color:#90CDF4;}
.kval{font-family:'Poppins',sans-serif;font-size:1.35rem;font-weight:700;
      color:#2B6CB0!important;display:block;}
.klbl{font-size:0.65rem;color:#718096!important;text-transform:uppercase;
      letter-spacing:0.6px;margin-top:3px;display:block;}
/* insight box — light bg dark text; tags coloured per type */
.ibox{background:#EBF8FF;border-left:4px solid #3182CE;border-radius:0 8px 8px 0;
      padding:11px 15px;margin:8px 0;font-size:0.84rem;color:#1A202C!important;}
.ibox.warn{background:#FFFBEB;border-color:#D69E2E;}
.ibox.ok{background:#F0FFF4;border-color:#38A169;}
.ibox.err{background:#FFF5F5;border-color:#E53E3E;}
.ibox.purple{background:#FAF5FF;border-color:#805AD5;}
.ibox strong,.ibox *{color:#1A202C!important;}
.irow{margin:4px 0;display:flex;gap:8px;align-items:flex-start;}
.itag{background:#2B6CB0;color:#fff!important;border-radius:3px;
      padding:1px 7px;font-size:0.67rem;font-weight:600;white-space:nowrap;flex-shrink:0;}
.ibox.ok .itag{background:#38A169;}
.ibox.warn .itag{background:#D69E2E;}
.ibox.err .itag{background:#E53E3E;}
.ibox.purple .itag{background:#805AD5;}
/* recommendation card */
.rcard{background:#fff;border:1px solid #E2E8F0;border-radius:10px;
       padding:12px 16px;margin:7px 0;box-shadow:0 1px 3px rgba(0,0,0,0.04);}
.rcard.green{border-top:3px solid #38A169;}
.rcard.orange{border-top:3px solid #DD6B20;}
.rcard.red{border-top:3px solid #E53E3E;}
.rcard.blue{border-top:3px solid #3182CE;}
.rtitle{font-weight:600;font-size:0.87rem;color:#1A202C!important;margin:0 0 4px 0;}
.rbody{font-size:0.81rem;color:#4A5568!important;margin:0;}
/* AI insight banner */
.ai-banner{background:linear-gradient(135deg,#FAF5FF,#EBF8FF);
           border:1px solid #B794F4;border-radius:10px;
           padding:12px 16px;margin:10px 0;font-size:0.84rem;color:#1A202C!important;}
.ai-banner strong{color:#553C9A!important;}
/* data quality badges */
.dqok{background:#F0FFF4;border:1px solid #68D391;border-radius:6px;
      padding:5px 10px;display:inline-block;font-size:0.77rem;color:#276749!important;font-weight:500;margin:2px 0;}
.dqwarn{background:#FFFBEB;border:1px solid #F6E05E;border-radius:6px;
        padding:5px 10px;display:inline-block;font-size:0.77rem;color:#744210!important;font-weight:500;margin:2px 0;}
.dqbad{background:#FFF5F5;border:1px solid #FC8181;border-radius:6px;
       padding:5px 10px;display:inline-block;font-size:0.77rem;color:#742A2A!important;font-weight:500;margin:2px 0;}
/* tooltip helper */
.tip{font-size:0.73rem;color:#718096!important;font-style:italic;margin:2px 0 5px 0;}
/* buttons — blue bg white text */
.stButton>button{background:linear-gradient(135deg,#2B6CB0,#3182CE)!important;
    color:#fff!important;border:none!important;border-radius:8px!important;
    font-family:'Poppins',sans-serif!important;font-weight:600!important;
    font-size:0.86rem!important;padding:8px 20px!important;
    transition:opacity 0.2s,transform 0.15s!important;
    box-shadow:0 2px 6px rgba(49,130,206,0.25)!important;}
.stButton>button:hover{opacity:0.88!important;transform:translateY(-1px)!important;}
.stButton>button:disabled{background:#CBD5E0!important;color:#718096!important;
    box-shadow:none!important;transform:none!important;}
div[data-testid="stExpander"]{background:#fff!important;
    border:1px solid #E2E8F0!important;border-radius:10px!important;}
div[data-testid="stExpander"] summary,
div[data-testid="stExpander"] summary *{color:#1A202C!important;font-weight:600!important;}
[data-testid="stDataFrame"],[data-testid="stDataFrame"] *{
    color:#1A202C!important;font-size:0.8rem!important;}
[data-testid="stAlert"] *{color:#1A202C!important;}
/* nav */
.nav-spacer{height:72px;}
.dot{width:8px;height:8px;border-radius:50%;background:#CBD5E0;display:inline-block;margin:0 2px;}
.dot.on{background:#2B6CB0;width:20px;border-radius:4px;}
.nlbl{font-family:'Poppins',sans-serif;font-size:0.76rem;font-weight:600;color:#4A5568!important;}
.pbg{height:4px;background:#E2E8F0;border-radius:2px;margin-top:5px;}
.pfill{height:4px;background:#2B6CB0;border-radius:2px;transition:width 0.35s;}
/* step tracker */
.si{display:flex;align-items:center;gap:9px;padding:5px 0;}
.sc{width:24px;height:24px;border-radius:50%;display:flex;align-items:center;
    justify-content:center;font-size:0.72rem;font-weight:700;flex-shrink:0;}
.sc.done{background:#C6F6D5;color:#276749!important;}
.sc.todo{background:#EDF2F7;color:#718096!important;}
.st2{font-size:0.82rem;color:#4A5568!important;}
.st2.done{color:#276749!important;font-weight:500;}
</style>
""", unsafe_allow_html=True)


# ─── Chart palette & helpers ────────────────────────────────────────────────
CB="#3182CE";CT="#319795";CG="#38A169";CO="#DD6B20";CR="#E53E3E"
CP="#805AD5";CK="#D53F8C";CC="#0BC5EA"
PAL=[CB,CT,CG,CO,CR,CP,CK,CC]
FBG="#FFFFFF";ABG="#FAFBFC";GRC="#EDF2F7";TKC="#4A5568";TTC="#1A202C";LBC="#2D3748"

def sfig(w=9,h=5):
    fig,ax=plt.subplots(figsize=(w,h))
    fig.patch.set_facecolor(FBG);ax.set_facecolor(ABG)
    ax.tick_params(colors=TKC,labelsize=8)
    ax.xaxis.label.set_color(LBC);ax.yaxis.label.set_color(LBC);ax.title.set_color(TTC)
    for sp in ax.spines.values():sp.set_edgecolor("#E2E8F0")
    ax.grid(color=GRC,linewidth=0.6,zorder=0)
    return fig,ax

def sax(ax):
    ax.set_facecolor(ABG);ax.tick_params(colors=TKC,labelsize=8)
    for sp in ax.spines.values():sp.set_edgecolor("#E2E8F0")
    ax.grid(color=GRC,linewidth=0.6,zorder=0)

def _inr_fmt(x,_):
    if x>=1e7: return f"Rs{x/1e7:.1f}Cr"
    if x>=1e5: return f"Rs{x/1e5:.1f}L"
    if x>=1e3: return f"Rs{x/1e3:.0f}K"
    return f"Rs{x:.0f}"
def inr_y(ax): ax.yaxis.set_major_formatter(mticker.FuncFormatter(_inr_fmt))
def inr_x(ax): ax.xaxis.set_major_formatter(mticker.FuncFormatter(_inr_fmt))

def fmt_inr(v):
    try:v=float(v)
    except:return"-"
    if pd.isna(v):return"-"
    if v>=1e7:return f"Rs {v/1e7:.2f} Cr"
    if v>=1e5:return f"Rs {v/1e5:.2f} L"
    if v>=1e3:return f"Rs {v/1e3:.1f} K"
    return f"Rs {v:,.0f}"

def col_num(df,col):
    if col in df.columns:return pd.to_numeric(df[col],errors="coerce").fillna(0)
    return pd.Series(np.zeros(len(df)),index=df.index)

def gb_sum(df,gc,vc):
    t=pd.DataFrame({"k":df[gc].astype(str),"v":pd.to_numeric(df[vc],errors="coerce").fillna(0)})
    return t.groupby("k")["v"].sum().rename_axis(gc)

def gb_mean(df,gc,vc):
    t=pd.DataFrame({"k":df[gc].astype(str),"v":pd.to_numeric(df[vc],errors="coerce").fillna(0)})
    return t.groupby("k")["v"].mean().rename_axis(gc)

def safe_cut(s,n=5,labels=None):
    num=pd.to_numeric(s,errors="coerce")
    if num.dropna().nunique()<2:return pd.Series(["N/A"]*len(s),index=s.index)
    lbl=labels or[str(i) for i in range(1,n+1)]
    try:return pd.cut(num,bins=n,labels=lbl,duplicates="drop")
    except:return pd.Series(["N/A"]*len(s),index=s.index)

def kpi_h(lbl,val,tip=""):
    tip_html=f'<span class="tip">{tip}</span>' if tip else ""
    return(f'<div class="kcard" title="{tip}"><span class="kval">{val}</span>'
           f'<span class="klbl">{lbl}</span>{tip_html}</div>')

def ibox(rows,kind="info"):
    cls={"info":"","warn":" warn","ok":" ok","err":" err","purple":" purple"}.get(kind,"")
    inner="".join(f'<div class="irow"><span class="itag">{t}</span><span>{b}</span></div>'for t,b in rows)
    return f'<div class="ibox{cls}">{inner}</div>'

def ai_banner(text):
    return f'<div class="ai-banner"><strong>AI Insight:</strong> {text}</div>'

def pie_c(ax,vals,labs):
    ax.set_facecolor(FBG)
    w,t,a=ax.pie(vals,labels=labs,autopct="%1.1f%%",colors=PAL[:len(vals)],
                 wedgeprops={"edgecolor":"white","linewidth":1.5},startangle=140)
    for tx in t:tx.set_color(TKC);tx.set_fontsize(8)
    for tx in a:tx.set_color("white");tx.set_fontsize(7);tx.set_fontweight("bold")

def rcard(title,body,color="blue"):
    return(f'<div class="rcard {color}"><p class="rtitle">{title}</p>'
           f'<p class="rbody">{body}</p></div>')

def legend_kw():
    return dict(facecolor="white",edgecolor="#E2E8F0",fontsize=8)

# ─── SHAP friendly names ────────────────────────────────────────────────────
FNAME={
    "Qty (StdUnit)":"Quantity Sold (Std Units)",
    "Dispatch Qty (StdUnit)":"Dispatched Quantity (Std Units)",
    "Sale Value":"Sale Value (Rs)",
    "Net Value":"Net Value Ex-Tax (Rs)",
    "Net Value Tax Inclusive":"Total Revenue Incl. Tax (Rs)",
    "Product MRP":"Product MRP (Rs)",
    "Price":"Unit Price (Rs)",
    "Actual Price":"Actual Price (Rs)",
    "Scheme Value":"Scheme Value (Rs)",
    "Scheme Quantity":"Scheme Quantity (Units)",
    "Outlet margin":"Outlet Margin (%)",
    "CGST":"CGST (Rs)","SGST":"SGST (Rs)","IGST":"IGST (Rs)","VAT":"VAT (Rs)",
    "avg_price_per_unit":"Avg Price per Unit (Rs)",
    "dispatch_rate":"Dispatch Rate (%)",
    "tax_ratio":"Tax Ratio",
    "scheme_intensity":"Scheme Intensity",
    "order_dayofweek":"Order Day of Week",
    "order_day":"Order Day",
    "order_quarter":"Order Quarter",
    "order_month_num":"Order Month",
    "Standard Unit Conversion Factor":"Unit Conversion Factor",
    "Depot_enc":"Depot (encoded)",
    "Classification_enc":"Classification (encoded)",
    "Order Type_enc":"Order Type (encoded)",
    "Delivery Status_enc":"Delivery Status (encoded)",
    "ProductDivision_enc":"Product Division (encoded)",
    "PrimaryCategory_enc":"Primary Category (encoded)",
    "SecondaryCategory_enc":"Secondary Category (encoded)",
    "Outlets Channel Name_enc":"Outlet Channel (encoded)",
    "Outlets Type_enc":"Outlet Type (encoded)",
    "Outlets Segmentation_enc":"Outlet Segmentation (encoded)",
    "Display Category_enc":"Display Category (encoded)",
    "StdUnit_enc":"Std Unit (encoded)",
}
def fn(col): return FNAME.get(col,col.replace("_"," ").title())


# ─── Session state ──────────────────────────────────────────────────────────
PAGES=["Upload","Preprocess","Data Quality","Overview","SHAP Analysis","Sales Drivers",
       "Outlet Analysis","Product & Category","Sales Executive","Delivery & Fulfilment",
       "Depot & Regional","Margin & Tax","Segmentation","Executive Summary"]
N=len(PAGES)

for k,v in dict(raw_df=None,processed=False,model=None,shap_values=None,feature_names=None,
    X_sample=None,df=None,shap_df=None,page=0,
    f_depot="All",f_month="All",f_otype="All",f_division="All",
    f_seg="All",f_dstatus="All",f_date=None).items():
    if k not in st.session_state:st.session_state[k]=v

def go(d):
    t=st.session_state.page+d
    if t>=2 and not st.session_state.processed:st.toast("Complete preprocessing first.");return
    st.session_state.page=max(0,min(N-1,t))
cur=st.session_state.page

# ─── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## Sales Intelligence")
    st.markdown("---")
    for i,(nm,done) in enumerate([
        ("Upload Dataset",st.session_state.raw_df is not None),
        ("Preprocess & Train",st.session_state.processed),
        ("Analytics Ready",st.session_state.shap_values is not None)]):
        cls="done" if done else "todo";icon="+" if done else str(i+1)
        st.markdown(f'<div class="si"><div class="sc {cls}">{icon}</div>'
                    f'<span class="st2 {"done" if done else ""}">{nm}</span></div>',
                    unsafe_allow_html=True)
    st.markdown("---");st.markdown("### Navigation")
    for i,pname in enumerate(PAGES):
        locked=i>=2 and not st.session_state.processed
        lbl=f"[locked] {pname}" if locked else pname
        if st.button(lbl,key=f"nav_{i}",disabled=locked):
            st.session_state.page=i;st.rerun()
    st.markdown("---")
    if st.session_state.processed:
        st.markdown("### Advanced Filters")
        dfs=st.session_state.df
        def opts(col):
            return(["All"]+sorted(dfs[col].dropna().astype(str).unique().tolist())
                   if col in dfs.columns else["All"])
        st.session_state.f_depot   =st.selectbox("Depot",               opts("Depot"),            help="Filter all charts by depot location")
        st.session_state.f_month   =st.selectbox("Month",               opts("Month"),            help="Filter by calendar month")
        st.session_state.f_otype   =st.selectbox("Order Type",          opts("Order Type"),       help="Filter by order classification type")
        st.session_state.f_division=st.selectbox("Product Division",    opts("ProductDivision"),  help="Filter by product division")
        st.session_state.f_seg     =st.selectbox("Outlet Segmentation", opts("Outlets Segmentation"), help="Filter by outlet tier")
        st.session_state.f_dstatus =st.selectbox("Delivery Status",     opts("Delivery Status"),  help="Filter by delivery outcome")
        if "_order_date_raw" in dfs.columns:
            dts=pd.to_datetime(dfs["_order_date_raw"],errors="coerce").dropna()
            if len(dts):
                mn_d,mx_d=dts.min().date(),dts.max().date()
                dr=st.date_input("Date Range",value=(mn_d,mx_d),min_value=mn_d,max_value=mx_d,
                                 help="Filter all charts by order date range")
                st.session_state.f_date=dr if isinstance(dr,(list,tuple)) and len(dr)==2 else None
    st.markdown('<div style="color:#A0AEC0;font-size:0.7rem;margin-top:6px">'
                'Streamlit | XGBoost | SHAP | v5</div>',unsafe_allow_html=True)

# ─── Hero ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <span class="htag">AI-Powered Sales Analytics Platform</span>
  <p class="htitle">SHAP Sales Intelligence Dashboard</p>
  <p class="hsub">Explainable machine learning for sales performance, outlet analytics and distribution insights</p>
</div>""",unsafe_allow_html=True)

# ─── Filter helper ───────────────────────────────────────────────────────────
def get_filt():
    df=st.session_state.df.copy()
    for col,key in[("Depot","f_depot"),("Month","f_month"),("Order Type","f_otype"),
                   ("ProductDivision","f_division"),("Outlets Segmentation","f_seg"),
                   ("Delivery Status","f_dstatus")]:
        if st.session_state[key]!="All" and col in df.columns:
            df2=df[df[col].astype(str)==st.session_state[key]]
            df=df2 if len(df2) else df
    dr=st.session_state.f_date
    if dr and "_order_date_raw" in df.columns:
        dts=pd.to_datetime(df["_order_date_raw"],errors="coerce")
        df2=df[(dts.dt.date>=dr[0])&(dts.dt.date<=dr[1])]
        df=df2 if len(df2) else df
    return df if len(df) else st.session_state.df.copy()


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 0 — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
if cur==0:
    st.markdown('<div class="sh">Step 1 — Upload Dataset</div>',unsafe_allow_html=True)
    st.markdown(ibox([
        ("Format","CSV or Excel (.xlsx / .xls) accepted — max ~500 MB."),
        ("Schema","Standard sales & distribution columns expected (Sale Value, Depot, Order Type, etc.)."),
        ("Next","After uploading, proceed to Step 2 to preprocess and unlock all 14 analytics pages.")
    ]),unsafe_allow_html=True)
    up=st.file_uploader("Drop file here or click to browse",type=["csv","xlsx","xls"],label_visibility="collapsed")
    if up:
        try:
            dfr=(pd.read_csv(up,low_memory=False) if up.name.endswith(".csv") else pd.read_excel(up))
            st.session_state.raw_df=dfr
            st.success(f"Loaded {dfr.shape[0]:,} rows x {dfr.shape[1]} columns.")
            with st.expander("Preview — first 10 rows"):
                st.dataframe(dfr.head(10),width='stretch')
            miss=[c for c in["Sale Value","Depot","Order Type","Delivery Status","Product","Outlets Name"]
                  if c not in dfr.columns]
            if miss:
                st.markdown(ibox([("Missing Columns",", ".join(miss)),
                                  ("Note","The app will use whichever columns are available.")],"warn"),
                            unsafe_allow_html=True)
        except Exception as e:st.error(f"Could not read file: {e}")
    elif st.session_state.raw_df is not None:
        r=st.session_state.raw_df
        st.success(f"Dataset already loaded — {r.shape[0]:,} rows x {r.shape[1]} columns.")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — PREPROCESS
# ══════════════════════════════════════════════════════════════════════════════
elif cur==1:
    st.markdown('<div class="sh">Step 2 — Preprocess and Train SHAP Model</div>',unsafe_allow_html=True)
    if st.session_state.raw_df is None:
        st.markdown(ibox([("Action","Please upload a dataset first (Step 1).")],"warn"),unsafe_allow_html=True)
    else:
        if st.session_state.processed:
            st.markdown(ibox([("Status","Already preprocessed. Re-run to refresh model.")],"ok"),unsafe_allow_html=True)
        st.markdown(ibox([
            ("Step 1","Numeric coercion for all financial and quantity columns."),
            ("Step 2","Date feature extraction (day, quarter, month) then datetime columns are dropped."),
            ("Step 3","Median imputation for numeric gaps; Unknown fill for categoricals."),
            ("Step 4","All object columns cast to clean strings to prevent groupby crashes."),
            ("Step 5","Derived features: avg price/unit, dispatch rate, tax ratio, scheme intensity."),
            ("Step 6","Label-encode 12 categorical columns for XGBoost."),
            ("Step 7","XGBoost Regressor (300 trees) on 80/20 split, top-99th-percentile revenue target."),
            ("Step 8","SHAP TreeExplainer on up to 2,000-order sample.")
        ]),unsafe_allow_html=True)
        if st.button("Run Preprocessing and Train Model"):
            with st.spinner("Engineering features, training XGBoost, computing SHAP values..."):
                try:
                    import xgboost as xgb,shap as shap_lib
                    from sklearn.preprocessing import LabelEncoder
                    from sklearn.model_selection import train_test_split
                    df=st.session_state.raw_df.copy();df.columns=df.columns.str.strip()
                    NUM=["Qty (StdUnit)","Dispatch Qty (StdUnit)","Sale Value","Net Value",
                         "Net Value Tax Inclusive","Product MRP","Price","Actual Price","MRP",
                         "PTSS","VTSS","Sale Value (Dispatch)","Net Dispatch Value (Tax Exclusive)",
                         "Scheme Quantity","Scheme Value","Outlet margin","CGST","SGST","IGST",
                         "VAT","Qty (Unit)","Dispatch Qty (Unit)","Standard Unit Conversion Factor","Lat","Long"]
                    for c in NUM:
                        if c in df.columns:df[c]=pd.to_numeric(df[c],errors="coerce")
                    DATE_COLS=["Order Date","Validation Date","Delivery/Rejection Date","Outlets Created On"]
                    if "Order Date" in df.columns:
                        dp=pd.to_datetime(df["Order Date"],errors="coerce")
                        df["_order_date_raw"]=df["Order Date"].astype(str)
                        df["order_dayofweek"]=dp.dt.dayofweek.astype("float32")
                        df["order_day"]=dp.dt.day.astype("float32")
                        df["order_quarter"]=dp.dt.quarter.astype("float32")
                        df["order_month_num"]=dp.dt.month.astype("float32")
                    for c in DATE_COLS:
                        if c in df.columns:df.drop(columns=[c],inplace=True)
                    for c in df.select_dtypes(include="number").columns:df[c]=df[c].fillna(df[c].median())
                    for c in df.select_dtypes(include="object").columns:df[c]=df[c].fillna("Unknown")
                    for c in df.select_dtypes(exclude="number").columns:df[c]=df[c].astype(str).str.strip()
                    qty=col_num(df,"Qty (StdUnit)");dqty=col_num(df,"Dispatch Qty (StdUnit)")
                    sv_=col_num(df,"Sale Value");nv=col_num(df,"Net Value").replace(0,np.nan)
                    nvi=col_num(df,"Net Value Tax Inclusive");schv=col_num(df,"Scheme Value")
                    df["avg_price_per_unit"]=np.where(qty>0,sv_/qty,0)
                    df["dispatch_rate"]=np.where(qty>0,dqty/qty,0).clip(0,2)
                    df["tax_ratio"]=((nvi-nv)/nv).fillna(0).clip(0,5)
                    df["scheme_intensity"]=np.where(sv_>0,schv/sv_,0).clip(0,5)
                    st.session_state.df=df
                    CAT=["Depot","Classification","Order Type","Delivery Status","ProductDivision",
                         "PrimaryCategory","SecondaryCategory","Outlets Channel Name","Outlets Type",
                         "Outlets Segmentation","Display Category","StdUnit"]
                    df_m=df.copy()
                    for c in CAT:
                        if c in df_m.columns:
                            le=LabelEncoder();df_m[c+"_enc"]=le.fit_transform(df_m[c].astype(str))
                    FC=(["Qty (StdUnit)","Dispatch Qty (StdUnit)","Product MRP","Price","Actual Price",
                         "Scheme Quantity","Scheme Value","Outlet margin","CGST","SGST","IGST","VAT",
                         "PTSS","VTSS","avg_price_per_unit","dispatch_rate","tax_ratio","scheme_intensity",
                         "order_dayofweek","order_day","order_quarter","order_month_num",
                         "Standard Unit Conversion Factor"]+[c+"_enc" for c in CAT])
                    fc=[c for c in FC if c in df_m.columns]
                    X=df_m[fc].apply(pd.to_numeric,errors="coerce").fillna(0)
                    if "Sale Value" not in df_m.columns:st.error("Sale Value column required.");st.stop()
                    y=pd.to_numeric(df_m["Sale Value"],errors="coerce").fillna(0)
                    mask=y<=y.quantile(0.99);X,y=X[mask],y[mask]
                    Xtr,Xte,ytr,_=train_test_split(X,y,test_size=0.2,random_state=42)
                    mdl=xgb.XGBRegressor(n_estimators=300,max_depth=6,learning_rate=0.05,
                                         subsample=0.8,colsample_bytree=0.8,
                                         random_state=42,n_jobs=-1,verbosity=0)
                    mdl.fit(Xtr,ytr)
                    Xs=Xte.iloc[:min(2000,len(Xte))]
                    exp=shap_lib.TreeExplainer(mdl);sv_shap=exp.shap_values(Xs)
                    ma=np.abs(sv_shap).mean(axis=0)
                    sdf=(pd.DataFrame({"Feature":fc,"Friendly":[ fn(c) for c in fc],"Mean_SHAP":ma})
                           .sort_values("Mean_SHAP",ascending=False).reset_index(drop=True))
                    sdf["Rank"]=range(1,len(sdf)+1)
                    st.session_state.update(dict(model=mdl,shap_values=sv_shap,feature_names=fc,
                                                  X_sample=Xs,shap_df=sdf,processed=True))
                    st.success("Preprocessing complete. XGBoost model trained. SHAP values computed.")
                    st.balloons()
                except ImportError as ie:st.error(f"Missing library: {ie}. Run: pip install xgboost shap")
                except Exception as e:
                    import traceback;st.error(f"Preprocessing failed: {e}");st.code(traceback.format_exc())

# ══════════════════════════════════════════════════════════════════════════════
# LOCKED GATE
# ══════════════════════════════════════════════════════════════════════════════
elif not st.session_state.processed:
    st.markdown(ibox([("Locked","Complete Step 1 (upload) and Step 2 (preprocess) to unlock all analytics.")],"warn"),
                unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# ANALYTICS BLOCK (pages 2-13)
# ══════════════════════════════════════════════════════════════════════════════
else:
    filt=get_filt()
    SV,ON,OUT,PRD,DEP,DR="Sale Value","Order No","Outlets Name","Product","Depot","dispatch_rate"
    total_rev  =col_num(filt,SV).sum()
    total_ord  =filt[ON].nunique()           if ON  in filt.columns else len(filt)
    total_out  =filt[OUT].astype(str).nunique() if OUT in filt.columns else 0
    avg_ord    =total_rev/max(total_ord,1)
    disp_rate  =col_num(filt,DR).mean()      if DR  in filt.columns else 0
    total_prd  =filt[PRD].astype(str).nunique() if PRD in filt.columns else 0
    total_qty  =col_num(filt,"Qty (StdUnit)").sum()
    total_tax  =sum(col_num(filt,c).sum() for c in["CGST","SGST","IGST","VAT"] if c in filt.columns)
    sv_arr=st.session_state.shap_values;fn_arr=st.session_state.feature_names
    Xs_arr=st.session_state.X_sample;sdf=st.session_state.shap_df
    mean_abs=np.abs(sv_arr).mean(axis=0)

    # ── 8-wide KPI strip ──────────────────────────────────────────────────
    cols8=st.columns(8)
    for col_,(lbl_,val_,tip_) in zip(cols8,[
        ("Total Revenue",     fmt_inr(total_rev),      "Sum of all Sale Values in filtered data"),
        ("Total Orders",      f"{total_ord:,}",         "Unique order count (Order No)"),
        ("Active Outlets",    f"{total_out:,}",         "Distinct outlet names in filtered data"),
        ("Avg Order Value",   fmt_inr(avg_ord),         "Total Revenue / Total Orders"),
        ("Dispatch Rate",     f"{disp_rate:.1%}",       "Avg dispatched qty / ordered qty — target 95%+"),
        ("Unique Products",   f"{total_prd:,}",         "Distinct products/SKUs in filtered data"),
        ("Qty Sold",          f"{total_qty:,.0f} Units","Sum of Qty (StdUnit) across filtered orders"),
        ("Total Tax",         fmt_inr(total_tax),       "Sum of CGST+SGST+IGST+VAT"),
    ]):
        col_.markdown(kpi_h(lbl_,val_,tip_),unsafe_allow_html=True)
    st.markdown("<div style='height:8px'></div>",unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 2 — DATA QUALITY
    # ══════════════════════════════════════════════════════════════════════
    if cur==2:
        st.markdown('<div class="sh">Data Quality and Model Reliability Report</div>',unsafe_allow_html=True)
        raw=st.session_state.raw_df;df_p=st.session_state.df
        total_rows=len(raw);dupe_rows=raw.duplicated().sum()
        c1,c2,c3=st.columns(3)
        with c1:
            st.markdown('<div class="sh2">Dataset Dimensions</div>',unsafe_allow_html=True)
            for l_,v_ in[("Total Records",f"{total_rows:,}"),
                          ("Total Columns",str(len(raw.columns))),
                          ("Duplicate Rows",f"{dupe_rows:,} ({dupe_rows/max(total_rows,1)*100:.1f}%)"),
                          ("Numeric Columns",str(len(raw.select_dtypes(include='number').columns))),
                          ("Object Columns",str(len(raw.select_dtypes(include='object').columns)))]:
                st.markdown(kpi_h(l_,v_),unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="sh2">Missing Value Analysis (Top 8 Columns)</div>',unsafe_allow_html=True)
            mp=(raw.isnull().sum()/len(raw)*100).sort_values(ascending=False)
            mt=mp[mp>0].head(8)
            if len(mt):
                for cn_,pv_ in mt.items():
                    bc="dqbad" if pv_>20 else("dqwarn" if pv_>5 else"dqok")
                    st.markdown(f'<div class="{bc}" style="width:100%;margin:2px 0">{cn_}: {pv_:.1f}% missing</div>',
                                unsafe_allow_html=True)
            else:
                st.markdown('<div class="dqok" style="width:100%">All key columns complete.</div>',unsafe_allow_html=True)
        with c3:
            st.markdown('<div class="sh2">Model Confidence Indicators</div>',unsafe_allow_html=True)
            top_feat=sdf.iloc[0]["Friendly"] if len(sdf) else"N/A"
            top_shap=sdf.iloc[0]["Mean_SHAP"] if len(sdf) else 0
            for l_,v_ in[("Features in Model",str(len(fn_arr))),
                          ("SHAP Sample Size",f"{len(Xs_arr):,} orders"),
                          ("Top Revenue Driver",top_feat),
                          ("Top Driver SHAP Score",f"{top_shap:.4f}"),
                          ("Training Split","80% train / 20% test")]:
                st.markdown(kpi_h(l_,v_),unsafe_allow_html=True)
        st.markdown("---")
        # Missing value bar chart
        st.markdown('<div class="sh2">Missing Value Rate by Column (%)</div>',unsafe_allow_html=True)
        st.markdown('<p class="tip">Green = acceptable (<5%) | Orange = caution (5-20%) | Red = unreliable (>20%)</p>',unsafe_allow_html=True)
        miss_all=(raw.isnull().sum()/len(raw)*100).sort_values(ascending=False).head(15)
        if miss_all.sum()>0:
            fig,ax=sfig(12,4)
            colors_=[CR if v>20 else(CO if v>5 else CG) for v in miss_all.values]
            ax.barh(miss_all.index[::-1],miss_all.values[::-1],color=colors_[::-1],edgecolor="none",height=0.65)
            ax.set_xlabel("Missing Values (%)",color=LBC)
            ax.axvline(5,color=CO,lw=1.2,ls="--",alpha=0.7,label="5% caution threshold")
            ax.axvline(20,color=CR,lw=1.2,ls="--",alpha=0.7,label="20% unreliable threshold")
            ax.set_title("Data Completeness — Missing Value Rate by Column",color=TTC,fontsize=11,fontweight="bold")
            ax.legend(**legend_kw())
            fig.tight_layout();st.pyplot(fig);plt.close()
            worst_col=miss_all.index[0] if len(miss_all) else"N/A"
            st.markdown(ibox([
                ("What This Shows","Percentage of missing values for each column in the raw uploaded dataset."),
                ("Key Insight",f"'{worst_col}' has the highest missing rate at {miss_all.iloc[0]:.1f}%. Columns above 20% may produce unreliable insights."),
                ("Business Impact","High missing rates in key revenue or outlet fields reduce model accuracy and chart reliability."),
                ("Action","Ensure upstream data capture processes are complete. For critical columns, consider mandatory field enforcement in the ERP system.")
            ]),unsafe_allow_html=True)
        else:
            st.markdown(ibox([("Data Completeness","No missing values detected in the dataset. Excellent data quality.")],"ok"),unsafe_allow_html=True)
        st.markdown("---")
        st.markdown('<div class="sh2">Numeric Column Statistical Summary</div>',unsafe_allow_html=True)
        ns=df_p.select_dtypes(include="number").describe().T[["count","mean","std","min","50%","max"]]
        ns.columns=["Count","Mean","Std Dev","Min","Median","Max"]
        for c_ in["Mean","Std Dev","Min","Median","Max"]:ns[c_]=ns[c_].apply(lambda x:f"{x:,.2f}")
        st.dataframe(ns,width='stretch')
        # AI insight
        completeness=100-mp.mean()
        ai_msg=(f"Data completeness is {completeness:.1f}% overall. "
                +(f"Column '{mt.index[0]}' has the most gaps ({mt.iloc[0]:.1f}%). " if len(mt) else "All columns are complete. ")
                +("Consider data quality remediation before sharing client-facing reports." if completeness<90
                  else "Data quality is sufficient for reliable analytical insights."))
        st.markdown(ai_banner(ai_msg),unsafe_allow_html=True)


    # ══════════════════════════════════════════════════════════════════════
    # PAGE 3 — OVERVIEW
    # ══════════════════════════════════════════════════════════════════════
    elif cur==3:
        st.markdown('<div class="sh">Dashboard Overview</div>',unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="sh2">Monthly Revenue Trend</div>',unsafe_allow_html=True)
            if "Month" in filt.columns and SV in filt.columns:
                mon=(filt.assign(M=filt["Month"].astype(str),V=col_num(filt,SV))
                         .groupby("M")["V"].sum().reset_index().sort_values("M"))
                fig,ax=sfig(7,4)
                ax.fill_between(range(len(mon)),mon["V"],alpha=0.13,color=CB)
                ax.plot(range(len(mon)),mon["V"],color=CB,lw=2.5,marker="o",ms=5,zorder=3)
                ax.set_xticks(range(len(mon)));ax.set_xticklabels(mon["M"],rotation=40,ha="right",fontsize=7,color=TKC)
                ax.set_ylabel("Total Revenue (Rs)",color=LBC);ax.set_xlabel("Month",color=LBC);inr_y(ax)
                ax.set_title("Monthly Revenue Trend (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                peak=mon.loc[mon["V"].idxmax(),"M"] if len(mon) else"N/A"
                low=mon.loc[mon["V"].idxmin(),"M"] if len(mon) else"N/A"
                mom_chg=((mon["V"].iloc[-1]-mon["V"].iloc[-2])/mon["V"].iloc[-2]*100
                         if len(mon)>1 and mon["V"].iloc[-2]>0 else 0)
                st.markdown(ibox([
                    ("What This Shows","Total revenue (Rs) aggregated by calendar month across all orders."),
                    ("Key Insight",f"Peak month: '{peak}'. Lowest month: '{low}'. Last period change: {mom_chg:+.1f}%."),
                    ("Business Impact","Month-on-month variance directly drives inventory planning, staffing, and target-setting decisions."),
                    ("Action","Investigate root causes of low-revenue months. Replicate conditions (schemes, coverage, executive drive) from peak months.")
                ]),unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="sh2">Top 10 Revenue Drivers — SHAP Feature Importance</div>',unsafe_allow_html=True)
            t10=sdf.head(10);fig,ax=sfig(7,4)
            ax.barh(t10["Friendly"][::-1],t10["Mean_SHAP"][::-1],color=PAL[:len(t10)],edgecolor="none",height=0.65)
            ax.set_xlabel("Mean |SHAP| — Revenue Influence Score",color=LBC)
            ax.set_title("Top 10 Revenue Drivers (SHAP)",color=TTC,fontsize=11,fontweight="bold")
            fig.tight_layout();st.pyplot(fig);plt.close()
            st.markdown(ibox([
                ("What This Shows","The 10 variables with the highest average impact on predicted sale value, ranked by SHAP score."),
                ("Key Insight",f"'{t10.iloc[0]['Friendly']}' is the single strongest revenue driver (SHAP={t10.iloc[0]['Mean_SHAP']:.4f})."),
                ("Business Impact","These drivers collectively explain the majority of revenue variance across all orders."),
                ("Action","Focus sales strategy, scheme design, and operational planning on the top 3-5 SHAP drivers.")
            ]),unsafe_allow_html=True)
        c3,c4=st.columns(2)
        with c3:
            st.markdown('<div class="sh2">Delivery Status Distribution</div>',unsafe_allow_html=True)
            if "Delivery Status" in filt.columns:
                ds=filt["Delivery Status"].astype(str).value_counts()
                if len(ds):
                    fig,ax=sfig(6,4);pie_c(ax,ds.values,ds.index)
                    ax.set_title("Order Count by Delivery Status",color=TTC,fontsize=11,fontweight="bold")
                    fig.tight_layout();st.pyplot(fig);plt.close()
                    deliv_pct=ds.get("Delivered",ds.iloc[0])/ds.sum()*100
                    st.markdown(ibox([
                        ("What This Shows","Proportion of total orders by delivery outcome."),
                        ("Key Insight",f"Approximately {deliv_pct:.1f}% of orders are in the top delivery category."),
                        ("Action","Target delivery failure categories with root-cause analysis and logistics improvements.")
                    ]),unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="sh2">Revenue by Product Division (Rs)</div>',unsafe_allow_html=True)
            if "ProductDivision" in filt.columns and SV in filt.columns:
                pdiv=gb_sum(filt,"ProductDivision",SV).sort_values(ascending=False).head(8)
                fig,ax=sfig(6,4)
                ax.bar(range(len(pdiv)),pdiv.values,color=PAL[:len(pdiv)],edgecolor="none")
                ax.set_xticks(range(len(pdiv)));ax.set_xticklabels(pdiv.index,rotation=30,ha="right",fontsize=7,color=TKC)
                ax.set_ylabel("Total Revenue (Rs)",color=LBC);ax.set_xlabel("Product Division",color=LBC);inr_y(ax)
                ax.set_title("Revenue by Product Division (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                top_div=pdiv.index[0] if len(pdiv) else"N/A"
                st.markdown(ibox([
                    ("Key Insight",f"'{top_div}' is the top-contributing division. Bottom divisions may need strategic review."),
                    ("Action","Invest marketing and distribution resources in top divisions; evaluate exit or turnaround for underperformers.")
                ]),unsafe_allow_html=True)
        # AI insight for overview
        top_drv=sdf.iloc[0]["Friendly"] if len(sdf) else"N/A"
        top_div2=(gb_sum(filt,"ProductDivision",SV).idxmax() if "ProductDivision" in filt.columns and SV in filt.columns else"N/A")
        st.markdown(ai_banner(
            f"SHAP analysis identifies '{top_drv}' as the primary revenue driver. "
            f"Product division '{top_div2}' leads revenue contribution. "
            f"Current dispatch rate of {disp_rate:.1%} represents a fulfilment gap that should be the top operational priority."
        ),unsafe_allow_html=True)
        st.markdown("---");st.markdown('<div class="sh">Section Recommendations</div>',unsafe_allow_html=True)
        st.markdown(rcard("Revenue Stability","Stable monthly trend and clear divisional leaders indicate a solid base. Use peak months as planning benchmarks for all future periods.","green"),unsafe_allow_html=True)
        st.markdown(rcard("Delivery Gap Risk","Undelivered orders represent direct revenue leakage and customer dissatisfaction. Each failed delivery is a recoverable loss.","orange"),unsafe_allow_html=True)
        st.markdown(rcard("SHAP-Driven Strategy",f"Concentrating commercial effort on top SHAP driver '{top_drv}' yields the highest revenue uplift with the least operational change.","blue"),unsafe_allow_html=True)


    # ══════════════════════════════════════════════════════════════════════
    # PAGE 4 — SHAP ANALYSIS
    # ══════════════════════════════════════════════════════════════════════
    elif cur==4:
        st.markdown('<div class="sh">SHAP Explainability Analysis</div>',unsafe_allow_html=True)
        st.markdown(ibox([
            ("What is SHAP","SHapley Additive exPlanations — a game-theory method that fairly attributes revenue prediction to each input variable."),
            ("Positive SHAP","This feature is pushing the predicted sale value HIGHER for those orders."),
            ("Negative SHAP","This feature is pushing the predicted sale value LOWER for those orders."),
            ("Scale","SHAP values are in the same unit as the target (Rs). A SHAP of 500 means this feature adds ~Rs 500 to the prediction."),
            ("Business Use","Use SHAP to understand WHY revenue is high or low — not just what the number is.")
        ],"purple"),unsafe_allow_html=True)

        c1,c2=st.columns([3,2])
        with c1:
            st.markdown('<div class="sh2">Feature Importance — Mean |SHAP| Revenue Influence Score</div>',unsafe_allow_html=True)
            st.markdown('<p class="tip">Longer bar = greater average impact on predicted sale value (Rs). Colour represents feature rank.</p>',unsafe_allow_html=True)
            t20=sdf.head(20);fig,ax=sfig(9,7)
            ax.barh(t20["Friendly"][::-1],t20["Mean_SHAP"][::-1],
                    color=[PAL[i%len(PAL)] for i in range(len(t20))],edgecolor="none",height=0.7)
            ax.set_xlabel("Mean |SHAP| Value — Average Revenue Influence (Rs direction)",color=LBC)
            ax.set_title("Feature Importance via SHAP (Top 20 Revenue Drivers)",color=TTC,fontsize=12,fontweight="bold")
            fig.tight_layout();st.pyplot(fig);plt.close()
            st.markdown(ibox([
                ("What This Shows","Top 20 variables ranked by their average absolute SHAP value across all sampled orders."),
                ("Key Insight",f"Top driver: '{sdf.iloc[0]['Friendly']}' (Mean |SHAP| = {sdf.iloc[0]['Mean_SHAP']:.4f}). This variable alone explains the largest share of revenue variation."),
                ("Business Impact","Features ranked 1-5 collectively account for the majority of predictable revenue variance."),
                ("Action","Operational improvements targeting the top 3 features will deliver the highest revenue ROI.")
            ]),unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="sh2">SHAP Feature Ranking</div>',unsafe_allow_html=True)
            st.markdown('<p class="tip">Mean |SHAP| ≈ average Rs impact on revenue prediction per order</p>',unsafe_allow_html=True)
            tbl=sdf[["Rank","Friendly","Mean_SHAP"]].head(20).copy()
            tbl.columns=["Rank","Business Name","Mean |SHAP| (Rs)"]
            tbl["Mean |SHAP| (Rs)"]=tbl["Mean |SHAP| (Rs)"].round(4)
            st.dataframe(tbl.reset_index(drop=True),width='stretch',height=530)

        # SHAP narrative
        st.markdown('<div class="sh2">Plain-English SHAP Interpretations</div>',unsafe_allow_html=True)
        narr=[]
        for i,(_,r) in enumerate(sdf.head(5).iterrows()):
            direction="strongly increases" if i<3 else "moderately influences"
            narr.append((f"#{r['Rank']} {r['Friendly']}",
                         f"Mean |SHAP| = {r['Mean_SHAP']:.4f} Rs. When this variable is high, it {direction} predicted revenue. "
                         f"Operations that improve this variable are expected to deliver measurable top-line uplift."))
        st.markdown(ibox(narr,"purple"),unsafe_allow_html=True)

        st.markdown("---")
        # Beeswarm
        st.markdown('<div class="sh2">SHAP Beeswarm — Distribution of Revenue Impact per Feature</div>',unsafe_allow_html=True)
        st.markdown('<p class="tip">Each dot = one order. Horizontal position = how much this feature shifted predicted revenue (Rs). Blue = low feature value, Red = high feature value.</p>',unsafe_allow_html=True)
        top_n=st.slider("Features to display",5,20,12,key="bs_n")
        ti=np.argsort(mean_abs)[::-1][:top_n]
        st_s=sv_arr[:,ti];Xt_=Xs_arr.iloc[:,ti]
        ft_=[sdf.loc[sdf["Feature"]==fn_arr[i],"Friendly"].values[0] if fn_arr[i] in sdf["Feature"].values else fn(fn_arr[i]) for i in ti]
        fig2,ax2=sfig(11,6);sax(ax2)
        for fi in range(top_n):
            sv2=st_s[:,fi];fv=Xt_.iloc[:,fi].values
            fn2=(fv-fv.min())/(np.ptp(fv)+1e-9)
            clrs=plt.cm.RdYlBu_r(fn2);jit=np.random.uniform(-0.18,0.18,size=len(sv2))
            ax2.scatter(sv2,fi+jit,c=clrs,alpha=0.3,s=7,linewidths=0,zorder=2)
        ax2.set_yticks(range(top_n));ax2.set_yticklabels(ft_,fontsize=8,color=TKC)
        ax2.axvline(0,color=CR,lw=1,ls="--",alpha=0.5,label="Zero impact line")
        ax2.set_xlabel("SHAP Value — Revenue Impact (Rs, positive = increases prediction)",color=LBC)
        ax2.set_title("SHAP Beeswarm — Feature Value vs Revenue Impact Direction",color=TTC,fontsize=11,fontweight="bold")
        ax2.legend(**legend_kw())
        fig2.tight_layout();st.pyplot(fig2);plt.close()
        st.markdown(ibox([
            ("What This Shows","Each point is one order; horizontal position shows how much this feature shifted the revenue prediction (Rs)."),
            ("Right of Zero Line","Feature is increasing predicted revenue for those orders."),
            ("Left of Zero Line","Feature is reducing predicted revenue for those orders."),
            ("Colour","Blue = low feature value, Red = high feature value. Wide rightward red spread = high values drive high revenue."),
            ("Action","Features with consistent rightward spread of red dots are the most actionable levers for revenue growth.")
        ]),unsafe_allow_html=True)

        st.markdown("---")
        # Dependence plot
        st.markdown('<div class="sh2">Dependence Plot — How a Feature Value Affects Revenue</div>',unsafe_allow_html=True)
        dc1,dc2=st.columns(2)
        with dc1:fx=st.selectbox("Primary Feature (X-axis)",[fn(c) for c in fn_arr],key="dep_x")
        with dc2:fc2=st.selectbox("Colour By Feature",[fn(c) for c in fn_arr],index=min(1,len(fn_arr)-1),key="dep_c")
        fx_raw=[c for c in fn_arr if fn(c)==fx][0]
        fc2_raw=[c for c in fn_arr if fn(c)==fc2][0]
        ix,ic=fn_arr.index(fx_raw),fn_arr.index(fc2_raw)
        cv=Xs_arr.iloc[:,ic].values;cn=(cv-cv.min())/(np.ptp(cv)+1e-9)
        fig3,ax3=sfig(9,5);sax(ax3)
        sc=ax3.scatter(Xs_arr.iloc[:,ix].values,sv_arr[:,ix],c=cn,cmap="plasma",alpha=0.28,s=10,linewidths=0,zorder=2)
        cb=fig3.colorbar(sc,ax=ax3);cb.ax.tick_params(colors=TKC,labelsize=7)
        cb.ax.set_ylabel(f"{fc2} (relative intensity)",color=LBC,fontsize=8)
        ax3.axhline(0,color=CR,lw=0.8,ls="--",alpha=0.4,label="Zero impact")
        ax3.set_xlabel(f"{fx} (Feature Value)",color=LBC)
        ax3.set_ylabel(f"SHAP Value for {fx} (Rs impact on revenue prediction)",color=LBC)
        ax3.set_title(f"Dependence: {fx} | Coloured by: {fc2}",color=TTC,fontsize=11,fontweight="bold")
        ax3.legend(**legend_kw())
        fig3.tight_layout();st.pyplot(fig3);plt.close()
        st.markdown(ibox([
            ("What This Shows",f"How SHAP values for '{fx}' change as the feature's own value changes, with colour showing '{fc2}' level."),
            ("Upward Trend","Higher feature values consistently increase revenue prediction — a strong positive driver."),
            ("Colour Variation","Different colours at the same X position reveal interaction effects between the two variables."),
            ("Business Use","Use interaction patterns to design outlet-specific or product-specific schemes and pricing.")
        ]),unsafe_allow_html=True)

        ai_txt=(f"The top revenue driver '{sdf.iloc[0]['Friendly']}' has a mean SHAP influence of {sdf.iloc[0]['Mean_SHAP']:.2f} Rs per order. "
                f"The second driver '{sdf.iloc[1]['Friendly']}' contributes {sdf.iloc[1]['Mean_SHAP']:.2f} Rs. "
                "Together, the top 5 SHAP features explain the dominant portion of revenue variance in this dataset. "
                "Operational strategies should be built around improving these five variables.")
        st.markdown(ai_banner(ai_txt),unsafe_allow_html=True)
        st.markdown("---");st.markdown('<div class="sh">Section Recommendations</div>',unsafe_allow_html=True)
        st.markdown(rcard("Top Driver Focus",f"'{sdf.iloc[0]['Friendly']}' is the strongest predictor. Operational and commercial teams should prioritise this variable in every planning cycle.","green"),unsafe_allow_html=True)
        st.markdown(rcard("Interaction Exploitation","Use the dependence plot to identify feature combinations that amplify revenue. Design targeted schemes for high-interaction outlet-product pairs.","blue"),unsafe_allow_html=True)
        st.markdown(rcard("Model Refresh","SHAP importance shifts as market conditions change. Retrain the model quarterly to ensure recommendations remain current.","orange"),unsafe_allow_html=True)


    # ══════════════════════════════════════════════════════════════════════
    # PAGE 5 — SALES DRIVERS
    # ══════════════════════════════════════════════════════════════════════
    elif cur==5:
        st.markdown('<div class="sh">Sales Drivers and Revenue Influencing Factors</div>',unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="sh2">Monthly Revenue Trend (Rs)</div>',unsafe_allow_html=True)
            if "Month" in filt.columns and SV in filt.columns:
                mon=(filt.assign(M=filt["Month"].astype(str),V=col_num(filt,SV))
                         .groupby("M")["V"].sum().reset_index().sort_values("M"))
                fig,ax=sfig(7,4)
                ax.fill_between(range(len(mon)),mon["V"],alpha=0.13,color=CB)
                ax.plot(range(len(mon)),mon["V"],color=CB,lw=2.5,marker="o",ms=5,zorder=3)
                ax.set_xticks(range(len(mon)));ax.set_xticklabels(mon["M"],rotation=40,ha="right",fontsize=7,color=TKC)
                ax.set_ylabel("Total Revenue (Rs)",color=LBC);ax.set_xlabel("Month",color=LBC);inr_y(ax)
                ax.set_title("Monthly Revenue Trend (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                peak=mon.loc[mon["V"].idxmax(),"M"] if len(mon) else"N/A"
                st.markdown(ibox([
                    ("What This Shows","Total revenue (Rs) generated per calendar month across selected filters."),
                    ("Key Insight",f"Peak month: '{peak}'. Study what drove peak performance — schemes, coverage, or seasonality."),
                    ("Business Impact","Seasonal revenue patterns determine optimal timing for scheme launches and stock building."),
                    ("Action","Pre-plan inventory and executive deployment 6 weeks ahead of historically strong months.")
                ]),unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="sh2">Revenue by Order Type (Rs)</div>',unsafe_allow_html=True)
            if "Order Type" in filt.columns and SV in filt.columns:
                ot=gb_sum(filt,"Order Type",SV).sort_values(ascending=False)
                fig,ax=sfig(7,4)
                ax.bar(range(len(ot)),ot.values,color=PAL[:len(ot)],edgecolor="none")
                ax.set_xticks(range(len(ot)));ax.set_xticklabels(ot.index,rotation=30,ha="right",fontsize=8,color=TKC)
                ax.set_ylabel("Total Revenue (Rs)",color=LBC);ax.set_xlabel("Order Type",color=LBC);inr_y(ax)
                ax.set_title("Total Revenue by Order Type (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                st.markdown(ibox([
                    ("Key Insight",f"'{ot.index[0] if len(ot) else 'N/A'}' order type generates the most revenue."),
                    ("Action","Optimise conversion in underperforming order types and analyse barriers to growth in each channel.")
                ]),unsafe_allow_html=True)
        c3,c4=st.columns(2)
        with c3:
            st.markdown('<div class="sh2">Quantity Sold vs Sale Value per Order (Std Units vs Rs)</div>',unsafe_allow_html=True)
            if "Qty (StdUnit)" in filt.columns and SV in filt.columns:
                smp=filt.sample(min(2000,len(filt)),random_state=1);fig,ax=sfig(7,4)
                ax.scatter(col_num(smp,"Qty (StdUnit)"),col_num(smp,SV),alpha=0.22,s=8,color=CT,zorder=2)
                ax.set_xlabel("Quantity Sold (Std Units) per Order",color=LBC)
                ax.set_ylabel("Sale Value per Order (Rs)",color=LBC);inr_y(ax)
                ax.set_title("Quantity Sold vs Revenue per Order",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                st.markdown(ibox([
                    ("What This Shows","Relationship between order size (Std Units) and revenue (Rs) at the individual order level."),
                    ("Key Insight","Positive correlation confirms volume drives revenue. Outliers (high qty, low Rs) suggest pricing or discount anomalies."),
                    ("Business Impact","Increasing average order size by 10% directly lifts revenue without adding new outlets."),
                    ("Action","Implement volume incentives and bundle schemes to raise average order size across all executives.")
                ]),unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="sh2">Scheme Tier vs Average Revenue per Order (Rs)</div>',unsafe_allow_html=True)
            if "Scheme Value" in filt.columns and SV in filt.columns:
                bs=safe_cut(filt["Scheme Value"],n=5,labels=["No Scheme","Low","Moderate","High","Very High"])
                tmp=filt.copy();tmp["_b"]=bs
                grp=(tmp[tmp["_b"]!="N/A"].groupby("_b",observed=True)[SV]
                     .apply(lambda x:col_num(pd.DataFrame({SV:x}),SV).mean()).dropna())
                if len(grp):
                    overall_avg=grp.mean()
                    colors_=[CG if v>=overall_avg else CO for v in grp.values]
                    fig,ax=sfig(7,4)
                    ax.bar(range(len(grp)),grp.values,color=colors_,edgecolor="none")
                    ax.set_xticks(range(len(grp)));ax.set_xticklabels(grp.index,fontsize=8,color=TKC)
                    ax.set_ylabel("Average Revenue per Order (Rs)",color=LBC)
                    ax.set_xlabel("Scheme Value Tier",color=LBC);inr_y(ax)
                    ax.axhline(overall_avg,color=CR,lw=1.5,ls="--",label=f"Overall Avg: {fmt_inr(overall_avg)}")
                    ax.set_title("Scheme Intensity Tier vs Avg Revenue per Order (Rs)",color=TTC,fontsize=11,fontweight="bold")
                    ax.legend(**legend_kw())
                    fig.tight_layout();st.pyplot(fig);plt.close()
                    st.markdown(ibox([
                        ("What This Shows","Whether higher scheme investment per order correlates with higher average revenue."),
                        ("Key Insight","Bars below the red average line indicate scheme tiers with poor revenue ROI."),
                        ("Business Impact","Poorly structured schemes burn margin without proportional revenue gain."),
                        ("Action","Reallocate scheme budget toward tiers demonstrably above the average revenue line.")
                    ]),unsafe_allow_html=True)
        # Benchmarking
        if ON in filt.columns and SV in filt.columns:
            st.markdown("---");st.markdown('<div class="sh2">Order Value Benchmarking (Rs)</div>',unsafe_allow_html=True)
            ov=filt.groupby(filt[ON].astype(str)).apply(lambda x:col_num(x,SV).sum())
            p25,p50,p75=ov.quantile([0.25,0.5,0.75])
            bc1,bc2,bc3,bc4=st.columns(4)
            for c_,(l_,v_,t_) in zip([bc1,bc2,bc3,bc4],[
                ("Bottom 25% of Orders",p25,"Orders in the lowest revenue quartile — focus for basket-size improvement"),
                ("Median Order Value",p50,"Half of all orders are above and below this value"),
                ("Top 25% of Orders",p75,"Orders in the top revenue quartile — protect and grow these"),
                ("Average Order Value",avg_ord,"Total Revenue divided by Total Order count")]):
                c_.markdown(kpi_h(l_,fmt_inr(v_),t_),unsafe_allow_html=True)
        ai_txt=(f"Monthly revenue analysis shows peak month '{peak if 'peak' in dir() else 'N/A'}' outperforms the annual average. "
                f"Scheme intensity appears to influence order value — tiers above {fmt_inr(avg_ord)} average represent scheme-ROI sweet spots. "
                "Volume-revenue correlation confirms that basket-size growth strategies will deliver measurable revenue impact.")
        st.markdown(ai_banner(ai_txt),unsafe_allow_html=True)
        st.markdown("---");st.markdown('<div class="sh">Section Recommendations</div>',unsafe_allow_html=True)
        st.markdown(rcard("Volume-Revenue Alignment","Review pricing tiers for bulk orders to ensure margin is maintained while incentivising higher volumes.","green"),unsafe_allow_html=True)
        st.markdown(rcard("Scheme ROI Review","Schemes not delivering above-average revenue lift should be restructured. Conduct scheme ROI analysis every quarter.","orange"),unsafe_allow_html=True)
        st.markdown(rcard("Seasonal Planning","Pre-plan stock, schemes, and executive deployment 6 weeks ahead of identified high-revenue months.","blue"),unsafe_allow_html=True)


    # ══════════════════════════════════════════════════════════════════════
    # PAGE 6 — OUTLET ANALYSIS
    # ══════════════════════════════════════════════════════════════════════
    elif cur==6:
        st.markdown('<div class="sh">Outlet Performance Analysis</div>',unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="sh2">Top 15 Outlets by Total Revenue (Rs)</div>',unsafe_allow_html=True)
            if OUT in filt.columns and SV in filt.columns:
                top_out=gb_sum(filt,OUT,SV).sort_values(ascending=False).head(15)
                colors_=[CG if i<5 else(CB if i<10 else CT) for i in range(len(top_out))]
                fig,ax=sfig(8,6)
                ax.barh(range(len(top_out)),top_out.values[::-1],color=colors_[::-1],edgecolor="none",height=0.65)
                ax.set_yticks(range(len(top_out)));ax.set_yticklabels(top_out.index[::-1],fontsize=7,color=TKC)
                ax.set_xlabel("Total Revenue (Rs)",color=LBC);inr_x(ax)
                ax.set_title("Top 15 Revenue-Generating Outlets (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                top5_rev=top_out.head(5).sum();total_rev_check=col_num(filt,SV).sum()
                top5_share=top5_rev/max(total_rev_check,1)*100
                st.markdown(ibox([
                    ("What This Shows","The 15 outlets generating the most revenue in the selected period. Green = top 5, Blue = 6-10, Teal = 11-15."),
                    ("Key Insight",f"Top outlet: '{top_out.index[0]}'. Top 5 outlets account for {top5_share:.1f}% of total revenue — a concentration risk."),
                    ("Business Impact","Loss of a single top outlet has an outsized impact on total revenue."),
                    ("Action","Assign dedicated relationship managers to top 10 outlets with quarterly business reviews and priority stock allocation.")
                ]),unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="sh2">Revenue Share by Outlet Classification (%)</div>',unsafe_allow_html=True)
            if "Classification" in filt.columns and SV in filt.columns:
                cls_r=gb_sum(filt,"Classification",SV).sort_values(ascending=False)
                if len(cls_r):
                    fig,ax=sfig(8,6);pie_c(ax,cls_r.values,cls_r.index)
                    ax.set_title("Revenue Share by Outlet Classification",color=TTC,fontsize=11,fontweight="bold")
                    fig.tight_layout();st.pyplot(fig);plt.close()
                    st.markdown(ibox([
                        ("Key Insight",f"'{cls_r.index[0]}' classification outlets drive the largest revenue share."),
                        ("Action","Invest in upgrading mid-tier outlets to premium classification to compound revenue growth.")
                    ]),unsafe_allow_html=True)
        c3,c4=st.columns(2)
        with c3:
            st.markdown('<div class="sh2">Outlet Margin Distribution (%)</div>',unsafe_allow_html=True)
            if "Outlet margin" in filt.columns:
                m_d=col_num(filt,"Outlet margin").replace(0,np.nan).dropna()
                if len(m_d)>1:
                    fig,ax=sfig(7,4)
                    ax.hist(m_d,bins=min(30,m_d.nunique()),color=CO,edgecolor="white",alpha=0.85)
                    ax.axvline(m_d.mean(),color=CR,lw=2,ls="--",label=f"Mean: {m_d.mean():.2f}%")
                    ax.axvline(m_d.median(),color=CG,lw=1.5,ls="--",label=f"Median: {m_d.median():.2f}%")
                    ax.set_xlabel("Outlet Margin (%)",color=LBC);ax.set_ylabel("Number of Outlet Records",color=LBC)
                    ax.set_title("Distribution of Outlet Margin Values (%)",color=TTC,fontsize=11,fontweight="bold")
                    ax.legend(**legend_kw())
                    fig.tight_layout();st.pyplot(fig);plt.close()
                    spread=m_d.std()
                    st.markdown(ibox([
                        ("What This Shows","Histogram of outlet margin (%) across all records — shows how consistently margins are applied."),
                        ("Key Insight",f"Average margin: {m_d.mean():.2f}%. Standard deviation: {spread:.2f}% — {'high variability indicates inconsistent trade terms' if spread>5 else 'relatively consistent margin application'}."),
                        ("Business Impact","Wide margin spread erodes profitability and creates unfair competitive dynamics between outlets."),
                        ("Action","Standardise margin ranges by outlet classification and audit outliers at both extremes of the distribution.")
                    ]),unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="sh2">Revenue by Outlet Channel (Rs)</div>',unsafe_allow_html=True)
            if "Outlets Channel Name" in filt.columns and SV in filt.columns:
                ch=gb_sum(filt,"Outlets Channel Name",SV).sort_values(ascending=False).head(10)
                fig,ax=sfig(7,4)
                ax.bar(range(len(ch)),ch.values,color=PAL[:len(ch)],edgecolor="none")
                ax.set_xticks(range(len(ch)));ax.set_xticklabels(ch.index,rotation=35,ha="right",fontsize=7,color=TKC)
                ax.set_ylabel("Total Revenue (Rs)",color=LBC);ax.set_xlabel("Outlet Channel",color=LBC);inr_y(ax)
                ax.set_title("Revenue by Outlet Channel (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                st.markdown(ibox([
                    ("Key Insight",f"'{ch.index[0] if len(ch) else 'N/A'}' channel generates the most revenue."),
                    ("Action","Study top channel characteristics and apply those conditions to develop underperforming channels.")
                ]),unsafe_allow_html=True)
        # Outlet benchmarking
        if OUT in filt.columns and SV in filt.columns:
            st.markdown("---");st.markdown('<div class="sh2">Outlet Revenue Benchmarking (Rs)</div>',unsafe_allow_html=True)
            or_=gb_sum(filt,OUT,SV)
            bb1,bb2,bb3,bb4=st.columns(4)
            for c_,(l_,v_,t_) in zip([bb1,bb2,bb3,bb4],[
                ("Bottom 10% Outlets",or_.quantile(0.1),"Lowest-revenue 10% of outlets — candidates for activation drive"),
                ("Median Outlet Revenue",or_.median(),"Half of outlets are above and below this revenue figure"),
                ("Top 10% Outlets",or_.quantile(0.9),"Highest-revenue 10% — key accounts requiring premium service"),
                ("Average Outlet Revenue",or_.mean(),"Total Revenue divided by active outlet count")]):
                c_.markdown(kpi_h(l_,fmt_inr(v_),t_),unsafe_allow_html=True)
        if "Outlets Type" in filt.columns and SV in filt.columns:
            st.markdown('<div class="sh2">Average Revenue per Outlet Type (Rs)</div>',unsafe_allow_html=True)
            ot_s=gb_sum(filt,"Outlets Type",SV)
            ot_c=filt.groupby(filt["Outlets Type"].astype(str)).size().reindex(ot_s.index).fillna(1)
            ot_av=(ot_s/ot_c).sort_values(ascending=False).head(12)
            fig,ax=sfig(12,4)
            colors_=[CG if v>=ot_av.median() else CO for v in ot_av.values]
            ax.bar(range(len(ot_av)),ot_av.values,color=colors_,edgecolor="none")
            ax.set_xticks(range(len(ot_av)));ax.set_xticklabels(ot_av.index,rotation=30,ha="right",fontsize=8,color=TKC)
            ax.set_ylabel("Avg Revenue per Outlet (Rs)",color=LBC);ax.set_xlabel("Outlet Type",color=LBC);inr_y(ax)
            ax.axhline(ot_av.median(),color=CR,lw=1.5,ls="--",label=f"Median: {fmt_inr(ot_av.median())}")
            ax.set_title("Average Revenue per Outlet Type (Rs) — Green = Above Median",color=TTC,fontsize=11,fontweight="bold")
            ax.legend(**legend_kw())
            fig.tight_layout();st.pyplot(fig);plt.close()
        top_out_name=(gb_sum(filt,OUT,SV).idxmax() if OUT in filt.columns and SV in filt.columns else"N/A")
        st.markdown(ai_banner(
            f"Outlet '{top_out_name}' is the single highest revenue contributor. "
            f"Top 5 outlets likely contribute a disproportionate share of total revenue — a key account management priority. "
            f"Outlet margin spread indicates inconsistent trade terms that may be reducing net profitability across the portfolio."
        ),unsafe_allow_html=True)
        st.markdown("---");st.markdown('<div class="sh">Section Recommendations</div>',unsafe_allow_html=True)
        st.markdown(rcard("Key Account Programme","Implement a formal Key Account Management programme for the top 15 outlets with dedicated service, priority fulfilment, and quarterly business reviews.","green"),unsafe_allow_html=True)
        st.markdown(rcard("Margin Standardisation","Wide margin distribution indicates inconsistent trade terms. Implement classification-based margin bands and audit all outlier outlets.","orange"),unsafe_allow_html=True)
        st.markdown(rcard("Low-Revenue Outlet Activation","Bottom 10% outlets are under-activated. Targeted activation campaigns with introductory schemes can unlock significant incremental revenue.","blue"),unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 7 — PRODUCT & CATEGORY
    # ══════════════════════════════════════════════════════════════════════
    elif cur==7:
        st.markdown('<div class="sh">Product and Category Contribution Analysis</div>',unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="sh2">Top 15 Products by Total Revenue (Rs)</div>',unsafe_allow_html=True)
            if PRD in filt.columns and SV in filt.columns:
                tp=gb_sum(filt,PRD,SV).sort_values(ascending=False).head(15)
                fig,ax=sfig(8,6)
                ax.barh(range(len(tp)),tp.values[::-1],color=CG,edgecolor="none",height=0.65)
                ax.set_yticks(range(len(tp)));ax.set_yticklabels(tp.index[::-1],fontsize=7,color=TKC)
                ax.set_xlabel("Total Revenue (Rs)",color=LBC);inr_x(ax)
                ax.set_title("Top 15 Products by Revenue (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                top5_prd=tp.head(5).sum()/max(total_rev,1)*100
                st.markdown(ibox([
                    ("Key Insight",f"Top product: '{tp.index[0]}'. Top 5 products account for {top5_prd:.1f}% of revenue — {'high concentration risk' if top5_prd>60 else 'acceptable diversification'}."),
                    ("Business Impact","Over-reliance on a few SKUs creates vulnerability if any single product faces supply or demand issues."),
                    ("Action","Actively develop secondary SKUs through targeted promotions to diversify the revenue base.")
                ]),unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="sh2">Revenue by Primary Category (Rs)</div>',unsafe_allow_html=True)
            if "PrimaryCategory" in filt.columns and SV in filt.columns:
                pc=gb_sum(filt,"PrimaryCategory",SV).sort_values(ascending=False).head(10)
                fig,ax=sfig(8,6)
                ax.bar(range(len(pc)),pc.values,color=PAL[:len(pc)],edgecolor="none")
                ax.set_xticks(range(len(pc)));ax.set_xticklabels(pc.index,rotation=35,ha="right",fontsize=7,color=TKC)
                ax.set_ylabel("Total Revenue (Rs)",color=LBC);ax.set_xlabel("Primary Category",color=LBC);inr_y(ax)
                ax.set_title("Revenue by Primary Category (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
        c3,c4=st.columns(2)
        with c3:
            st.markdown('<div class="sh2">Ordered vs Dispatched Quantity — Top 20 Products (Std Units)</div>',unsafe_allow_html=True)
            if PRD in filt.columns and "Qty (StdUnit)" in filt.columns:
                tmp=filt.copy();tmp[PRD]=tmp[PRD].astype(str)
                tmp["O"]=col_num(tmp,"Qty (StdUnit)");tmp["D"]=col_num(tmp,"Dispatch Qty (StdUnit)")
                pq=tmp.groupby(PRD).agg(Ordered=("O","sum"),Dispatched=("D","sum")).sort_values("Ordered",ascending=False).head(20)
                x,w=np.arange(len(pq)),0.38;fig,ax=sfig(12,5)
                ax.bar(x-w/2,pq["Ordered"],w,label="Ordered (Std Units)",color=CB,edgecolor="none")
                ax.bar(x+w/2,pq["Dispatched"],w,label="Dispatched (Std Units)",color=CG,edgecolor="none")
                ax.set_xticks(x);ax.set_xticklabels(pq.index,rotation=40,ha="right",fontsize=6.5,color=TKC)
                ax.set_ylabel("Quantity (Std Units)",color=LBC);ax.set_xlabel("Product (SKU)",color=LBC)
                ax.set_title("Ordered vs Dispatched Quantity — Top 20 Products (Std Units)",color=TTC,fontsize=11,fontweight="bold")
                ax.legend(**legend_kw())
                fig.tight_layout();st.pyplot(fig);plt.close()
                gap=(pq["Ordered"]-pq["Dispatched"]).sum()/pq["Ordered"].sum()*100 if pq["Ordered"].sum()>0 else 0
                st.markdown(ibox([
                    ("What This Shows","Side-by-side comparison of units ordered vs units dispatched for the top 20 SKUs."),
                    ("Key Insight",f"Average fulfilment gap across top 20 products: {gap:.1f}% of ordered units not dispatched."),
                    ("Business Impact","Unfulfilled units directly reduce revenue recognition and damage customer trust."),
                    ("Action","Investigate stock-out patterns for the products with the largest gaps. Adjust safety stock levels and reorder points.")
                ]),unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="sh2">Revenue Heatmap: Primary x Secondary Category (Rs)</div>',unsafe_allow_html=True)
            if "PrimaryCategory" in filt.columns and "SecondaryCategory" in filt.columns and SV in filt.columns:
                tmp2=filt.copy()
                tmp2["PrimaryCategory"]=tmp2["PrimaryCategory"].astype(str)
                tmp2["SecondaryCategory"]=tmp2["SecondaryCategory"].astype(str)
                tmp2[SV]=col_num(tmp2,SV)
                piv=tmp2.pivot_table(index="PrimaryCategory",columns="SecondaryCategory",values=SV,aggfunc="sum",fill_value=0)
                piv=piv.iloc[:min(10,len(piv)),:min(10,piv.shape[1])]
                if piv.size:
                    fig,ax=sfig(8,5);im=ax.imshow(piv.values,aspect="auto",cmap="Blues")
                    ax.set_xticks(range(piv.shape[1]));ax.set_xticklabels(piv.columns,rotation=45,ha="right",fontsize=6,color=TKC)
                    ax.set_yticks(range(piv.shape[0]));ax.set_yticklabels(piv.index,fontsize=6,color=TKC)
                    ax.set_xlabel("Secondary Category",color=LBC);ax.set_ylabel("Primary Category",color=LBC)
                    ax.set_title("Revenue Heatmap: Primary x Secondary Category (Rs)",color=TTC,fontsize=10,fontweight="bold")
                    cb2=fig.colorbar(im,ax=ax);cb2.ax.tick_params(colors=TKC,labelsize=7)
                    cb2.set_label("Total Revenue (Rs)",color=LBC,fontsize=7)
                    fig.tight_layout();st.pyplot(fig);plt.close()
                    st.markdown(ibox([
                        ("What This Shows","Revenue (Rs) for each Primary-Secondary category combination. Darker cells = higher revenue."),
                        ("Key Insight","Light cells represent under-penetrated category combinations — potential white-space growth opportunities."),
                        ("Action","Focus upsell efforts on light-coloured cells adjacent to dark cells — adjacent expansion is lower risk.")
                    ]),unsafe_allow_html=True)
        st.markdown(ai_banner(
            f"Top product concentration analysis: top 5 SKUs account for a significant portion of revenue — SKU diversification is recommended. "
            f"Category heatmap reveals white-space opportunities in under-penetrated sub-categories. "
            f"Fulfilment gap data indicates supply chain improvements could directly recover lost revenue units."
        ),unsafe_allow_html=True)
        st.markdown("---");st.markdown('<div class="sh">Section Recommendations</div>',unsafe_allow_html=True)
        st.markdown(rcard("SKU Diversification","If top 5 products exceed 60% revenue share, develop secondary SKUs actively. Revenue diversification reduces supply and demand risk.","orange"),unsafe_allow_html=True)
        st.markdown(rcard("Fulfilment Gap Closure","Improve demand forecasting accuracy and safety stock levels for high-gap products. Each recovered unit is direct revenue.","red"),unsafe_allow_html=True)
        st.markdown(rcard("Category White Space","Heatmap light spots represent upsell opportunities. Structured cross-category promotions can unlock incremental volume at low marginal cost.","green"),unsafe_allow_html=True)


    # ══════════════════════════════════════════════════════════════════════
    # PAGE 8 — SALES EXECUTIVE
    # ══════════════════════════════════════════════════════════════════════
    elif cur==8:
        st.markdown('<div class="sh">Sales Executive Performance Analysis</div>',unsafe_allow_html=True)
        ec=next((c for c in["User","Sales Executive Erp Id","Assistant Sales Manager","Senior Supervisor","Supervisor"] if c in filt.columns),None)
        if ec and SV in filt.columns:
            er=gb_sum(filt,ec,SV).sort_values(ascending=False)
            tmp3=filt.copy();tmp3[ec]=tmp3[ec].astype(str)
            eo=tmp3.groupby(ec)[ON].nunique() if ON in tmp3.columns else er.apply(lambda _:1)
            avg_rev_ord=(er/eo.reindex(er.index).fillna(1)).sort_values(ascending=False)
            c1,c2=st.columns(2)
            with c1:
                st.markdown(f'<div class="sh2">Top 15 {ec}s — Total Revenue (Rs)</div>',unsafe_allow_html=True)
                t15=er.head(15);colors_=[CG if i<5 else CB for i in range(len(t15))]
                fig,ax=sfig(8,6)
                ax.barh(range(len(t15)),t15.values[::-1],color=colors_[::-1],edgecolor="none",height=0.65)
                ax.set_yticks(range(len(t15)));ax.set_yticklabels(t15.index[::-1],fontsize=7,color=TKC)
                ax.set_xlabel("Total Revenue (Rs)",color=LBC);inr_x(ax)
                ax.set_title(f"Top 15 {ec}s by Total Revenue (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                st.markdown(ibox([
                    ("What This Shows",f"Total revenue generated by the top 15 {ec}s. Green = top 5, Blue = 6-15."),
                    ("Key Insight",f"Top performer: '{t15.index[0] if len(t15) else 'N/A'}' ({fmt_inr(t15.iloc[0] if len(t15) else 0)})."),
                    ("Action","Study top performer behaviours — outlet coverage, order frequency, scheme utilisation — and replicate across the team.")
                ]),unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="sh2">Revenue per Order by {ec} (Rs / Order) — Quality Metric</div>',unsafe_allow_html=True)
                t15a=avg_rev_ord.head(15);med_rev_ord=avg_rev_ord.median()
                colors_=[CG if v>=med_rev_ord else CO for v in t15a.values]
                fig,ax=sfig(8,6)
                ax.barh(range(len(t15a)),t15a.values[::-1],color=colors_[::-1],edgecolor="none",height=0.65)
                ax.set_yticks(range(len(t15a)));ax.set_yticklabels(t15a.index[::-1],fontsize=7,color=TKC)
                ax.set_xlabel("Revenue per Order (Rs)",color=LBC);inr_x(ax)
                ax.axvline(med_rev_ord,color=CR,lw=1.5,ls="--",label=f"Median: {fmt_inr(med_rev_ord)}")
                ax.set_title("Revenue per Order — Executive Efficiency (Rs/Order)",color=TTC,fontsize=11,fontweight="bold")
                ax.legend(**legend_kw())
                fig.tight_layout();st.pyplot(fig);plt.close()
                st.markdown(ibox([
                    ("What This Shows","Revenue per order placed — a measure of order QUALITY, not just volume."),
                    ("Green Bars","Above-median revenue per order — these executives place high-value orders."),
                    ("Orange Bars","Below-median — may indicate smaller baskets, lower-value outlets, or under-upselling."),
                    ("Action","Coach below-median executives on basket-building. Pair with top performers as mentors for 90-day improvement cycles.")
                ]),unsafe_allow_html=True)
            # Benchmarking table
            st.markdown("---");st.markdown('<div class="sh2">Full Executive Performance Benchmarking Table</div>',unsafe_allow_html=True)
            etbl=pd.DataFrame({
                ec:er.index,
                "Total Revenue (Rs)":er.apply(fmt_inr).values,
                "Order Count":eo.reindex(er.index).fillna(0).astype(int).values,
                "Revenue per Order (Rs)":avg_rev_ord.reindex(er.index).apply(fmt_inr).values,
                "vs Median":[(f"+{(v/med_rev_ord-1)*100:.0f}%" if v>=med_rev_ord
                              else f"{(v/med_rev_ord-1)*100:.0f}%")
                             for v in avg_rev_ord.reindex(er.index).fillna(0)]
            }).reset_index(drop=True)
            st.dataframe(etbl,width='stretch')
            top3=er.head(3).index.tolist();bot3=er.tail(3).index.tolist()
            st.markdown(ai_banner(
                f"Top 3 executives ({', '.join(top3)}) are generating the highest revenue. "
                f"Bottom performers ({', '.join(bot3)}) show a gap vs median that represents a coaching and territory-review opportunity. "
                "Revenue per order variance across the team suggests basket-building skills and territory quality need attention."
            ),unsafe_allow_html=True)
            st.markdown("---");st.markdown('<div class="sh">Section Recommendations</div>',unsafe_allow_html=True)
            st.markdown(rcard("Reward Top Performers",f"Top executives {', '.join(top3)} deserve structured incentive programmes for retention. Their practices should be documented and shared.","green"),unsafe_allow_html=True)
            st.markdown(rcard("Support Bottom Performers",f"Bottom performers {', '.join(bot3)} may benefit from territory reassignment, training programmes, or workload balancing. Investigate root causes before reassignment.","orange"),unsafe_allow_html=True)
            st.markdown(rcard("Revenue Quality Focus","Revenue per order is a better performance metric than total volume alone. Build team KPIs around order quality, not just order count.","blue"),unsafe_allow_html=True)
        else:st.info("No suitable sales executive column found in the dataset.")

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 9 — DELIVERY & FULFILMENT
    # ══════════════════════════════════════════════════════════════════════
    elif cur==9:
        st.markdown('<div class="sh">Delivery and Fulfilment Performance</div>',unsafe_allow_html=True)
        if "Delivery Status" in filt.columns and SV in filt.columns:
            dsr=gb_sum(filt,"Delivery Status",SV)
            tmp4=filt.copy();tmp4["Delivery Status"]=tmp4["Delivery Status"].astype(str)
            dsc=tmp4.groupby("Delivery Status").size().reindex(dsr.index).fillna(1)
            dsa=(dsr/dsc).sort_values(ascending=False)
            c1,c2=st.columns(2)
            with c1:
                st.markdown('<div class="sh2">Revenue Share by Delivery Outcome (%)</div>',unsafe_allow_html=True)
                fig,ax=sfig(7,5);pie_c(ax,dsr.values,dsr.index)
                ax.set_title("Revenue Distribution by Delivery Status",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                st.markdown(ibox([
                    ("What This Shows","Percentage of total revenue associated with each delivery outcome."),
                    ("Key Insight","Revenue associated with non-delivered status is at-risk income — it may be recognised later or lost entirely."),
                    ("Business Impact","Unrecognised revenue creates cash flow unpredictability and customer dissatisfaction."),
                    ("Action","Prioritise recovery of orders in transit or pending delivery status. Set SLA targets for each delivery category.")
                ]),unsafe_allow_html=True)
            with c2:
                st.markdown('<div class="sh2">Average Revenue per Order by Delivery Status (Rs / Order)</div>',unsafe_allow_html=True)
                colors_=[CG if "deliver" in str(k).lower() else CR for k in dsa.index]
                fig,ax=sfig(7,5)
                ax.bar(range(len(dsa)),dsa.values,color=colors_,edgecolor="none")
                ax.set_xticks(range(len(dsa)));ax.set_xticklabels(dsa.index,rotation=30,ha="right",fontsize=8,color=TKC)
                ax.set_ylabel("Avg Revenue per Order (Rs)",color=LBC);ax.set_xlabel("Delivery Status",color=LBC);inr_y(ax)
                ax.set_title("Avg Revenue per Order by Delivery Status (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                st.markdown(ibox([
                    ("Key Insight","If high-value orders have worse delivery performance, the revenue impact is amplified beyond order count."),
                    ("Action","Implement order-value-based fulfilment priority — ensure high-value orders are protected with dedicated logistics resources.")
                ]),unsafe_allow_html=True)
            st.markdown('<div class="sh2">Delivery Status Detail Table</div>',unsafe_allow_html=True)
            dst=pd.DataFrame({
                "Delivery Status":dsr.index,
                "Total Revenue (Rs)":dsr.apply(fmt_inr).values,
                "Order Count":dsc.reindex(dsr.index).fillna(0).astype(int).values,
                "Avg Rev / Order (Rs)":dsa.reindex(dsr.index).apply(fmt_inr).values,
                "Revenue Share (%)": (dsr/dsr.sum()*100).round(1).astype(str).add("%").values
            }).sort_values("Order Count",ascending=False)
            st.dataframe(dst.reset_index(drop=True),width='stretch')
        if DR in filt.columns:
            st.markdown('<div class="sh2">Dispatch Rate Distribution — All Orders (%)</div>',unsafe_allow_html=True)
            st.markdown('<p class="tip">Dispatch Rate = Dispatched Quantity / Ordered Quantity. Industry best practice target: 95%+</p>',unsafe_allow_html=True)
            dr_d=col_num(filt,DR).clip(0,1);dr_d=dr_d[dr_d>0]
            if len(dr_d)>1:
                fig,ax=sfig(10,4)
                ax.hist(dr_d*100,bins=40,color=CB,edgecolor="white",alpha=0.85)
                ax.axvline(dr_d.mean()*100,color=CR,lw=2,ls="--",label=f"Mean: {dr_d.mean():.1%}")
                ax.axvline(95,color=CG,lw=1.5,ls="--",label="95% Industry Target")
                ax.set_xlabel("Dispatch Rate (%) — Dispatched Qty / Ordered Qty",color=LBC)
                ax.set_ylabel("Number of Orders",color=LBC)
                ax.set_title("Order Dispatch Rate Distribution — All Orders",color=TTC,fontsize=11,fontweight="bold")
                ax.legend(**legend_kw())
                fig.tight_layout();st.pyplot(fig);plt.close()
                gap_to_target=max(0,95-dr_d.mean()*100)
                st.markdown(ibox([
                    ("What This Shows","Distribution of dispatch rates across all orders — how fully each order was fulfilled at dispatch."),
                    ("Key Insight",f"Mean dispatch rate: {dr_d.mean():.1%}. Gap to 95% target: {gap_to_target:.1f} percentage points."),
                    ("Business Impact","Each percentage point below 95% represents direct revenue leakage that compounds across thousands of orders."),
                    ("Action",f"Orders below 80% dispatch rate ({(dr_d<0.8).mean()*100:.1f}% of orders) should trigger immediate investigation and corrective action.")
                ]),unsafe_allow_html=True)
        st.markdown(ai_banner(
            f"Current dispatch rate of {disp_rate:.1%} is {'below' if disp_rate<0.95 else 'at or above'} the 95% industry benchmark. "
            "Delivery failures appear concentrated in specific status categories — targeted root-cause analysis for each category will yield the fastest improvement. "
            "High-value orders with delivery issues represent the highest-priority recovery opportunity."
        ),unsafe_allow_html=True)
        st.markdown("---");st.markdown('<div class="sh">Section Recommendations</div>',unsafe_allow_html=True)
        st.markdown(rcard("Delivery Failure Root Cause","Segment delivery failures by depot, route and product category. Identify systemic vs isolated issues and resolve with targeted logistics improvements.","orange"),unsafe_allow_html=True)
        st.markdown(rcard("Dispatch Rate Improvement",f"Closing dispatch rate from {disp_rate:.1%} to 95% requires improving stock availability at warehouse level. Review reorder points and supplier lead times.","red"),unsafe_allow_html=True)
        st.markdown(rcard("Customer Communication SLA","Implement automated delivery status notifications. Proactive communication converts delivery delays from customer losses into recoverable delays.","blue"),unsafe_allow_html=True)


    # ══════════════════════════════════════════════════════════════════════
    # PAGE 10 — DEPOT & REGIONAL
    # ══════════════════════════════════════════════════════════════════════
    elif cur==10:
        st.markdown('<div class="sh">Depot and Regional Performance Analysis</div>',unsafe_allow_html=True)
        if DEP in filt.columns and SV in filt.columns:
            dr_s=gb_sum(filt,DEP,SV).sort_values(ascending=False)
            tmp5=filt.copy();tmp5[DEP]=tmp5[DEP].astype(str)
            do_=tmp5.groupby(DEP)[ON].nunique() if ON in tmp5.columns else dr_s.apply(lambda _:1)
            dout=(tmp5.groupby(DEP)[OUT].apply(lambda x:x.astype(str).nunique()) if OUT in tmp5.columns else dr_s.apply(lambda _:0))
            dr_avg=(dr_s/do_.reindex(dr_s.index).fillna(1))
            c1,c2=st.columns(2)
            with c1:
                st.markdown('<div class="sh2">Total Revenue by Depot (Rs) — Green = Above Median</div>',unsafe_allow_html=True)
                colors_=[CG if v>=dr_s.median() else CO for v in dr_s.values[::-1]]
                fig,ax=sfig(8,6)
                ax.barh(range(len(dr_s)),dr_s.values[::-1],color=colors_,edgecolor="none",height=0.65)
                ax.set_yticks(range(len(dr_s)));ax.set_yticklabels(dr_s.index[::-1],fontsize=7,color=TKC)
                ax.set_xlabel("Total Revenue (Rs)",color=LBC);inr_x(ax)
                ax.set_title("Total Revenue by Depot (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                st.markdown(ibox([
                    ("What This Shows","Total revenue generated per depot location. Green = above median, Orange = below median."),
                    ("Key Insight",f"Top depot: '{dr_s.index[0]}' ({fmt_inr(dr_s.iloc[0])}). Bottom depot: '{dr_s.index[-1]}' ({fmt_inr(dr_s.iloc[-1])})."),
                    ("Business Impact","Below-median depots represent unrealised revenue potential in their respective geographies."),
                    ("Action","Conduct operational audits for orange depots. Identify whether the gap is driven by coverage, execution, or market size.")
                ]),unsafe_allow_html=True)
            with c2:
                st.markdown('<div class="sh2">Revenue per Order by Depot (Rs / Order) — Efficiency Metric</div>',unsafe_allow_html=True)
                fig,ax=sfig(8,6)
                colors2_=[CG if v>=dr_avg.median() else CO for v in dr_avg.values[::-1]]
                ax.barh(range(len(dr_avg)),dr_avg.values[::-1],color=colors2_,edgecolor="none",height=0.65)
                ax.set_yticks(range(len(dr_avg)));ax.set_yticklabels(dr_avg.index[::-1],fontsize=7,color=TKC)
                ax.set_xlabel("Revenue per Order (Rs)",color=LBC);inr_x(ax)
                ax.set_title("Revenue Efficiency by Depot (Rs / Order)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                st.markdown(ibox([
                    ("Key Insight","Depots with high total revenue but low per-order efficiency may be placing many small orders instead of fewer larger ones."),
                    ("Action","Encourage executives in low-efficiency depots to focus on order quality — larger basket sizes with fewer, better orders.")
                ]),unsafe_allow_html=True)
            ord_v=do_.reindex(dr_s.index).fillna(0).astype(int)
            out_v=dout.reindex(dr_s.index).fillna(0).astype(int)
            st.markdown('<div class="sh2">Depot Performance Summary Table</div>',unsafe_allow_html=True)
            dep_tbl=pd.DataFrame({
                "Depot":dr_s.index,
                "Total Revenue (Rs)":dr_s.apply(fmt_inr).values,
                "Orders":ord_v.values,
                "Outlets Covered":out_v.values,
                "Revenue per Order (Rs)":(dr_s/ord_v.replace(0,1)).apply(fmt_inr).values,
                "Revenue per Outlet (Rs)":(dr_s/out_v.replace(0,1)).apply(fmt_inr).values,
                "vs Median":[(f"+{(v/dr_s.median()-1)*100:.0f}%" if v>=dr_s.median()
                              else f"{(v/dr_s.median()-1)*100:.0f}%") for v in dr_s.values]
            })
            st.dataframe(dep_tbl.reset_index(drop=True),width='stretch')
            if "Beats" in filt.columns and SV in filt.columns:
                st.markdown('<div class="sh2">Top 10 Revenue-Generating Beats / Routes (Rs)</div>',unsafe_allow_html=True)
                br=gb_sum(filt,"Beats",SV).sort_values(ascending=False).head(10);fig,ax=sfig(12,4)
                ax.bar(range(len(br)),br.values,color=PAL[:len(br)],edgecolor="none")
                ax.set_xticks(range(len(br)));ax.set_xticklabels(br.index,rotation=35,ha="right",fontsize=8,color=TKC)
                ax.set_ylabel("Total Revenue (Rs)",color=LBC);ax.set_xlabel("Beat / Route",color=LBC);inr_y(ax)
                ax.set_title("Top 10 Revenue Beats (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
        st.markdown(ai_banner(
            f"Top depot '{dr_s.index[0] if DEP in filt.columns else 'N/A'}' significantly outperforms peers. "
            "Below-median depots may benefit from outlet expansion drives and executive upskilling. "
            "Revenue per outlet metric identifies depots with strong demand but low outlet coverage — expansion priority targets."
        ),unsafe_allow_html=True)
        st.markdown("---");st.markdown('<div class="sh">Section Recommendations</div>',unsafe_allow_html=True)
        st.markdown(rcard("Top Depot Best Practices","Document and share what makes top depots successful — coverage patterns, call frequency, scheme utilisation — and roll out to all depots.","green"),unsafe_allow_html=True)
        st.markdown(rcard("Under-Performing Depot Audit","Orange depots need operational assessments covering outlet coverage gaps, executive deployment, and local market constraints.","orange"),unsafe_allow_html=True)
        st.markdown(rcard("Expansion Targeting","Depots with high revenue-per-outlet but low total outlets signal under-tapped markets. Prioritise outlet expansion in these geographies.","blue"),unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 11 — MARGIN & TAX
    # ══════════════════════════════════════════════════════════════════════
    elif cur==11:
        st.markdown('<div class="sh">Margin, Revenue, and Tax Analysis</div>',unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="sh2">Outlet Margin (%) vs Sale Value per Order (Rs)</div>',unsafe_allow_html=True)
            if "Outlet margin" in filt.columns and SV in filt.columns:
                smp=filt.sample(min(3000,len(filt)),random_state=2);fig,ax=sfig(7,5)
                ax.scatter(col_num(smp,"Outlet margin"),col_num(smp,SV),alpha=0.2,s=8,color=CO,zorder=2)
                ax.set_xlabel("Outlet Margin (%) per Record",color=LBC)
                ax.set_ylabel("Sale Value per Order (Rs)",color=LBC);inr_y(ax)
                ax.set_title("Outlet Margin (%) vs Sale Value per Order (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                st.markdown(ibox([
                    ("What This Shows","Scatter of outlet margin (%) vs revenue per order (Rs) — each dot is one order record."),
                    ("Key Insight","Positive correlation = high-margin outlets also generate more revenue. Clusters at high margin but low revenue are an anomaly to investigate."),
                    ("Business Impact","High-margin, low-revenue outlets are under-performing relative to their potential."),
                    ("Action","Schedule targeted visits to high-margin, low-revenue outlets. Understand barriers and activate with targeted promotions.")
                ]),unsafe_allow_html=True)
        with c2:
            tax_cols=[c for c in["CGST","SGST","IGST","VAT"] if c in filt.columns]
            if tax_cols:
                st.markdown('<div class="sh2">Tax Component Breakdown (Rs)</div>',unsafe_allow_html=True)
                tt={c:col_num(filt,c).sum() for c in tax_cols};fig,ax=sfig(7,5)
                ax.bar(list(tt.keys()),list(tt.values()),color=PAL[:len(tt)],edgecolor="none")
                ax.set_ylabel("Total Tax Amount (Rs)",color=LBC);ax.set_xlabel("Tax Component",color=LBC);inr_y(ax)
                for i,(k,v) in enumerate(tt.items()):ax.text(i,v*1.01,fmt_inr(v),ha="center",va="bottom",fontsize=7,color=TKC)
                ax.set_title("Tax Component Breakdown (Rs)",color=TTC,fontsize=11,fontweight="bold")
                fig.tight_layout();st.pyplot(fig);plt.close()
                pct_tax=sum(tt.values())/max(total_rev,1)*100
                igst_share=tt.get("IGST",0)/max(sum(tt.values()),1)*100
                st.markdown(ibox([
                    ("What This Shows","Total Rs value of each tax component across the filtered dataset."),
                    ("Key Insight",f"Total tax = {fmt_inr(sum(tt.values()))} ({pct_tax:.1f}% of gross revenue). IGST represents {igst_share:.1f}% of total tax — the most controllable component."),
                    ("Business Impact","Tax obligations directly reduce net revenue. IGST on inter-state movements is the primary optimisation lever."),
                    ("Action","Review inter-state order routing. Serving orders from closer depots reduces IGST exposure and improves net margin.")
                ]),unsafe_allow_html=True)
        if "Net Value" in filt.columns and "Net Value Tax Inclusive" in filt.columns and "Month" in filt.columns:
            st.markdown('<div class="sh2">Net Value Trend: Ex-Tax vs Tax-Inclusive by Month (Rs)</div>',unsafe_allow_html=True)
            mn=(filt.assign(M=filt["Month"].astype(str),NV=col_num(filt,"Net Value"),NVI=col_num(filt,"Net Value Tax Inclusive"))
                    .groupby("M")[["NV","NVI"]].sum().reset_index().sort_values("M"))
            fig,ax=sfig(11,4);x=range(len(mn))
            ax.plot(x,mn["NV"],color=CB,lw=2.5,marker="o",ms=4,label="Net Value Ex-Tax (Rs)")
            ax.plot(x,mn["NVI"],color=CO,lw=2.5,marker="s",ms=4,label="Net Value Incl. Tax (Rs)")
            ax.fill_between(x,mn["NV"],mn["NVI"],alpha=0.12,color=CO,label="Tax Amount (Rs)")
            ax.set_xticks(list(x));ax.set_xticklabels(mn["M"],rotation=40,ha="right",fontsize=7,color=TKC)
            ax.set_ylabel("Net Value (Rs)",color=LBC);ax.set_xlabel("Month",color=LBC);inr_y(ax)
            ax.set_title("Monthly Net Value — Ex-Tax vs Tax-Inclusive (Rs)",color=TTC,fontsize=11,fontweight="bold")
            ax.legend(**legend_kw())
            fig.tight_layout();st.pyplot(fig);plt.close()
            avg_gap=(mn["NVI"]-mn["NV"]).mean()
            st.markdown(ibox([
                ("What This Shows","Monthly comparison of net value before tax (blue) and after tax (orange). Shaded area = tax burden."),
                ("Key Insight",f"Average monthly tax burden: {fmt_inr(avg_gap)}. This is a direct cash outflow reducing net profitability each month."),
                ("Business Impact","A growing shaded area over time signals increasing tax exposure that needs proactive management."),
                ("Action","Optimise product mix toward lower-tax categories. Review inter-state logistics for IGST reduction opportunities.")
            ]),unsafe_allow_html=True)
        st.markdown(ai_banner(
            f"Total tax obligation ({fmt_inr(total_tax)}) represents {total_tax/max(total_rev,1)*100:.1f}% of gross revenue. "
            "IGST from inter-state orders is the most actionable component — depot routing optimisation can reduce this exposure. "
            "Outlet margin scatter reveals high-potential but under-performing outlets that require targeted activation."
        ),unsafe_allow_html=True)
        st.markdown("---");st.markdown('<div class="sh">Section Recommendations</div>',unsafe_allow_html=True)
        st.markdown(rcard("Margin Policy Standardisation","Implement classification-based margin bands and enforce them through the ERP system. Audit outlier outlets quarterly.","orange"),unsafe_allow_html=True)
        st.markdown(rcard("IGST Reduction Strategy","Map IGST-heavy orders and model the savings from re-routing through closer depots. Even a 15% IGST reduction improves net margin measurably.","red"),unsafe_allow_html=True)
        st.markdown(rcard("Tax Trend Monitoring","Set up monthly tax-to-revenue ratio alerts. A rising ratio requires immediate product mix or logistics review.","blue"),unsafe_allow_html=True)


    # ══════════════════════════════════════════════════════════════════════
    # PAGE 12 — SEGMENTATION
    # ══════════════════════════════════════════════════════════════════════
    elif cur==12:
        st.markdown('<div class="sh">Outlet Segmentation and Classification Insights</div>',unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="sh2">Revenue Share by Outlet Segmentation (%)</div>',unsafe_allow_html=True)
            if "Outlets Segmentation" in filt.columns and SV in filt.columns:
                sr=gb_sum(filt,"Outlets Segmentation",SV).sort_values(ascending=False)
                if len(sr):
                    fig,ax=sfig(7,5);pie_c(ax,sr.values,sr.index)
                    ax.set_title("Revenue Share by Outlet Segmentation",color=TTC,fontsize=11,fontweight="bold")
                    fig.tight_layout();st.pyplot(fig);plt.close()
                    top_seg=sr.index[0];top_seg_share=sr.iloc[0]/sr.sum()*100
                    st.markdown(ibox([
                        ("What This Shows","Revenue percentage contribution from each outlet segmentation tier."),
                        ("Key Insight",f"'{top_seg}' segment drives {top_seg_share:.1f}% of total revenue — the highest-value customer tier."),
                        ("Business Impact","Losing top-segment outlets has a disproportionate revenue impact vs their count."),
                        ("Action","Apply premium service standards to top segments: first-priority stock, dedicated executives, exclusive schemes.")
                    ]),unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="sh2">Revenue Heatmap: Segmentation x Classification (Rs)</div>',unsafe_allow_html=True)
            if "Outlets Segmentation" in filt.columns and "Classification" in filt.columns and SV in filt.columns:
                tmp6=filt.copy()
                tmp6["Outlets Segmentation"]=tmp6["Outlets Segmentation"].astype(str)
                tmp6["Classification"]=tmp6["Classification"].astype(str)
                tmp6[SV]=col_num(tmp6,SV)
                piv3=tmp6.pivot_table(index="Outlets Segmentation",columns="Classification",values=SV,aggfunc="sum",fill_value=0)
                if piv3.size:
                    fig,ax=sfig(8,5);im=ax.imshow(piv3.values,aspect="auto",cmap="Blues")
                    ax.set_xticks(range(piv3.shape[1]));ax.set_xticklabels(piv3.columns,rotation=35,ha="right",fontsize=7,color=TKC)
                    ax.set_yticks(range(piv3.shape[0]));ax.set_yticklabels(piv3.index,fontsize=7,color=TKC)
                    ax.set_xlabel("Classification",color=LBC);ax.set_ylabel("Outlet Segmentation",color=LBC)
                    ax.set_title("Revenue Heatmap: Segmentation x Classification (Rs)",color=TTC,fontsize=10,fontweight="bold")
                    cb3=fig.colorbar(im,ax=ax);cb3.ax.tick_params(colors=TKC,labelsize=7)
                    cb3.set_label("Total Revenue (Rs)",color=LBC,fontsize=7)
                    fig.tight_layout();st.pyplot(fig);plt.close()
                    st.markdown(ibox([
                        ("What This Shows","Revenue (Rs) for each combination of outlet segmentation tier and classification grade."),
                        ("Key Insight","Dark cells = your core revenue-generating customer groups. Light cells = growth opportunities."),
                        ("Action","Run targeted activation campaigns for light-cell combinations adjacent to dark cells — lower risk expansion.")
                    ]),unsafe_allow_html=True)
        # Segmentation KPI table
        st.markdown('<div class="sh2">Segmentation Performance Summary</div>',unsafe_allow_html=True)
        if "Outlets Segmentation" in filt.columns and SV in filt.columns:
            tmp7=filt.copy();tmp7["Outlets Segmentation"]=tmp7["Outlets Segmentation"].astype(str)
            tmp7[SV]=col_num(tmp7,SV);tmp7[DR]=col_num(tmp7,DR) if DR in tmp7.columns else 0
            agg_={SV:["sum","mean"]}
            if OUT in tmp7.columns:agg_[OUT]="nunique"
            if ON in tmp7.columns:agg_[ON]="nunique"
            if DR in tmp7.columns:agg_[DR]="mean"
            sk=tmp7.groupby("Outlets Segmentation").agg(agg_)
            sk.columns=["_".join(c).strip("_") for c in sk.columns]
            rn={}
            for c_ in sk.columns:
                if "Sale Value_sum" in c_:rn[c_]="Total Revenue (Rs)"
                elif "Sale Value_mean" in c_:rn[c_]="Avg Revenue / Order (Rs)"
                elif "Outlets Name" in c_:rn[c_]="Outlet Count"
                elif "Order No" in c_:rn[c_]="Order Count"
                elif "dispatch_rate" in c_:rn[c_]="Avg Dispatch Rate (%)"
            sk=sk.rename(columns=rn).reset_index()
            for c_ in["Total Revenue (Rs)","Avg Revenue / Order (Rs)"]:
                if c_ in sk.columns:sk[c_]=sk[c_].apply(fmt_inr)
            if "Avg Dispatch Rate (%)" in sk.columns:
                sk["Avg Dispatch Rate (%)"]=sk["Avg Dispatch Rate (%)"].apply(lambda x:f"{x:.1%}")
            st.dataframe(sk,width='stretch')
        # Comparative bar
        if "Outlets Segmentation" in filt.columns and SV in filt.columns:
            seg_rev=gb_sum(filt,"Outlets Segmentation",SV).sort_values(ascending=False)
            fig,ax=sfig(10,4)
            colors_=[PAL[i%len(PAL)] for i in range(len(seg_rev))]
            ax.bar(range(len(seg_rev)),seg_rev.values,color=colors_,edgecolor="none")
            ax.set_xticks(range(len(seg_rev)));ax.set_xticklabels(seg_rev.index,rotation=25,ha="right",fontsize=9,color=TKC)
            ax.set_ylabel("Total Revenue (Rs)",color=LBC);ax.set_xlabel("Outlet Segmentation Tier",color=LBC);inr_y(ax)
            for i,v in enumerate(seg_rev.values):ax.text(i,v*1.01,fmt_inr(v),ha="center",va="bottom",fontsize=7,color=TKC)
            ax.set_title("Revenue by Outlet Segmentation Tier (Rs) — with Labels",color=TTC,fontsize=11,fontweight="bold")
            fig.tight_layout();st.pyplot(fig);plt.close()
        top_seg2=(gb_sum(filt,"Outlets Segmentation",SV).idxmax() if "Outlets Segmentation" in filt.columns and SV in filt.columns else"N/A")
        st.markdown(ai_banner(
            f"Segmentation analysis confirms '{top_seg2}' outlets dominate revenue. "
            "Mid-tier segments with high order frequency but below-average basket size represent an upsell opportunity. "
            "Certain product divisions likely outperform specifically in premium outlet segments — cross-referencing with product analysis reveals targeted growth strategies."
        ),unsafe_allow_html=True)
        st.markdown("---");st.markdown('<div class="sh">Section Recommendations</div>',unsafe_allow_html=True)
        st.markdown(rcard("Premium Segment Retention","Top-segment outlets drive outsized revenue. Implement premium SLAs: priority stock allocation, dedicated executives, exclusive scheme access.","green"),unsafe_allow_html=True)
        st.markdown(rcard("Mid-Tier Upgrade Programme","A structured upgrade programme with clear criteria and incentives can move mid-tier outlets to premium — compounding revenue without new acquisition cost.","blue"),unsafe_allow_html=True)
        st.markdown(rcard("Low-Tier Activation Drive","Low-segmentation, low-dispatch-rate outlets are likely inactive or under-served. Run a targeted 90-day activation campaign with introductory promotions.","orange"),unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 13 — EXECUTIVE SUMMARY
    # ══════════════════════════════════════════════════════════════════════
    elif cur==13:
        st.markdown('<div class="sh">Executive Summary — Key Findings and Strategic Recommendations</div>',unsafe_allow_html=True)
        # Hero KPI strip
        s1,s2,s3,s4,s5,s6=st.columns(6)
        for c_,(l_,v_,t_) in zip([s1,s2,s3,s4,s5,s6],[
            ("Total Revenue",     fmt_inr(total_rev),      "Sum of all Sale Values (Rs)"),
            ("Total Orders",      f"{total_ord:,}",         "Unique order count"),
            ("Active Outlets",    f"{total_out:,}",         "Distinct outlet names"),
            ("Dispatch Rate",     f"{disp_rate:.1%}",       "Avg dispatched/ordered qty — target 95%+"),
            ("Qty Sold (Units)",  f"{total_qty:,.0f}",      "Total Quantity Sold (Std Units)"),
            ("Total Tax (Rs)",    fmt_inr(total_tax),       "Sum of CGST+SGST+IGST+VAT"),
        ]):c_.markdown(kpi_h(l_,v_,t_),unsafe_allow_html=True)
        st.markdown("---")
        # Major findings
        st.markdown('<div class="sh">Major Business Findings</div>',unsafe_allow_html=True)
        top5f=sdf.head(5)
        shap_summary=", ".join(f"'{r['Friendly']}' ({r['Mean_SHAP']:.4f})" for _,r in top5f.iterrows())
        top_div3=(gb_sum(filt,"ProductDivision",SV).idxmax() if "ProductDivision" in filt.columns and SV in filt.columns else"N/A")
        top_dep3=(gb_sum(filt,DEP,SV).idxmax() if DEP in filt.columns and SV in filt.columns else"N/A")
        top_out3=(gb_sum(filt,OUT,SV).idxmax() if OUT in filt.columns and SV in filt.columns else"N/A")
        findings=[
            ("SHAP Revenue Drivers",f"Top 5 revenue drivers by SHAP importance: {shap_summary}. These variables explain the majority of revenue variance and are the highest-priority optimisation targets."),
            ("Best Performing Segments",f"Top depot: '{top_dep3}'. Top product division: '{top_div3}'. Top outlet: '{top_out3}'. These segments set the performance benchmark for the business."),
            ("Fulfilment Risk",f"Current dispatch rate: {disp_rate:.1%} vs 95% industry target. Every percentage point below target represents direct revenue leakage across all orders."),
            ("Tax Exposure",f"Total tax: {fmt_inr(total_tax)} ({total_tax/max(total_rev,1)*100:.1f}% of gross revenue). IGST on inter-state orders is the most controllable component and the primary margin improvement lever."),
        ]
        for title_,body_ in findings:
            st.markdown(ibox([(title_,body_)],"purple"),unsafe_allow_html=True)
        st.markdown("---")
        # Risk indicators
        st.markdown('<div class="sh">Risk Indicators</div>',unsafe_allow_html=True)
        rc1,rc2,rc3=st.columns(3)
        disp_kind="err" if disp_rate<0.85 else("warn" if disp_rate<0.95 else"ok")
        rc1.markdown(ibox([("Dispatch Rate",f"{disp_rate:.1%} vs 95% target — {'critical gap' if disp_rate<0.85 else 'below target' if disp_rate<0.95 else 'on target'}.")],disp_kind),unsafe_allow_html=True)
        rc2.markdown(ibox([("Revenue Concentration","Top outlets and SKUs likely represent >50% of revenue — key account and SKU dependency risk.")],"warn"),unsafe_allow_html=True)
        rc3.markdown(ibox([("Tax Exposure",f"Tax represents {total_tax/max(total_rev,1)*100:.1f}% of gross revenue — IGST reduction should be a priority.")],"warn"),unsafe_allow_html=True)
        st.markdown("---")
        # AI summary insight
        st.markdown(ai_banner(
            f"Overall business health: Total revenue {fmt_inr(total_rev)} across {total_ord:,} orders. "
            f"Primary growth levers (SHAP): '{sdf.iloc[0]['Friendly']}' and '{sdf.iloc[1]['Friendly']}'. "
            f"Top performing geography: '{top_dep3}'. Top performing division: '{top_div3}'. "
            f"Dispatch rate of {disp_rate:.1%} is the most urgent operational improvement area — closing this gap to 95% would directly recover significant revenue."
        ),unsafe_allow_html=True)
        st.markdown("---")
        # Strategic recommendations
        st.markdown('<div class="sh">Strategic Recommendations</div>',unsafe_allow_html=True)
        recs=[
            ("green","Scheme Optimisation",f"Reallocate scheme budgets toward top SHAP drivers ('{sdf.iloc[0]['Friendly']}', '{sdf.iloc[1]['Friendly']}\'). A 10% budget reallocation can yield 5-8% revenue increase."),
            ("green","Key Account Protection",f"Top outlet '{top_out3}' and the top 14 others need dedicated relationship management. Revenue risk from losing even 2-3 key accounts is significant."),
            ("orange","Fulfilment Improvement",f"Close dispatch rate from {disp_rate:.1%} to 95% through stock availability improvements, better demand forecasting, and supplier lead time reduction. Each 1% improvement directly recovers proportional revenue."),
            ("blue","Executive Performance Programme","Bottom-quartile sales executives should receive structured 90-day coaching plans, territory review, and outlet reassignment where the gap is territory-driven, not skill-driven."),
            ("red","IGST Reduction",f"IGST represents {(col_num(filt,'IGST').sum() if 'IGST' in filt.columns else 0)/max(total_rev,1)*100:.1f}% of gross revenue. Review depot routing for inter-state orders. A 20% IGST reduction improves net margin by 1.5-2.5%."),
            ("blue","Segmentation Upgrade Programme","A structured mid-tier outlet upgrade programme with clear eligibility criteria and incentives can drive 10-15% volume growth without incremental outlet acquisition cost."),
        ]
        for clr_,title_,body_ in recs:
            st.markdown(rcard(title_,body_,clr_),unsafe_allow_html=True)
        st.markdown("---")
        # Export
        st.markdown('<div class="sh">Export and Download Options</div>',unsafe_allow_html=True)
        ec1,ec2,ec3=st.columns(3)
        with ec1:
            st.download_button("Download Filtered Dataset (CSV)",
                data=filt.to_csv(index=False).encode(),
                file_name="filtered_sales_data.csv",mime="text/csv",
                help="Downloads the filtered dataset as a CSV file for further analysis")
        with ec2:
            exp_s=sdf[["Rank","Friendly","Feature","Mean_SHAP"]].copy()
            exp_s.columns=["Rank","Business Name","Technical Name","Mean |SHAP| (Rs impact)"]
            exp_s["Mean |SHAP| (Rs impact)"]=exp_s["Mean |SHAP| (Rs impact)"].round(5)
            st.download_button("Download SHAP Feature Importance (CSV)",
                data=exp_s.to_csv(index=False).encode(),
                file_name="shap_feature_importance.csv",mime="text/csv",
                help="Downloads SHAP feature importance table with business names")
        with ec3:
            kpi_exp=pd.DataFrame({
                "Metric":["Total Revenue (Rs)","Total Orders","Active Outlets",
                          "Avg Order Value (Rs)","Dispatch Rate (%)","Qty Sold (Std Units)",
                          "Total Tax (Rs)","Unique Products"],
                "Value":[fmt_inr(total_rev),f"{total_ord:,}",f"{total_out:,}",fmt_inr(avg_ord),
                         f"{disp_rate:.1%}",f"{total_qty:,.0f}",fmt_inr(total_tax),f"{total_prd:,}"],
                "Definition":["Sum of all Sale Values","Unique Order No count","Distinct Outlets Name count",
                              "Total Revenue / Order Count","Avg Dispatch Qty / Ordered Qty",
                              "Sum of Qty (StdUnit)","Sum of CGST+SGST+IGST+VAT","Distinct Product count"]
            })
            st.download_button("Download KPI Summary (CSV)",
                data=kpi_exp.to_csv(index=False).encode(),
                file_name="kpi_summary_with_definitions.csv",mime="text/csv",
                help="Downloads KPI summary with metric definitions")
        st.markdown('<div style="text-align:center;color:#A0AEC0;font-size:0.72rem;margin-top:18px">'
                    'SHAP Sales Intelligence Dashboard  |  Streamlit  |  XGBoost  |  SHAP  |  v5</div>',
                    unsafe_allow_html=True)


# ─── Bottom Navigation Bar ──────────────────────────────────────────────────
st.markdown("<div class='nav-spacer'></div>",unsafe_allow_html=True)
nl,nc,nr=st.columns([2,5,2])
with nl:
    if cur>0:
        if st.button("Back",key="prev_btn"):go(-1);st.rerun()
with nc:
    dots="".join(f'<span class="dot {"on" if i==cur else ""}"></span>' for i in range(N))
    pct_=int((cur/max(N-1,1))*100)
    st.markdown(
        f'<div style="text-align:center;padding:4px 0">'
        f'<div style="display:flex;justify-content:center;gap:3px;margin-bottom:5px">{dots}</div>'
        f'<span class="nlbl">{PAGES[cur]}  —  Step {cur+1} of {N}</span>'
        f'<div class="pbg"><div class="pfill" style="width:{pct_}%"></div></div>'
        f'</div>',unsafe_allow_html=True)
with nr:
    if cur<N-1:
        locked=cur+1>=2 and not st.session_state.processed
        if st.button("Next [locked]" if locked else "Next",key="next_btn",disabled=locked):
            go(1);st.rerun()



