import io
import os
from copy import copy
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName


# ============================================================
# PERSONAL FINANCE WEB DASHBOARD
# Streamlit + Pandas + Plotly + Custom CSS
# Designed around:
# "Personal_Finance_Dashboard - Year - Month.xlsx"
# ============================================================

st.set_page_config(
    page_title="Personal Finance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------
# Theme / CSS
# ---------------------------
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Lao:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700;800&display=swap');

:root{
    --bg:#0f172a;
    --panel:#1e293b;
    --panel2:#111827;
    --teal:#00a896;
    --blue:#028090;
    --emerald:#02c39a;
    --orange:#f4a261;
    --red:#ef476f;
    --purple:#7c5cff;
    --text:#e5eef7;
    --muted:#94a3b8;
    --line:rgba(255,255,255,.10);
}

html, body, [class*="css"] {
    font-family: 'Inter','Noto Sans Lao',sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at 10% 0%, rgba(0,168,150,.16), transparent 28%),
        radial-gradient(circle at 90% 5%, rgba(2,128,144,.18), transparent 30%),
        linear-gradient(135deg,#0f172a 0%,#111827 48%,#0b1220 100%);
    color:var(--text);
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#0b1220,#111827);
    border-right:1px solid var(--line);
}

[data-testid="stSidebar"] * { color:var(--text); }

.block-container { padding-top:1.2rem; padding-bottom:3rem; }

.hero {
    padding: 26px 30px;
    border-radius: 24px;
    border:1px solid rgba(2,195,154,.28);
    background:linear-gradient(135deg,rgba(30,41,59,.82),rgba(15,23,42,.68));
    box-shadow:0 18px 50px rgba(0,0,0,.28), inset 0 1px 0 rgba(255,255,255,.05);
    backdrop-filter:blur(18px);
    margin-bottom:20px;
}

.hero h1 { margin:0; font-size:2rem; font-weight:800; letter-spacing:-.03em; }
.hero p { margin:.35rem 0 0; color:var(--muted); }

.kpi {
    padding:18px 18px 16px;
    min-height:145px;
    border-radius:20px;
    border:1px solid rgba(255,255,255,.10);
    background:linear-gradient(145deg,rgba(30,41,59,.88),rgba(15,23,42,.72));
    box-shadow:0 12px 34px rgba(0,0,0,.22), inset 0 1px 0 rgba(255,255,255,.04);
    backdrop-filter:blur(16px);
    position:relative;
    overflow:hidden;
}
.kpi:before {
    content:"";
    position:absolute; left:0; top:0; width:4px; height:100%;
    background:var(--accent,#00a896);
    box-shadow:0 0 20px var(--accent,#00a896);
}
.kpi .label { color:#cbd5e1; font-size:.82rem; font-weight:600; }
.kpi .value { font-size:1.45rem; font-weight:800; margin-top:8px; }
.kpi .sub { color:var(--muted); font-size:.76rem; margin-top:5px; }
.badge {
    display:inline-block; padding:5px 10px; border-radius:999px;
    font-size:.72rem; font-weight:800; margin-top:9px;
    border:1px solid rgba(255,255,255,.12);
}
.good { color:#8ff3d7; background:rgba(2,195,154,.12); }
.warn { color:#ffd19b; background:rgba(244,162,97,.12); }
.bad { color:#ff9db4; background:rgba(239,71,111,.12); }
.info { color:#9bdcf0; background:rgba(2,128,144,.12); }

.section {
    margin:20px 0 10px;
    font-size:1.08rem; font-weight:800;
    letter-spacing:-.01em;
}
.section small { color:var(--muted); font-weight:500; }

.panel {
    padding:18px;
    border-radius:20px;
    border:1px solid rgba(255,255,255,.09);
    background:rgba(30,41,59,.60);
    box-shadow:0 12px 34px rgba(0,0,0,.20);
    backdrop-filter:blur(16px);
}

.progress-wrap { margin:10px 0 8px; }
.progress-bg {
    width:100%; height:14px; border-radius:999px;
    background:rgba(255,255,255,.08); overflow:hidden;
}
.progress-fill {
    height:100%; border-radius:999px;
    background:linear-gradient(90deg,#00a896,#02c39a);
    box-shadow:0 0 18px rgba(2,195,154,.35);
}

.ratio-box {
    padding:14px; border-radius:16px;
    background:rgba(15,23,42,.55);
    border:1px solid rgba(255,255,255,.07);
}
.big-score {
    font-size:3.2rem; line-height:1; font-weight:900;
    background:linear-gradient(90deg,#00a896,#02c39a,#028090);
    -webkit-background-clip:text; -webkit-text-fill-color:transparent;
}

div[data-testid="stMetric"] {
    background:rgba(30,41,59,.62);
    border:1px solid rgba(255,255,255,.08);
    padding:14px;
    border-radius:16px;
}

.stTabs [data-baseweb="tab-list"] {
    gap:6px;
    background:rgba(15,23,42,.58);
    padding:7px;
    border-radius:16px;
}
.stTabs [data-baseweb="tab"] {
    border-radius:12px;
    padding:9px 15px;
}
.stTabs [aria-selected="true"] {
    background:rgba(0,168,150,.18);
}

div[data-testid="stDataEditor"] {
    border-radius:16px;
    overflow:hidden;
}

button[kind="primary"] {
    background:linear-gradient(90deg,#00a896,#028090);
    border:0;
}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# Helpers
# ============================================================

MONTH_NAMES_LO = [
    "ມັງກອນ", "ກຸມພາ", "ມີນາ", "ເມສາ", "ພຶດສະພາ", "ມິຖຸນາ",
    "ກໍລະກົດ", "ສິງຫາ", "ກັນຍາ", "ຕຸລາ", "ພະຈິກ", "ທັນວາ"
]
MONTH_NAMES_EN = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

def fmt_money(x):
    try:
        if pd.isna(x):
            return "₭ 0"
        return f"₭ {float(x):,.0f}"
    except Exception:
        return "₭ 0"

def fmt_pct(x):
    try:
        return f"{float(x)*100:.1f}%"
    except Exception:
        return "0.0%"

def clean_numeric(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)

def safe_date_series(s):
    return pd.to_datetime(s, errors="coerce")

def status_badge(status):
    text = str(status or "")
    if "GOOD" in text or "COMPLETE" in text or "Healthy" in text:
        cls = "good"
    elif "WARNING" in text or "BUILD" in text or "MEDIUM" in text:
        cls = "warn"
    elif "ATTENTION" in text or "NEED" in text or "OVERDUE" in text or "HIGH" in text:
        cls = "bad"
    else:
        cls = "info"
    return f'<span class="badge {cls}">{text}</span>'

def score_status(score):
    if score >= 80:
        return "🟢 GOOD"
    if score >= 60:
        return "🟡 WARNING"
    return "🔴 NEED ATTENTION"

def change_pct(current, previous):
    if previous is None or pd.isna(previous) or previous == 0:
        return np.nan
    return (current - previous) / previous

def month_label(m):
    if 1 <= int(m) <= 12:
        return f"{int(m):02d} · {MONTH_NAMES_LO[int(m)-1]}"
    return str(m)

def styled_fig(fig, height=330):
    fig.update_layout(
        template="plotly_dark",
        height=height,
        margin=dict(l=10,r=10,t=35,b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#dbeafe", family="Inter"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        hoverlabel=dict(bgcolor="#0f172a"),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="rgba(255,255,255,.07)")
    return fig


# ============================================================
# Excel load / parse
# ============================================================

@st.cache_data(show_spinner=False)
def read_excel_bytes(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    return {s: pd.read_excel(xls, sheet_name=s, header=None) for s in xls.sheet_names}

def nonempty_rows(df):
    if df.empty:
        return df.copy()
    return df.dropna(how="all").copy()

def parse_transactions(raw):
    cols = ["Date","Type","Category","Description","Account","Need_Want","Amount","Month","Year","Note"]
    if raw.shape[1] < 10:
        raw = raw.reindex(columns=range(10))
    df = raw.iloc[2:, :10].copy()
    df.columns = cols
    df = nonempty_rows(df)
    # Keep rows that have any meaningful transaction field.
    mask = df[["Date","Type","Category","Description","Account","Need_Want","Amount","Note"]].notna().any(axis=1)
    df = df.loc[mask].reset_index(drop=True)
    df["Date"] = safe_date_series(df["Date"])
    df["Amount"] = clean_numeric(df["Amount"])
    df["Month"] = df["Date"].dt.month.astype("Int64")
    df["Year"] = df["Date"].dt.year.astype("Int64")
    return df

def parse_debt(raw):
    cols = ["Debt ID","Debt Name","Lender","Original Principal","Principal Paid",
            "Remaining Principal","Interest Rate","Monthly Payment","Due Date",
            "Start Date","End Date","Status","Notes"]
    raw = raw.reindex(columns=range(max(13, raw.shape[1])))
    df = raw.iloc[2:, :13].copy()
    df.columns = cols
    df = nonempty_rows(df)
    mask = df[["Debt ID","Debt Name","Lender","Original Principal","Principal Paid",
               "Interest Rate","Monthly Payment","Due Date","Start Date","End Date","Notes"]].notna().any(axis=1)
    df = df.loc[mask].reset_index(drop=True)
    for c in ["Original Principal","Principal Paid","Monthly Payment","Interest Rate"]:
        df[c] = clean_numeric(df[c])
    for c in ["Due Date","Start Date","End Date"]:
        df[c] = safe_date_series(df[c])
    df["Remaining Principal"] = (df["Original Principal"] - df["Principal Paid"]).clip(lower=0)
    today = pd.Timestamp.today().normalize()
    def debt_status(r):
        if not r["Debt ID"] and not r["Debt Name"]:
            return ""
        if r["Remaining Principal"] <= 0:
            return "✅ PAID"
        if pd.notna(r["Due Date"]) and r["Due Date"] < today:
            return "⚠ OVERDUE"
        return "🔵 ACTIVE"
    df["Status"] = df.apply(debt_status, axis=1)
    return df

def parse_assets(raw):
    cols = ["Asset ID","Asset Name","Category","Purchase Date","Purchase Cost",
            "Current Value","Account","Liquidity","Notes"]
    raw = raw.reindex(columns=range(max(9, raw.shape[1])))
    df = raw.iloc[2:, :9].copy()
    df.columns = cols
    df = nonempty_rows(df)
    mask = df[["Asset ID","Asset Name","Category","Purchase Date","Purchase Cost",
               "Current Value","Account","Liquidity","Notes"]].notna().any(axis=1)
    df = df.loc[mask].reset_index(drop=True)
    for c in ["Purchase Cost","Current Value"]:
        df[c] = clean_numeric(df[c])
    df["Purchase Date"] = safe_date_series(df["Purchase Date"])
    return df

def parse_emergency(raw):
    # A3:G3 in the workbook
    vals = raw.iloc[2, :7].tolist() if raw.shape[0] >= 3 else [0,3,0,0,0,0,""]
    return {
        "Monthly Essential Expense": float(pd.to_numeric(vals[0], errors="coerce") or 0),
        "Target Months": float(pd.to_numeric(vals[1], errors="coerce") or 3),
        "Current Fund": float(pd.to_numeric(vals[3], errors="coerce") or 0),
    }

def parse_saving_log(raw):
    # Allocation log is A13:G...
    if raw.shape[0] <= 13:
        return pd.DataFrame(columns=["Date","Type","Category","Account/Platform","Amount","Goal","Notes"])
    cols = ["Date","Type","Category","Account/Platform","Amount","Goal","Notes"]
    x = raw.reindex(columns=range(7)).iloc[13:, :7].copy()
    x.columns = cols
    x = nonempty_rows(x)
    if x.empty:
        return pd.DataFrame(columns=cols)
    x["Date"] = safe_date_series(x["Date"])
    x["Amount"] = clean_numeric(x["Amount"])
    return x.reset_index(drop=True)

SETTINGS_MAP = {
    "Years": ("A", 0, int),
    "Months #": ("B", 0, int),
    "Month Names": ("C", 0, str),
    "Income Categories": ("D", 0, str),
    "Expense Categories": ("E", 0, str),
    "Accounts": ("F", 0, str),
    "Need / Want": ("G", 0, str),
    "Debt Status": ("H", 0, str),
    "Asset Categories": ("I", 0, str),
    "Investment Categories": ("J", 0, str),
    "EF Target Months": ("K", 0, float),
    "Transaction Types": ("L", 0, str),
    "Liquidity": ("M", 0, str),
    "Saving Categories": ("N", 0, str),
    "Saving / Investment Types": ("O", 0, str),
}

def parse_settings(raw):
    settings = {}
    for name, (letter, _, caster) in SETTINGS_MAP.items():
        col = ord(letter) - ord("A")
        vals = raw.iloc[2:, col].tolist() if raw.shape[1] > col else []
        clean = []
        for v in vals:
            if pd.isna(v) or str(v).strip() == "":
                continue
            try:
                if caster is int:
                    clean.append(int(float(v)))
                elif caster is float:
                    clean.append(float(v))
                else:
                    clean.append(str(v))
            except Exception:
                clean.append(str(v))
        settings[name] = clean
    return settings


# ============================================================
# Session state
# ============================================================

def init_state(file_bytes):
    if st.session_state.get("loaded_signature") == hash(file_bytes):
        return
    sheets = read_excel_bytes(file_bytes)
    required = [
        "01_Dashboard","02_Income_Expense","03_Cash_Flow","04_Emergency_Fund",
        "05_Debt","06_Assets","07_Net_Worth","08_Saving_Investment",
        "09_Monthly_Report","10_Settings"
    ]
    missing = [x for x in required if x not in sheets]
    if missing:
        st.error("Missing sheets: " + ", ".join(missing))
        st.stop()

    st.session_state.loaded_signature = hash(file_bytes)
    st.session_state.original_bytes = file_bytes
    st.session_state.raw_sheets = sheets
    st.session_state.transactions = parse_transactions(sheets["02_Income_Expense"])
    st.session_state.debts = parse_debt(sheets["05_Debt"])
    st.session_state.assets = parse_assets(sheets["06_Assets"])
    st.session_state.emergency = parse_emergency(sheets["04_Emergency_Fund"])
    st.session_state.saving_log = parse_saving_log(sheets["08_Saving_Investment"])
    st.session_state.settings = parse_settings(sheets["10_Settings"])

    years = st.session_state.settings.get("Years", [])
    tx_years = st.session_state.transactions["Year"].dropna().astype(int).unique().tolist()
    all_years = sorted(set(years + tx_years))
    if not all_years:
        all_years = [datetime.now().year]
    st.session_state.selected_year = all_years[-1]
    st.session_state.selected_month = datetime.now().month if datetime.now().month in range(1,13) else 1


# ============================================================
# Financial engine
# ============================================================

def calc_metrics(tx, debts, assets, emergency, year, month):
    t = tx.copy()
    if not t.empty:
        t["Date"] = safe_date_series(t["Date"])
        t["Amount"] = clean_numeric(t["Amount"])
        t["Month"] = t["Date"].dt.month
        t["Year"] = t["Date"].dt.year

    selected = t[(t["Year"] == year) & (t["Month"] == month)] if not t.empty else t
    previous_month = 12 if month == 1 else month - 1
    previous_year = year - 1 if month == 1 else year
    previous = t[(t["Year"] == previous_year) & (t["Month"] == previous_month)] if not t.empty else t

    def total(frame, typ):
        return float(frame.loc[frame["Type"].astype(str).str.lower() == typ.lower(), "Amount"].sum())

    income = total(selected, "Income")
    expense = total(selected, "Expense")
    prev_income = total(previous, "Income")
    prev_expense = total(previous, "Expense")
    cash_flow = income - expense
    prev_cash_flow = prev_income - prev_expense

    saving = float(selected.loc[
        (selected["Type"] == "Expense") & (selected["Need_Want"] == "Saving"), "Amount"
    ].sum())
    investment = float(selected.loc[
        (selected["Type"] == "Expense") & (selected["Need_Want"] == "Investment"), "Amount"
    ].sum())

    all_saving = float(t.loc[(t["Type"] == "Expense") & (t["Need_Want"] == "Saving"), "Amount"].sum())
    all_investment = float(t.loc[(t["Type"] == "Expense") & (t["Need_Want"] == "Investment"), "Amount"].sum())

    debt_remaining = float(clean_numeric(debts["Remaining Principal"]).sum()) if not debts.empty else 0
    debt_monthly = float(clean_numeric(debts["Monthly Payment"]).sum()) if not debts.empty else 0
    debt_original = float(clean_numeric(debts["Original Principal"]).sum()) if not debts.empty else 0
    debt_paid = float(clean_numeric(debts["Principal Paid"]).sum()) if not debts.empty else 0

    assets_total = float(clean_numeric(assets["Current Value"]).sum()) if not assets.empty else 0
    ef_fund = float(emergency.get("Current Fund", 0))
    ef_target = float(emergency.get("Monthly Essential Expense", 0) * emergency.get("Target Months", 0))
    ef_progress = ef_fund / ef_target if ef_target > 0 else 0
    ef_months = ef_fund / emergency.get("Monthly Essential Expense", 1) if emergency.get("Monthly Essential Expense", 0) > 0 else 0

    net_worth = assets_total - debt_remaining

    # Need/Want
    need = float(selected.loc[selected["Need_Want"] == "Need", "Amount"].sum())
    want = float(selected.loc[selected["Need_Want"] == "Want", "Amount"].sum())

    # Ratios
    saving_rate = saving / income if income else 0
    investment_rate = investment / income if income else 0
    debt_to_income = debt_monthly / income if income else 0

    # Score: intentionally transparent and conservative.
    cash_score = 20 if cash_flow >= 0 else max(0, 20 + (cash_flow / max(income,1))*20)
    saving_score = min(20, max(0, saving_rate / 0.20 * 20))
    ef_score = min(20, max(0, ef_progress * 20))
    debt_score = 20 if debt_to_income <= 0.20 else max(0, 20 - (debt_to_income-0.20)*40)
    invest_score = min(10, max(0, investment_rate / 0.10 * 10))

    # Net worth trend based on monthly series.
    monthly = monthly_metrics(t, debts, assets, year)
    nw_trend_score = 10
    if len(monthly) >= 2:
        first = monthly.iloc[0]["Net Worth"]
        last = monthly.iloc[-1]["Net Worth"]
        if last > first:
            nw_trend_score = 10
        elif last == first:
            nw_trend_score = 6
        else:
            nw_trend_score = 2

    score = int(round(min(100, max(0, cash_score + saving_score + ef_score + debt_score + invest_score + nw_trend_score))))

    if ef_progress >= 1:
        ef_status = "🟢 GOOD / COMPLETE"
    elif ef_progress >= .5:
        ef_status = "🟡 WARNING / BUILDING"
    else:
        ef_status = "🔴 NEED ATTENTION / BUILD FUND"

    return {
        "income": income, "expense": expense, "cash_flow": cash_flow,
        "prev_income": prev_income, "prev_expense": prev_expense, "prev_cash_flow": prev_cash_flow,
        "saving": saving, "investment": investment,
        "all_saving": all_saving, "all_investment": all_investment,
        "assets": assets_total, "debt": debt_remaining, "debt_monthly": debt_monthly,
        "debt_original": debt_original, "debt_paid": debt_paid,
        "net_worth": net_worth, "ef_fund": ef_fund, "ef_target": ef_target,
        "ef_progress": ef_progress, "ef_months": ef_months,
        "need": need, "want": want,
        "saving_rate": saving_rate, "investment_rate": investment_rate,
        "debt_to_income": debt_to_income, "score": score,
        "score_status": score_status(score), "ef_status": ef_status,
    }

def monthly_metrics(tx, debts, assets, year):
    rows = []
    t = tx.copy()
    if t.empty:
        t = pd.DataFrame(columns=["Date","Type","Category","Need_Want","Amount"])
    t["Date"] = safe_date_series(t.get("Date", pd.Series(dtype="datetime64[ns]")))
    t["Amount"] = clean_numeric(t.get("Amount", pd.Series(dtype=float)))
    t["Month"] = t["Date"].dt.month
    t["Year"] = t["Date"].dt.year

    debt_total = float(clean_numeric(debts["Remaining Principal"]).sum()) if not debts.empty else 0
    asset_total = float(clean_numeric(assets["Current Value"]).sum()) if not assets.empty else 0
    nw = asset_total - debt_total

    # The original workbook uses current asset/debt balances for the selected month.
    # The web app therefore uses the same balance-sheet snapshot for the 12-month view
    # unless historical snapshots are present.
    for m in range(1,13):
        x = t[(t["Year"] == year) & (t["Month"] == m)]
        inc = float(x.loc[x["Type"] == "Income","Amount"].sum())
        exp = float(x.loc[x["Type"] == "Expense","Amount"].sum())
        sav = float(x.loc[(x["Type"] == "Expense") & (x["Need_Want"] == "Saving"),"Amount"].sum())
        inv = float(x.loc[(x["Type"] == "Expense") & (x["Need_Want"] == "Investment"),"Amount"].sum())
        rows.append({
            "Month": m,
            "Month Name": MONTH_NAMES_LO[m-1],
            "Income": inc,
            "Expense": exp,
            "Cash Flow": inc-exp,
            "Saving": sav,
            "Investment": inv,
            "Debt": debt_total,
            "Assets": asset_total,
            "Net Worth": nw,
            "Saving Rate": sav/inc if inc else 0,
            "Investment Rate": inv/inc if inc else 0,
        })
    return pd.DataFrame(rows)


# ============================================================
# Excel export while preserving the original workbook
# ============================================================

def set_defined_name(wb, name, sheet, cell_range):
    ref = f"'{sheet}'!{cell_range}"
    if name in wb.defined_names:
        wb.defined_names[name].attr_text = ref
    else:
        wb.defined_names.add(DefinedName(name, attr_text=ref))

def copy_style_row(ws, source_row, target_row, max_col):
    if source_row > ws.max_row:
        return
    for c in range(1, max_col+1):
        src = ws.cell(source_row,c)
        dst = ws.cell(target_row,c)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)

def clear_rows(ws, start_row, end_row, max_col):
    for r in range(start_row, end_row+1):
        for c in range(1,max_col+1):
            ws.cell(r,c).value = None

def write_df_rows(ws, df, start_row, max_col):
    old_end = max(ws.max_row, start_row + len(df) + 10)
    clear_rows(ws, start_row, old_end, max_col)
    for i, (_, row) in enumerate(df.iterrows(), start=start_row):
        copy_style_row(ws, start_row, i, max_col)
        for c, value in enumerate(row.tolist(), start=1):
            ws.cell(i,c).value = None if pd.isna(value) else value

def prepare_export():
    wb = load_workbook(io.BytesIO(st.session_state.original_bytes))

    # ---------------- Transactions ----------------
    tx = st.session_state.transactions.copy()
    tx["Date"] = safe_date_series(tx["Date"])
    tx["Amount"] = clean_numeric(tx["Amount"])
    tx["Month"] = tx["Date"].dt.month
    tx["Year"] = tx["Date"].dt.year
    ws = wb["02_Income_Expense"]
    clear_rows(ws, 3, max(ws.max_row, 5000), 10)
    for i, (_, r) in enumerate(tx.iterrows(), start=3):
        copy_style_row(ws, 3, i, 10)
        vals = [
            r["Date"], r["Type"], r["Category"], r["Description"], r["Account"],
            r["Need_Want"], float(r["Amount"]) if pd.notna(r["Amount"]) else None,
            f'=IF(A{i}="","",MONTH(A{i}))',
            f'=IF(A{i}="","",YEAR(A{i}))',
            r["Note"]
        ]
        for c,v in enumerate(vals,1):
            ws.cell(i,c).value = None if pd.isna(v) else v

    # Expand transaction named ranges for future Excel entries.
    set_defined_name(wb, "TxnDate", "02_Income_Expense", "$A$3:$A$5000")
    set_defined_name(wb, "TxnType", "02_Income_Expense", "$B$3:$B$5000")
    set_defined_name(wb, "TxnCat", "02_Income_Expense", "$C$3:$C$5000")
    set_defined_name(wb, "TxnAcc", "02_Income_Expense", "$E$3:$E$5000")
    set_defined_name(wb, "TxnNW", "02_Income_Expense", "$F$3:$F$5000")
    set_defined_name(wb, "TxnAmount", "02_Income_Expense", "$G$3:$G$5000")
    set_defined_name(wb, "TxnMonth", "02_Income_Expense", "$H$3:$H$5000")
    set_defined_name(wb, "TxnYear", "02_Income_Expense", "$I$3:$I$5000")

    # Expand data validations to match future rows.
    ws.data_validations.dataValidation = [
        dv for dv in ws.data_validations.dataValidation
        if not any(str(col) in str(dv.sqref) for col in ["B","C","E","F"])
    ]
    for formula, rng in [
        ("=TypeList","B3:B5000"),
        ("=AllCategories","C3:C5000"),
        ("=AccountsList","E3:E5000"),
        ("=NeedWantList","F3:F5000"),
    ]:
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        dv.add(rng)
        ws.add_data_validation(dv)

    # ---------------- Emergency Fund ----------------
    ef = st.session_state.emergency
    ws = wb["04_Emergency_Fund"]
    ws["A3"] = ef["Monthly Essential Expense"]
    ws["B3"] = ef["Target Months"]
    ws["C3"] = "=A3*B3"
    ws["D3"] = ef["Current Fund"]
    ws["E3"] = "=MAX(C3-D3,0)"
    ws["F3"] = '=IFERROR(D3/C3,0)'
    ws["G3"] = '=IF(C3=0,"—",IF(F3>=1,"🟢 COMPLETE / ສຳເລັດ",IF(F3>=0.75,"🟢 GOOD / ດີ",IF(F3>=0.5,"🟡 BUILDING / ກຳລັງສະສົມ","🔴 NEED TO BUILD / ຕ້ອງເລັ່ງສະສົມ"))))'

    # ---------------- Debt ----------------
    debt = st.session_state.debts.copy()
    ws = wb["05_Debt"]
    clear_rows(ws, 3, max(ws.max_row, 200), 13)
    for i, (_,r) in enumerate(debt.iterrows(), start=3):
        copy_style_row(ws, 3, i, 13)
        vals = [
            r["Debt ID"], r["Debt Name"], r["Lender"],
            float(r["Original Principal"]) if pd.notna(r["Original Principal"]) else None,
            float(r["Principal Paid"]) if pd.notna(r["Principal Paid"]) else None,
            f'=IF(D{i}="","",MAX(D{i}-E{i},0))',
            float(r["Interest Rate"]) if pd.notna(r["Interest Rate"]) else None,
            float(r["Monthly Payment"]) if pd.notna(r["Monthly Payment"]) else None,
            r["Due Date"], r["Start Date"], r["End Date"],
            f'=IF(D{i}="","",IF(F{i}<=0,"✅ PAID / ຈ່າຍໝົດ",IF(AND(I{i}<>"",I{i}<TODAY()),"⚠ OVERDUE / ເກີນກຳນົດ","🔵 ACTIVE / ດຳເນີນຢູ່")))',
            r["Notes"]
        ]
        for c,v in enumerate(vals,1):
            ws.cell(i,c).value = None if pd.isna(v) else v
    end = max(33, 2 + len(debt))
    ws["B37"] = f"=SUM(F3:F{end})"
    ws["B38"] = f"=SUM(H3:H{end})"
    ws["B41"] = f'=IFERROR(SUM(E3:E{end})/SUM(D3:D{end}),0)'

    # ---------------- Assets ----------------
    assets = st.session_state.assets.copy()
    ws = wb["06_Assets"]
    clear_rows(ws, 3, max(ws.max_row, 500), 9)
    for i,(_,r) in enumerate(assets.iterrows(), start=3):
        copy_style_row(ws, 3, i, 9)
        vals = [r["Asset ID"],r["Asset Name"],r["Category"],r["Purchase Date"],
                float(r["Purchase Cost"]) if pd.notna(r["Purchase Cost"]) else None,
                float(r["Current Value"]) if pd.notna(r["Current Value"]) else None,
                r["Account"],r["Liquidity"],r["Notes"]]
        for c,v in enumerate(vals,1):
            ws.cell(i,c).value = None if pd.isna(v) else v
    asset_end = max(48, 2 + len(assets))
    ws["B60"] = f"=SUM(F3:F{asset_end})"
    # Preserve the original category summary but extend it to all entered rows.
    for r in range(54,59):
        ws.cell(r,2).value = f'=SUMIFS($F$3:$F${asset_end},$C$3:$C${asset_end},A{r})'

    set_defined_name(wb, "AssetTotal", "06_Assets", "$B$60")
    set_defined_name(wb, "AssetEmergencyFund", "06_Assets", "$B$54")

    # ---------------- Saving / Investment allocation log ----------------
    log = st.session_state.saving_log.copy()
    ws = wb["08_Saving_Investment"]
    clear_rows(ws, 14, max(ws.max_row, 300), 7)
    for i,(_,r) in enumerate(log.iterrows(), start=14):
        copy_style_row(ws, 14, i, 7)
        vals = [r["Date"],r["Type"],r["Category"],r["Account/Platform"],
                float(r["Amount"]) if pd.notna(r["Amount"]) else None,r["Goal"],r["Notes"]]
        for c,v in enumerate(vals,1):
            ws.cell(i,c).value = None if pd.isna(v) else v

    # ---------------- Settings ----------------
    settings = st.session_state.settings
    ws = wb["10_Settings"]
    col_letters = {name:meta[0] for name,meta in SETTINGS_MAP.items()}
    for name, letter in col_letters.items():
        col = ord(letter)-ord("A")+1
        # Keep enough room for future entries.
        for r in range(3,101):
            ws.cell(r,col).value = None
        for idx,val in enumerate(settings.get(name,[]), start=3):
            ws.cell(idx,col).value = val

    # All Categories is an auto-maintained helper column Q.
    all_cats = list(dict.fromkeys(
        [str(x) for x in settings.get("Income Categories",[]) if str(x).strip()] +
        [str(x) for x in settings.get("Expense Categories",[]) if str(x).strip()]
    ))
    for r in range(3,101):
        ws.cell(r,17).value = None
    for i,val in enumerate(all_cats,start=3):
        ws.cell(i,17).value = val

    # Update named ranges without deleting existing names.
    name_to_col = {
        "YearsList":"A","MonthsList":"B","MonthNamesList":"C",
        "IncomeCategories":"D","ExpenseCategories":"E","AccountsList":"F",
        "NeedWantList":"G","DebtStatusList":"H","AssetCategories":"I",
        "InvestmentCategories":"J","EFTargetMonthsList":"K","TypeList":"L",
        "LiquidityList":"M","SavingCategoryList":"N","SavingInvestTypeList":"O",
        "AllCategories":"Q",
    }
    for nm,letter in name_to_col.items():
        vals = settings.get({
            "YearsList":"Years","MonthsList":"Months #","MonthNamesList":"Month Names",
            "IncomeCategories":"Income Categories","ExpenseCategories":"Expense Categories",
            "AccountsList":"Accounts","NeedWantList":"Need / Want","DebtStatusList":"Debt Status",
            "AssetCategories":"Asset Categories","InvestmentCategories":"Investment Categories",
            "EFTargetMonthsList":"EF Target Months","TypeList":"Transaction Types",
            "LiquidityList":"Liquidity","SavingCategoryList":"Saving Categories",
            "SavingInvestTypeList":"Saving / Investment Types"
        }.get(nm,""), [])
        if nm == "AllCategories":
            vals = all_cats
        n = max(1,len(vals))
        set_defined_name(wb,nm,"10_Settings",f"${letter}$3:${letter}${2+n}")

    # Calculation settings: force Excel to recalculate formulas when opened.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ============================================================
# File selection
# ============================================================

st.sidebar.markdown("## 📊 Personal Finance")
st.sidebar.caption("Modern Dark Glassmorphic Edition")

uploaded = st.sidebar.file_uploader(
    "📁 Upload Excel Workbook",
    type=["xlsx"],
    help="Upload Personal_Finance_Dashboard - Year - Month.xlsx",
)

default_candidates = [
    "Personal_Finance_Dashboard - Year - Month.xlsx",
    "/mnt/data/Personal_Finance_Dashboard - Year - Month.xlsx",
]

file_bytes = None
if uploaded is not None:
    file_bytes = uploaded.getvalue()
else:
    for p in default_candidates:
        if os.path.exists(p):
            with open(p,"rb") as f:
                file_bytes = f.read()
            break

if file_bytes is None:
    st.markdown(
        """
        <div class="hero">
            <h1>📊 PERSONAL FINANCE DASHBOARD</h1>
            <p>Upload your Excel workbook from the left sidebar to start.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.info("Required workbook: Personal_Finance_Dashboard - Year - Month.xlsx")
    st.stop()

init_state(file_bytes)


# ============================================================
# Sidebar selectors
# ============================================================

settings = st.session_state.settings
tx = st.session_state.transactions
debts = st.session_state.debts
assets = st.session_state.assets
emergency = st.session_state.emergency

years = sorted(set(settings.get("Years",[]) + tx["Year"].dropna().astype(int).unique().tolist()))
if not years:
    years = [datetime.now().year]

year_index = years.index(st.session_state.get("selected_year", years[-1])) if st.session_state.get("selected_year", years[-1]) in years else len(years)-1
selected_year = st.sidebar.selectbox("📅 Year", years, index=year_index)
selected_month = st.sidebar.selectbox(
    "🗓️ Month",
    list(range(1,13)),
    index=int(st.session_state.get("selected_month", datetime.now().month))-1,
    format_func=month_label,
)
st.session_state.selected_year = selected_year
st.session_state.selected_month = selected_month

st.sidebar.markdown("---")
st.sidebar.markdown("### 💾 Export")
try:
    export_bytes = prepare_export()
    st.sidebar.download_button(
        "💾 Save & Download Excel",
        data=export_bytes,
        file_name="Personal_Finance_Dashboard_Web_Updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
except Exception as e:
    st.sidebar.error(f"Export error: {e}")

st.sidebar.markdown("---")
st.sidebar.caption("Currency: LAK (₭)")
st.sidebar.caption("Edits are recalculated on each Streamlit interaction.")


# ============================================================
# Current metrics
# ============================================================

metrics = calc_metrics(
    st.session_state.transactions,
    st.session_state.debts,
    st.session_state.assets,
    st.session_state.emergency,
    selected_year,
    selected_month,
)


# ============================================================
# Header
# ============================================================

st.markdown(
    f"""
<div class="hero">
    <div style="display:flex;justify-content:space-between;gap:20px;align-items:center;flex-wrap:wrap;">
        <div>
            <h1>📊 PERSONAL FINANCE DASHBOARD</h1>
            <p>ຕິດຕາມສະຖານະການເງິນສ່ວນບຸກຄົນ · {month_label(selected_month)} {selected_year} · LAK (₭)</p>
        </div>
        <div style="text-align:right">
            <div style="font-size:.78rem;color:#94a3b8">CURRENT NET WORTH</div>
            <div style="font-size:1.8rem;font-weight:900">{fmt_money(metrics["net_worth"])}</div>
            <div>{status_badge(metrics["score_status"])} <span style="margin-left:8px;color:#94a3b8">Score</span> <b>{metrics["score"]}/100</b></div>
        </div>
    </div>
</div>
""",
    unsafe_allow_html=True,
)


# ============================================================
# Six main tabs
# ============================================================

tabs = st.tabs([
    "📊 Executive Overview",
    "💵 Transactions Log",
    "💳 Debt & Assets",
    "🛟 Emergency & Investment",
    "📅 Monthly Report",
    "⚙️ System Settings",
])


# ============================================================
# TAB 1 — Executive Overview
# ============================================================
with tabs[0]:
    kpis = [
        ("💰","Total Income",metrics["income"],metrics["prev_income"],"#02c39a"),
        ("💸","Total Expense",metrics["expense"],metrics["prev_expense"],"#f4a261"),
        ("💵","Net Cash Flow",metrics["cash_flow"],metrics["prev_cash_flow"],"#028090"),
        ("🛟","Emergency Fund",metrics["ef_fund"],None,"#00a896"),
        ("💳","Total Debt",metrics["debt"],None,"#ef476f"),
        ("🏠","Total Assets",metrics["assets"],None,"#02c39a"),
        ("💎","Net Worth",metrics["net_worth"],None,"#7c5cff"),
        ("📈","Saving + Investment",metrics["saving"]+metrics["investment"],None,"#f4a261"),
    ]
    cols = st.columns(4)
    for i,(icon,label,value,prev,accent) in enumerate(kpis):
        ch = change_pct(value,prev)
        ch_text = f"MoM {fmt_pct(ch)}" if pd.notna(ch) else "Current selected month"
        cols[i%4].markdown(
            f"""
<div class="kpi" style="--accent:{accent}">
    <div class="label">{icon} {label}</div>
    <div class="value">{fmt_money(value)}</div>
    <div class="sub">{ch_text}</div>
</div>
""",
            unsafe_allow_html=True,
        )

    st.markdown('<div class="section">📊 Financial Performance <small>12-month view</small></div>',unsafe_allow_html=True)
    monthly = monthly_metrics(st.session_state.transactions, debts, assets, selected_year)

    c1,c2 = st.columns(2)
    with c1:
        fig = go.Figure()
        fig.add_bar(x=[month_label(x) for x in monthly["Month"]], y=monthly["Income"], name="Income")
        fig.add_bar(x=[month_label(x) for x in monthly["Month"]], y=monthly["Expense"], name="Expense")
        fig.update_layout(barmode="group", title="Income vs Expense")
        st.plotly_chart(styled_fig(fig),use_container_width=True)
    with c2:
        fig = px.line(monthly, x="Month Name", y="Cash Flow", markers=True, title="Cash Flow Trend")
        fig.add_hline(y=0,line_dash="dot")
        st.plotly_chart(styled_fig(fig),use_container_width=True)

    c1,c2,c3 = st.columns(3)
    with c1:
        st.markdown('<div class="panel">',unsafe_allow_html=True)
        st.markdown("### 🛡️ Emergency Fund")
        p=min(100,max(0,metrics["ef_progress"]*100))
        st.markdown(
            f"""
            <div style="display:flex;justify-content:space-between">
                <span>{fmt_money(metrics["ef_fund"])}</span><b>{p:.1f}%</b>
            </div>
            <div class="progress-wrap"><div class="progress-bg"><div class="progress-fill" style="width:{p}%"></div></div></div>
            <div style="color:#94a3b8;font-size:.82rem">Target: {fmt_money(metrics["ef_target"])} · {metrics["ef_months"]:.1f} months covered</div>
            <div>{status_badge(metrics["ef_status"])}</div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown('</div>',unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="panel">',unsafe_allow_html=True)
        st.markdown("### 💳 Debt Health")
        st.metric("Debt-to-Income",fmt_pct(metrics["debt_to_income"]))
        st.progress(min(1,metrics["debt_to_income"]/0.5 if metrics["debt_to_income"] else 0))
        st.caption(f"Monthly payment: {fmt_money(metrics['debt_monthly'])}")
        st.caption(f"Paid: {fmt_pct(metrics['debt_paid']/metrics['debt_original'] if metrics['debt_original'] else 0)}")
        st.markdown('</div>',unsafe_allow_html=True)

    with c3:
        st.markdown('<div class="panel">',unsafe_allow_html=True)
        st.markdown("### 🧠 Financial Health")
        st.markdown(f'<div class="big-score">{metrics["score"]}</div>',unsafe_allow_html=True)
        st.markdown(f"**{metrics['score_status']}**")
        st.caption("Score considers cash flow, saving rate, emergency fund, debt burden, investment rate and net-worth trend.")
        st.markdown('</div>',unsafe_allow_html=True)

    st.markdown('<div class="section">💎 Balance Sheet</div>',unsafe_allow_html=True)
    fig = go.Figure()
    fig.add_bar(y=["Assets","Debt","Net Worth"], x=[metrics["assets"],metrics["debt"],metrics["net_worth"]], orientation="h")
    fig.update_layout(title="Assets vs Debt vs Net Worth")
    st.plotly_chart(styled_fig(fig,280),use_container_width=True)


# ============================================================
# TAB 2 — Transactions
# ============================================================
with tabs[1]:
    st.markdown("### 💵 Transactions Log")
    st.caption("Double-click cells to edit. Add or delete rows with the editor. Month/Year are calculated from Date.")

    tdf = st.session_state.transactions.copy()
    if not tdf.empty:
        tdf["Date"] = safe_date_series(tdf["Date"])

    category_options = sorted(set(settings.get("Income Categories",[]) + settings.get("Expense Categories",[])))
    type_options = settings.get("Transaction Types",["Income","Expense"])
    account_options = settings.get("Accounts",["Cash","Bank","Other"])
    nw_options = settings.get("Need / Want",["Need","Want","Saving","Investment"])

    edited = st.data_editor(
        tdf,
        key="transactions_editor",
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Date": st.column_config.DateColumn("Date",format="DD/MM/YYYY"),
            "Type": st.column_config.SelectboxColumn("Type",options=type_options,required=False),
            "Category": st.column_config.SelectboxColumn("Category",options=category_options,required=False),
            "Account": st.column_config.SelectboxColumn("Account",options=account_options,required=False),
            "Need_Want": st.column_config.SelectboxColumn("Need / Want",options=nw_options,required=False),
            "Amount": st.column_config.NumberColumn("Amount (₭)",min_value=0,step=1000,format="%.0f"),
            "Month": st.column_config.NumberColumn("Month",disabled=True),
            "Year": st.column_config.NumberColumn("Year",disabled=True),
        },
        disabled=["Month","Year"],
    )
    edited["Date"] = safe_date_series(edited["Date"])
    edited["Amount"] = clean_numeric(edited["Amount"])
    edited["Month"] = edited["Date"].dt.month.astype("Int64")
    edited["Year"] = edited["Date"].dt.year.astype("Int64")
    st.session_state.transactions = edited

    st.markdown("---")
    selected_tx = edited[(edited["Year"] == selected_year) & (edited["Month"] == selected_month)] if not edited.empty else edited

    c1,c2 = st.columns(2)
    with c1:
        exp = selected_tx[selected_tx["Type"]=="Expense"].groupby("Category",dropna=False)["Amount"].sum().reset_index()
        if not exp.empty:
            fig = px.pie(exp,names="Category",values="Amount",hole=.62,title="🍩 Expense by Category")
            st.plotly_chart(styled_fig(fig,340),use_container_width=True)
        else:
            st.info("No expense data for selected month.")
    with c2:
        nw = selected_tx[selected_tx["Type"]=="Expense"].groupby("Need_Want",dropna=False)["Amount"].sum().reset_index()
        if not nw.empty:
            fig = px.pie(nw,names="Need_Want",values="Amount",hole=.62,title="Need vs Want / Saving / Investment")
            st.plotly_chart(styled_fig(fig,340),use_container_width=True)
        else:
            st.info("No Need/Want data for selected month.")


# ============================================================
# TAB 3 — Debt & Assets
# ============================================================
with tabs[2]:
    st.markdown("### 💳 Debt Tracker")
    debt_edit = st.data_editor(
        st.session_state.debts.copy(),
        key="debt_editor",
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Original Principal": st.column_config.NumberColumn("Original Principal (₭)",min_value=0,step=1000,format="%.0f"),
            "Principal Paid": st.column_config.NumberColumn("Principal Paid (₭)",min_value=0,step=1000,format="%.0f"),
            "Remaining Principal": st.column_config.NumberColumn("Remaining (₭)",disabled=True,format="%.0f"),
            "Interest Rate": st.column_config.NumberColumn("Interest Rate",min_value=0,step=.001,format="%.2%"),
            "Monthly Payment": st.column_config.NumberColumn("Monthly Payment (₭)",min_value=0,step=1000,format="%.0f"),
            "Due Date": st.column_config.DateColumn("Due Date",format="DD/MM/YYYY"),
            "Start Date": st.column_config.DateColumn("Start Date",format="DD/MM/YYYY"),
            "End Date": st.column_config.DateColumn("End Date",format="DD/MM/YYYY"),
            "Status": st.column_config.TextColumn("Status",disabled=True),
        },
        disabled=["Remaining Principal","Status"],
    )
    for c in ["Due Date","Start Date","End Date"]:
        debt_edit[c] = safe_date_series(debt_edit[c])
    for c in ["Original Principal","Principal Paid","Monthly Payment","Interest Rate"]:
        debt_edit[c] = clean_numeric(debt_edit[c])
    debt_edit["Remaining Principal"]=(debt_edit["Original Principal"]-debt_edit["Principal Paid"]).clip(lower=0)
    today=pd.Timestamp.today().normalize()
    def ui_debt_status(r):
        if not r["Debt ID"] and not r["Debt Name"]:
            return ""
        if r["Remaining Principal"]<=0:
            return "✅ PAID"
        if pd.notna(r["Due Date"]) and r["Due Date"]<today:
            return "⚠ OVERDUE"
        return "🔵 ACTIVE"
    debt_edit["Status"]=debt_edit.apply(ui_debt_status,axis=1)
    st.session_state.debts=debt_edit

    st.markdown('<div class="section">Balance Sheet</div>',unsafe_allow_html=True)
    d1,d2,d3,d4=st.columns(4)
    d1.metric("Total Debt",fmt_money(debt_edit["Remaining Principal"].sum()))
    d2.metric("Monthly Payment",fmt_money(debt_edit["Monthly Payment"].sum()))
    orig=debt_edit["Original Principal"].sum()
    paid=debt_edit["Principal Paid"].sum()
    d3.metric("Debt Paid",fmt_pct(paid/orig if orig else 0))
    d4.metric("Debt-to-Income",fmt_pct(metrics["debt_to_income"]))

    st.markdown("### 🏠 Asset Tracker")
    asset_edit = st.data_editor(
        st.session_state.assets.copy(),
        key="asset_editor",
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Purchase Date": st.column_config.DateColumn("Purchase Date",format="DD/MM/YYYY"),
            "Purchase Cost": st.column_config.NumberColumn("Purchase Cost (₭)",min_value=0,step=1000,format="%.0f"),
            "Current Value": st.column_config.NumberColumn("Current Value (₭)",min_value=0,step=1000,format="%.0f"),
            "Category": st.column_config.SelectboxColumn("Category",options=settings.get("Asset Categories",[]),required=False),
            "Account": st.column_config.SelectboxColumn("Account",options=settings.get("Accounts",[]),required=False),
            "Liquidity": st.column_config.SelectboxColumn("Liquidity",options=settings.get("Liquidity",[]),required=False),
        },
    )
    asset_edit["Purchase Date"]=safe_date_series(asset_edit["Purchase Date"])
    asset_edit["Purchase Cost"]=clean_numeric(asset_edit["Purchase Cost"])
    asset_edit["Current Value"]=clean_numeric(asset_edit["Current Value"])
    st.session_state.assets=asset_edit

    fig=go.Figure()
    av=float(asset_edit["Current Value"].sum())
    dv=float(debt_edit["Remaining Principal"].sum())
    nw=av-dv
    fig.add_bar(
        y=["Assets","Debt","Net Worth"],
        x=[av,dv,nw],
        orientation="h",
        text=[fmt_money(av),fmt_money(dv),fmt_money(nw)],
        textposition="auto",
    )
    fig.update_layout(title="Horizontal Net Worth Distribution")
    st.plotly_chart(styled_fig(fig,300),use_container_width=True)


# ============================================================
# TAB 4 — Emergency + Investment
# ============================================================
with tabs[3]:
    st.markdown("### 🛟 Emergency Fund")
    ef_df=pd.DataFrame([st.session_state.emergency])
    ef_edit=st.data_editor(
        ef_df,
        key="ef_editor",
        num_rows="fixed",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Monthly Essential Expense": st.column_config.NumberColumn("Monthly Essential Expense (₭)",min_value=0,step=1000,format="%.0f"),
            "Target Months": st.column_config.NumberColumn("Target Months",min_value=1,max_value=24,step=1,format="%.0f"),
            "Current Fund": st.column_config.NumberColumn("Current Fund (₭)",min_value=0,step=1000,format="%.0f"),
        },
    )
    row=ef_edit.iloc[0]
    st.session_state.emergency={
        "Monthly Essential Expense":float(row["Monthly Essential Expense"]),
        "Target Months":float(row["Target Months"]),
        "Current Fund":float(row["Current Fund"]),
    }
    ef_target=st.session_state.emergency["Monthly Essential Expense"]*st.session_state.emergency["Target Months"]
    ef_current=st.session_state.emergency["Current Fund"]
    ef_p=ef_current/ef_target if ef_target else 0
    c1,c2,c3=st.columns(3)
    c1.metric("Target Fund",fmt_money(ef_target))
    c2.metric("Current Fund",fmt_money(ef_current))
    c3.metric("Months Covered",f"{ef_current/st.session_state.emergency['Monthly Essential Expense']:.1f}" if st.session_state.emergency["Monthly Essential Expense"] else "0.0")
    st.markdown(
        f"""
        <div class="panel">
            <div style="display:flex;justify-content:space-between">
                <b>Emergency Fund Progress</b><b>{min(100,ef_p*100):.1f}%</b>
            </div>
            <div class="progress-wrap"><div class="progress-bg"><div class="progress-fill" style="width:{min(100,ef_p*100)}%"></div></div></div>
            <div style="color:#94a3b8">Target = Monthly Essential Expense × Target Months</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### 📈 Saving & Investment")
    selected_tx=st.session_state.transactions[
        (st.session_state.transactions["Year"]==selected_year)&
        (st.session_state.transactions["Month"]==selected_month)
    ]
    saving=float(selected_tx.loc[(selected_tx["Type"]=="Expense")&(selected_tx["Need_Want"]=="Saving"),"Amount"].sum())
    inv=float(selected_tx.loc[(selected_tx["Type"]=="Expense")&(selected_tx["Need_Want"]=="Investment"),"Amount"].sum())
    inc=float(selected_tx.loc[selected_tx["Type"]=="Income","Amount"].sum())
    sr=saving/inc if inc else 0
    ir=inv/inc if inc else 0

    a,b,c,d=st.columns(4)
    a.metric("Monthly Saving",fmt_money(saving))
    b.metric("Monthly Investment",fmt_money(inv))
    c.metric("Saving Rate",fmt_pct(sr))
    d.metric("Investment Rate",fmt_pct(ir))

    st.markdown("#### Allocation Log")
    log=st.session_state.saving_log.copy()
    log_edit=st.data_editor(
        log,
        key="saving_log_editor",
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={
            "Date":st.column_config.DateColumn("Date",format="DD/MM/YYYY"),
            "Type":st.column_config.SelectboxColumn("Type",options=["Saving","Investment"],required=False),
            "Category":st.column_config.SelectboxColumn("Category",options=settings.get("Saving Categories",[]),required=False),
            "Amount":st.column_config.NumberColumn("Amount (₭)",min_value=0,step=1000,format="%.0f"),
        },
    )
    log_edit["Date"]=safe_date_series(log_edit["Date"])
    log_edit["Amount"]=clean_numeric(log_edit["Amount"])
    st.session_state.saving_log=log_edit


# ============================================================
# TAB 5 — Monthly Report
# ============================================================
with tabs[4]:
    st.markdown(f"### 📅 Monthly Financial Report · {selected_year}")
    monthly=monthly_metrics(st.session_state.transactions,st.session_state.debts,st.session_state.assets,selected_year)

    display=monthly.copy()
    display["Month"]=display["Month"].map(month_label)
    st.dataframe(
        display[["Month","Income","Expense","Cash Flow","Saving","Investment","Debt","Assets","Net Worth","Saving Rate","Investment Rate"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Income":st.column_config.NumberColumn("Income (₭)",format="%.0f"),
            "Expense":st.column_config.NumberColumn("Expense (₭)",format="%.0f"),
            "Cash Flow":st.column_config.NumberColumn("Cash Flow (₭)",format="%.0f"),
            "Saving":st.column_config.NumberColumn("Saving (₭)",format="%.0f"),
            "Investment":st.column_config.NumberColumn("Investment (₭)",format="%.0f"),
            "Debt":st.column_config.NumberColumn("Debt (₭)",format="%.0f"),
            "Assets":st.column_config.NumberColumn("Assets (₭)",format="%.0f"),
            "Net Worth":st.column_config.NumberColumn("Net Worth (₭)",format="%.0f"),
            "Saving Rate":st.column_config.NumberColumn("Saving Rate",format="%.1%"),
            "Investment Rate":st.column_config.NumberColumn("Investment Rate",format="%.1%"),
        }
    )

    c1,c2=st.columns(2)
    with c1:
        fig=go.Figure()
        fig.add_scatter(x=monthly["Month Name"],y=monthly["Cash Flow"],mode="lines+markers",name="Cash Flow")
        fig.add_hline(y=0,line_dash="dot")
        fig.update_layout(title="Smooth Cash Flow Trend")
        st.plotly_chart(styled_fig(fig,350),use_container_width=True)
    with c2:
        fig=go.Figure()
        fig.add_scatter(x=monthly["Month Name"],y=monthly["Assets"],mode="lines",name="Assets")
        fig.add_scatter(x=monthly["Month Name"],y=monthly["Debt"],mode="lines",name="Debt")
        fig.add_scatter(x=monthly["Month Name"],y=monthly["Net Worth"],mode="lines+markers",name="Net Worth")
        fig.update_layout(title="Assets · Debt · Net Worth")
        st.plotly_chart(styled_fig(fig,350),use_container_width=True)

    # Annual summary
    best_income=monthly.loc[monthly["Income"].idxmax()]
    worst_cf=monthly.loc[monthly["Cash Flow"].idxmin()]
    highest_exp=monthly.loc[monthly["Expense"].idxmax()]
    avg_exp=monthly["Expense"].mean()
    avg_saving=monthly["Saving"].mean()

    c=st.columns(5)
    c[0].metric("Highest Income Month",month_label(best_income["Month"]))
    c[1].metric("Worst Cash Flow",month_label(worst_cf["Month"]))
    c[2].metric("Highest Expense",month_label(highest_exp["Month"]))
    c[3].metric("Avg Monthly Expense",fmt_money(avg_exp))
    c[4].metric("Avg Monthly Saving",fmt_money(avg_saving))


# ============================================================
# TAB 6 — Settings
# ============================================================
with tabs[5]:
    st.markdown("### ⚙️ System Settings")
    st.caption("Master data used by dropdowns. Add/delete rows directly. All Categories is automatically rebuilt from Income + Expense Categories.")

    # Show the most useful master lists as independent dynamic editors.
    editable_lists = [
        ("Years","📅 Years",int),
        ("Income Categories","💰 Income Categories",str),
        ("Expense Categories","💸 Expense Categories",str),
        ("Accounts","🏦 Accounts",str),
        ("Need / Want","🎯 Need / Want",str),
        ("Debt Status","💳 Debt Status",str),
        ("Asset Categories","🏠 Asset Categories",str),
        ("Investment Categories","📈 Investment Categories",str),
        ("EF Target Months","🛟 EF Target Months",float),
        ("Transaction Types","🔄 Transaction Types",str),
        ("Liquidity","💧 Liquidity",str),
        ("Saving Categories","💰 Saving Categories",str),
        ("Saving / Investment Types","📊 Saving / Investment Types",str),
    ]

    cols=st.columns(3)
    for idx,(key,title,typ) in enumerate(editable_lists):
        with cols[idx%3]:
            base=pd.DataFrame({key:st.session_state.settings.get(key,[])})
            ed=st.data_editor(
                base,
                key=f"settings_{key}",
                num_rows="dynamic",
                hide_index=True,
                use_container_width=True,
                column_config={
                    key: st.column_config.NumberColumn(title) if typ in (int,float) else st.column_config.TextColumn(title)
                },
            )
            vals=[]
            for v in ed[key].tolist():
                if pd.isna(v) or str(v).strip()=="":
                    continue
                try:
                    vals.append(int(float(v)) if typ is int else float(v) if typ is float else str(v))
                except Exception:
                    vals.append(str(v))
            st.session_state.settings[key]=vals

    all_categories=list(dict.fromkeys(
        st.session_state.settings.get("Income Categories",[]) +
        st.session_state.settings.get("Expense Categories",[])
    ))
    st.markdown("### 🔗 All Categories Helper")
    st.dataframe(pd.DataFrame({"All Categories":all_categories}),use_container_width=True,hide_index=True)

    st.info("Tip: after changing settings, use the dropdowns in Transactions / Assets again. Download the updated Excel to keep the new master lists.")


# ============================================================
# Footer
# ============================================================
st.markdown("---")
st.caption(
    "Personal Finance Web Dashboard · LAK (₭) · Streamlit + Pandas + Plotly · "
    "The workbook remains the source of truth; web calculations are derived from the actual workbook data."
)
