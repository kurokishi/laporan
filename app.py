# streamlit_financial_analyzer_idx_xbrl.py
# Streamlit app to parse IDX XBRL-style Excel and compute financial ratios
# Usage: streamlit run streamlit_financial_analyzer_idx_xbrl.py

import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="IDX XBRL Financial Analyzer", layout="wide")
st.title("IDX XBRL Financial Analyzer (UNVR / IDX Excel)")

# ---------------- Utilities ----------------
def clean_number(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return x
    s = str(x).strip()
    if s in ["-", "—", ""]:
        return np.nan
    # remove footnotes/letters at end
    s = re.sub(r"[a-zA-Z\*\)\(]+$", "", s).strip()
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    s = s.replace(",", "").replace(" ", "")
    s = re.sub(r"[^\d\.\-]", "", s)
    try:
        val = float(s)
        return -val if neg else val
    except:
        return np.nan

def read_all_sheets(file):
    try:
        sheets = pd.read_excel(file, sheet_name=None, header=None)
        return sheets
    except Exception as e:
        st.error(f"Gagal baca file: {e}")
        return {}

def parse_context_sheet(ctx_df):
    """
    Try to build mapping context_id -> human period label (e.g., 'FY2022', '2022-03-31', 'Q1-2022').
    ctx_df expected raw (header=None). We'll try to find rows/cols that contain context id and dates.
    """
    # Normalize into two columns: id and period-string (best effort)
    # Many IDX XBRL Excel have two-column layout: [contextId, period/endDate or instant]
    ctx_df = ctx_df.dropna(how='all').fillna('')
    # convert to string grid and search
    text = ctx_df.astype(str).apply(lambda row: ' | '.join(row.values), axis=1)
    mapping = {}
    # search rows that look like "contextId <something> 2022" or contain 'instant' 'startDate' etc
    for i, row in ctx_df.iterrows():
        # bring all cells in row to single string
        cells = [str(x).strip() for x in row.tolist() if str(x).strip() != '']
        if len(cells) >= 2:
            key = cells[0]
            # find any date-like token in cells
            date_tokens = [c for c in cells[1:] if re.search(r"\b(19|20)\d{2}\b", c)]
            if date_tokens:
                label = ' / '.join(date_tokens)
            else:
                label = ' / '.join(cells[1:])
            mapping[key] = label
    # fallback: try first column as keys and second column as value if no mapping found
    if not mapping and ctx_df.shape[1] >= 2:
        for i, row in ctx_df.iterrows():
            key = str(row.iloc[0]).strip()
            val = str(row.iloc[1]).strip()
            if key and key.lower() not in ['context', 'id', '']:
                mapping[key] = val
    return mapping

def sheet_to_period_table(sheet_df, context_map):
    """
    Convert an XBRL-like financial sheet where:
    - first column = element/label (string)
    - subsequent columns = context ids (or period labels)
    Returns DataFrame indexed by period label with columns = element labels.
    """
    if sheet_df is None or sheet_df.shape[1] < 2:
        return pd.DataFrame()
    # find header row if any: try to detect if first row contains context ids (strings that match context_map keys)
    header_row = 0
    raw = sheet_df.copy()
    raw = raw.dropna(how='all').reset_index(drop=True)
    # convert potential header row to strings
    first_row = raw.iloc[header_row].astype(str).tolist()
    # If cells in first_row (except first col) match context_map keys, treat as header
    match_count = sum(1 for cell in first_row[1:] if cell in context_map)
    if match_count >= 1:
        header = first_row
        data = raw[1:].copy()
        data.columns = header
    else:
        # Maybe header already present as column names (we read with header=None) - let's force first row as labels if it looks like labels
        data = raw.copy()
    # Ensure first column is label
    labels = data.iloc[:,0].astype(str).tolist()
    values = data.iloc[:,1:].copy()
    # Use column names as context ids if they are in context_map, otherwise attempt to interpret them as period labels
    colnames = list(values.columns)
    period_labels = []
    for c in colnames:
        cstr = str(c)
        if cstr in context_map:
            period_labels.append(context_map[cstr])
        else:
            # clean and use as-is (maybe already a year like '2022')
            period_labels.append(cstr)
    # build pivot table: rows are periods, columns are labels
    long = []
    for row_idx, lbl in enumerate(labels):
        for col_idx, col in enumerate(values.columns):
            period = period_labels[col_idx]
            raw_val = values.iloc[row_idx, col_idx]
            val = clean_number(raw_val)
            long.append((period, lbl, val))
    if not long:
        return pd.DataFrame()
    df_long = pd.DataFrame(long, columns=['period', 'account', 'value'])
    pivot = df_long.pivot_table(index='period', columns='account', values='value', aggfunc='first')
    # try to sort periods meaningfully
    try:
        # extract year if present
        pivot.index = pd.Index([str(x) for x in pivot.index])
    except:
        pass
    return pivot

# Common keywords for mapping (Indonesian + English)
COMMON_ROW_KEYWORDS = {
    "total_revenue": ["pendapatan", "total pendapatan", "total revenue", "revenue", "penjualan"],
    "net_income": ["laba bersih", "laba tahun berjalan", "net income", "profit (loss)", "profit"],
    "total_assets": ["total aset", "total assets"],
    "total_equity": ["ekuitas", "total equity", "jumlah ekuitas"],
    "total_liabilities": ["total liabilitas", "total liabilities", "liabilitas"],
    "current_assets": ["aset lancar", "current assets"],
    "current_liabilities": ["liabilitas lancar", "current liabilities"],
    "cash": ["kas", "kas dan setara kas", "cash and cash equivalents", "cash"]
}

def find_account_column_candidates(df, keywords):
    """Return first matching column name in df.columns for each keyword list"""
    found = {}
    if df is None or df.empty:
        return found
    cols = [str(c).lower() for c in df.columns]
    for canonical, kws in keywords.items():
        for kw in kws:
            for i, c in enumerate(cols):
                if kw in c:
                    found[canonical] = list(df.columns)[i]
                    break
            if canonical in found:
                break
    return found

def compute_basic_ratios(income_df, balance_df, cash_df):
    # Align periods (union)
    periods = list(sorted(set(list(income_df.index) + list(balance_df.index) + list(cash_df.index)), key=lambda x: str(x)))
    rows = []
    for p in periods:
        def get_col_val(df, matches):
            if df is None or df.empty:
                return np.nan
            for m in matches:
                if m in df.columns:
                    try:
                        return df.loc[p, m]
                    except:
                        # try string match
                        try:
                            return df.loc[str(p), m]
                        except:
                            continue
            return np.nan
        # use inferred column names from actual DF
        income_cols = list(income_df.columns)
        balance_cols = list(balance_df.columns)
        cash_cols = list(cash_df.columns)

        # find by matching keywords inside column names
        income_map = find_account_column_candidates(income_df, {
            "total_revenue": COMMON_ROW_KEYWORDS["total_revenue"],
            "net_income": COMMON_ROW_KEYWORDS["net_income"],
        })
        balance_map = find_account_column_candidates(balance_df, {
            "total_assets": COMMON_ROW_KEYWORDS["total_assets"],
            "total_equity": COMMON_ROW_KEYWORDS["total_equity"],
            "total_liabilities": COMMON_ROW_KEYWORDS["total_liabilities"],
            "current_assets": COMMON_ROW_KEYWORDS["current_assets"],
            "current_liabilities": COMMON_ROW_KEYWORDS["current_liabilities"],
            "cash": COMMON_ROW_KEYWORDS["cash"]
        })
        cash_map = find_account_column_candidates(cash_df, {
            "operating_cash_flow": ["arus kas dari aktivitas operasi", "operating cash flow", "cash flow from operating activities"]
        })

        revenue = get_col_val(income_df, income_map.values()) if income_map.get("total_revenue") else np.nan
        net_income = get_col_val(income_df, income_map.values()) if income_map.get("net_income") else np.nan
        total_assets = get_col_val(balance_df, balance_map.values()) if balance_map.get("total_assets") else np.nan
        total_equity = get_col_val(balance_df, balance_map.values()) if balance_map.get("total_equity") else np.nan
        total_liab = get_col_val(balance_df, balance_map.values()) if balance_map.get("total_liabilities") else np.nan
        current_assets = get_col_val(balance_df, balance_map.values()) if balance_map.get("current_assets") else np.nan
        current_liab = get_col_val(balance_df, balance_map.values()) if balance_map.get("current_liabilities") else np.nan
        cash = get_col_val(balance_df, balance_map.values()) if balance_map.get("cash") else np.nan
        ocf = get_col_val(cash_df, cash_map.values()) if cash_map.get("operating_cash_flow") else np.nan

        roe = net_income / total_equity if not any(pd.isna([net_income, total_equity])) else np.nan
        roa = net_income / total_assets if not any(pd.isna([net_income, total_assets])) else np.nan
        der = total_liab / total_equity if not any(pd.isna([total_liab, total_equity])) else np.nan
        current_ratio = current_assets / current_liab if not any(pd.isna([current_assets, current_liab])) else np.nan
        net_margin = net_income / revenue if not any(pd.isna([net_income, revenue])) else np.nan
        rows.append({
            "period": p,
            "revenue": revenue,
            "net_income": net_income,
            "total_assets": total_assets,
            "total_equity": total_equity,
            "total_liabilities": total_liab,
            "cash": cash,
            "operating_cash_flow": ocf,
            "ROE": roe,
            "ROA": roa,
            "DER": der,
            "Current Ratio": current_ratio,
            "Net Margin": net_margin
        })
    df = pd.DataFrame(rows).set_index("period")
    return df

# ---------------- Streamlit UI ----------------
uploaded = st.file_uploader("Upload Excel (IDX XBRL-style) atau pilih contoh", type=["xls", "xlsx"])
if not uploaded:
    st.info("Unggah file Excel XBRL IDX (contoh: FinancialStatement-2022-I-UNVR.xlsx).")
    st.stop()

with st.spinner("Membaca file..."):
    raw_sheets = read_all_sheets(uploaded)
    if not raw_sheets:
        st.stop()

st.write("Sheet ditemukan:", list(raw_sheets.keys()))

# Build context mapping if present
context_map = {}
if 'Context' in raw_sheets:
    context_map = parse_context_sheet(raw_sheets['Context'])
    st.write("Context mapping (sample 10):", dict(list(context_map.items())[:10]))
else:
    st.warning("Sheet 'Context' tidak ditemukan — akan coba gunakan header kolom langsung sebagai periode.")

# pick candidate sheet names for financial statements (common IDX codes)
candidates = {name.lower(): name for name in raw_sheets.keys()}
balance_sheet_keys = [k for k in candidates if '1210000' in k or 'balance' in k or 'posisi' in k or 'neraca' in k]
income_sheet_keys = [k for k in candidates if '1311000' in k or 'income' in k or 'laba' in k or 'labarugi' in k or 'laba rugi' in k]
cash_sheet_keys = [k for k in candidates if '1410000' in k or 'cash' in k or 'arus' in k]

st.markdown("Pencarian sheet otomatis:")
st.write("Balance candidates:", balance_sheet_keys)
st.write("Income candidates:", income_sheet_keys)
st.write("Cashflow candidates:", cash_sheet_keys)

# UI to choose sheets
col1, col2, col3 = st.columns(3)
balance_choice = col1.selectbox("Pilih sheet Neraca (Balance Sheet)", options=[None] + list(raw_sheets.keys()), index=1 if balance_sheet_keys else 0)
income_choice = col2.selectbox("Pilih sheet Laba Rugi (Income Statement)", options=[None] + list(raw_sheets.keys()), index=1 if income_sheet_keys else 0)
cash_choice = col3.selectbox("Pilih sheet Arus Kas (Cash Flow)", options=[None] + list(raw_sheets.keys()), index=1 if cash_sheet_keys else 0)

if not (balance_choice and income_choice and cash_choice):
    st.warning("Pilih ketiga sheet (neraca, laba rugi, arus kas) agar analisis lengkap.")
    st.stop()

# Convert to period-indexed tables
with st.spinner("Memproses sheet menjadi tabel periodik..."):
    income_table = sheet_to_period_table(raw_sheets[income_choice], context_map)
    balance_table = sheet_to_period_table(raw_sheets[balance_choice], context_map)
    cash_table = sheet_to_period_table(raw_sheets[cash_choice], context_map)

st.subheader("Preview (beberapa kolom teratas)")
st.write("Income (preview):")
st.dataframe(income_table.head(10))
st.write("Balance (preview):")
st.dataframe(balance_table.head(10))
st.write("Cashflow (preview):")
st.dataframe(cash_table.head(10))

# Compute ratios
ratios = compute_basic_ratios(income_table, balance_table, cash_table)
if ratios.empty:
    st.error("Gagal menghitung rasio — data ter-parsing kosong atau tidak terdeteksi kolom yang relevan.")
else:
    st.subheader("Rasio Dasar")
    st.dataframe(ratios.style.format("{:.2f}", subset=["revenue","net_income","total_assets","total_equity","total_liabilities","cash","operating_cash_flow"]))
    st.dataframe(ratios.style.format("{:.2%}", subset=["ROE","ROA","Net Margin"]).fillna(""))

    st.subheader("Grafik Tren")
    selectable = ["revenue","net_income","ROE","ROA","DER","Current Ratio","Net Margin"]
    sel = st.multiselect("Pilih metrik", options=selectable, default=["revenue","ROE"])
    if sel:
        df_plot = ratios[sel].copy()
        df_plot = df_plot.reset_index().melt(id_vars="period", var_name="metric", value_name="value")
        fig = px.line(df_plot, x="period", y="value", color="metric", markers=True)
        st.plotly_chart(fig, use_container_width=True)

    csv = ratios.reset_index().to_csv(index=False).encode('utf-8')
    st.download_button("Download ratios CSV", data=csv, file_name="ratios_idx_xbrl.csv", mime="text/csv")

st.caption("Catatan: parser ini dibuat robust untuk banyak variasi format XBRL-Excel IDX, tetapi ada perusahaan yang pakai layout unik — kalau ada mismatch, kirim file contoh dan aku akan sesuaikan mapping label ke akun yang benar.")
