# streamlit_financial_analyzer_pdf_excel.py
# Streamlit app: baca Excel IDX XBRL OR PDF laporan keuangan IDX, ekstrak tabel & hitung rasio
# Jalankan: streamlit run streamlit_financial_analyzer_pdf_excel.py

import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import tempfile
import os

# PDF table extraction libs
import camelot
import pdfplumber

import plotly.express as px

st.set_page_config(page_title="IDX Financial Analyzer (PDF + Excel)", layout="wide")
st.title("IDX Financial Analyzer — PDF & Excel")

# ---------------- helpers ----------------
def clean_number(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return x
    s = str(x).strip()
    if s in ["-", "—", ""]:
        return np.nan
    # parentheses as negative
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    # remove footnotes/letters
    s = re.sub(r"[^\d\.\-\,\s]", "", s)
    s = s.replace(",", "").replace(" ", "")
    try:
        val = float(s)
        return -val if neg else val
    except:
        return np.nan

# keyword sets to detect which table is which
BALANCE_KEYWORDS = ['laporan posisi keuangan','statement of financial position','total assets','jumlah aset','aset']
INCOME_KEYWORDS  = ['laba rugi','income statement','statement of profit','penjualan','revenue','sales','pendapatan']
CASH_KEYWORDS    = ['arus kas','cash flows','statement of cash flows','kas dan setara kas','arus kas dari aktivitas']

def text_contains_any(cell, keywords):
    if not isinstance(cell, str):
        cell = str(cell)
    s = cell.lower()
    return any(kw in s for kw in keywords)

# Convert a raw DataFrame (like camelot table.df) to "period-indexed" table:
# heuristics: first row(s) include period headings (years or '31 December 2024'), first col = account labels
def table_to_period_df(df_raw):
    # df_raw is camelot's table.df (all strings)
    df = df_raw.copy().replace(r'^\s*$', np.nan, regex=True).dropna(how='all', axis=0).dropna(how='all', axis=1)
    if df.shape[1] < 2:
        return pd.DataFrame()
    # find header row: a row where many columns look like years or dates
    header_idx = 0
    for i in range(min(4, df.shape[0])):
        row = df.iloc[i].astype(str).tolist()
        year_like = sum(1 for cell in row if re.search(r"(19|20)\d{2}", str(cell)))
        if year_like >= 1:
            header_idx = i
            break
    # build table using header row as columns
    header = df.iloc[header_idx].astype(str).tolist()
    body = df.iloc[header_idx+1:].copy()
    body.columns = header
    # ensure first column are labels
    labels = body.iloc[:,0].astype(str).str.strip().tolist()
    value_cols = list(body.columns[1:])
    # build pivot: rows = period, columns = account labels
    rows = []
    for lbl in labels:
        for col in value_cols:
            rawval = body.at[body.index[labels.index(lbl)], col]
            rows.append((str(col).strip(), lbl.strip(), clean_number(rawval)))
    if not rows:
        return pd.DataFrame()
    df_long = pd.DataFrame(rows, columns=['period','account','value'])
    pivot = df_long.pivot_table(index='period', columns='account', values='value', aggfunc='first')
    # try to sort periods (prefer year-like)
    try:
        pivot.index = pd.Index([p.strip() for p in pivot.index])
    except:
        pass
    return pivot

# try extract tables from PDF using camelot; fallback to pdfplumber for pages where camelot fails
def extract_tables_from_pdf(file_bytes):
    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp.write(file_bytes)
    tmp.close()
    tables = []
    try:
        cam_tables = camelot.read_pdf(tmp.name, pages='all', flavor='stream')  # or 'lattice' if lines present
        for t in cam_tables:
            tables.append(t.df)
    except Exception as e:
        st.warning(f"camelot extraction error: {e}")
    # fallback: pdfplumber (extract per page as table-like using heuristics)
    if not tables:
        try:
            with pdfplumber.open(tmp.name) as pdf:
                for page in pdf.pages:
                    tbls = page.extract_tables()
                    for t in tbls:
                        df = pd.DataFrame(t[1:], columns=t[0])
                        tables.append(df)
        except Exception as e:
            st.error(f"pdfplumber extraction failed: {e}")
    os.unlink(tmp.name)
    return tables

def detect_statement_tables(raw_tables):
    """Return dict with keys: income, balance, cash (value = index into raw_tables)"""
    mapping = {'income':[], 'balance':[], 'cash':[]}
    for i, df in enumerate(raw_tables):
        # examine some cells for keywords
        sample = " ".join(df.astype(str).stack().head(30).astype(str).str.lower().tolist())
        if text_contains_any(sample, BALANCE_KEYWORDS):
            mapping['balance'].append(i)
        if text_contains_any(sample, INCOME_KEYWORDS):
            mapping['income'].append(i)
        if text_contains_any(sample, CASH_KEYWORDS):
            mapping['cash'].append(i)
    return mapping

# compute basic ratios from period-indexed tables
def compute_basic_ratios(income_df, balance_df, cash_df):
    periods = sorted(set(list(income_df.index)+list(balance_df.index)+list(cash_df.index)), key=lambda x: str(x))
    rows=[]
    for p in periods:
        # best-effort find account names via contains
        def find_col(df, keywords):
            if df is None or df.empty:
                return None
            for c in df.columns:
                s = str(c).lower()
                if any(kw in s for kw in keywords):
                    return c
            return None
        rev_col = find_col(income_df, ['pendapatan','revenue','sales','penjualan','total revenue'])
        ni_col  = find_col(income_df, ['laba bersih','net income','total profit','jumlah laba'])
        ta_col  = find_col(balance_df, ['total aset','total assets','jumlah aset'])
        te_col  = find_col(balance_df, ['ekuitas','total equity','jumlah ekuitas'])
        tl_col  = find_col(balance_df, ['liabilitas','total liabilities','jumlah liabilitas'])
        ca_col  = find_col(balance_df, ['kas','kas dan setara kas','cash and cash equivalents'])
        current_assets_col = find_col(balance_df, ['aset lancar','current assets'])
        current_liab_col = find_col(balance_df, ['liabilitas lancar','current liabilities'])
        try:
            revenue = income_df.at[p, rev_col] if rev_col in income_df.columns and p in income_df.index else np.nan
        except: revenue = np.nan
        try:
            net_income = income_df.at[p, ni_col] if ni_col in income_df.columns and p in income_df.index else np.nan
        except: net_income = np.nan
        try:
            total_assets = balance_df.at[p, ta_col] if ta_col in balance_df.columns and p in balance_df.index else np.nan
        except: total_assets = np.nan
        try:
            total_equity = balance_df.at[p, te_col] if te_col in balance_df.columns and p in balance_df.index else np.nan
        except: total_equity = np.nan
        try:
            total_liab = balance_df.at[p, tl_col] if tl_col in balance_df.columns and p in balance_df.index else np.nan
        except: total_liab = np.nan
        try:
            cash = balance_df.at[p, ca_col] if ca_col in balance_df.columns and p in balance_df.index else np.nan
        except: cash = np.nan
        try:
            current_assets = balance_df.at[p, current_assets_col] if current_assets_col in balance_df.columns and p in balance_df.index else np.nan
        except: current_assets = np.nan
        try:
            current_liab = balance_df.at[p, current_liab_col] if current_liab_col in balance_df.columns and p in balance_df.index else np.nan
        except: current_liab = np.nan

        roe = net_income/total_equity if not pd.isna(net_income) and not pd.isna(total_equity) and total_equity!=0 else np.nan
        roa = net_income/total_assets if not pd.isna(net_income) and not pd.isna(total_assets) and total_assets!=0 else np.nan
        der = total_liab/total_equity if not pd.isna(total_liab) and not pd.isna(total_equity) and total_equity!=0 else np.nan
        cr  = current_assets/current_liab if not pd.isna(current_assets) and not pd.isna(current_liab) and current_liab!=0 else np.nan
        net_margin = net_income/revenue if not pd.isna(net_income) and not pd.isna(revenue) and revenue!=0 else np.nan

        rows.append({
            'period': p,
            'revenue': revenue,
            'net_income': net_income,
            'total_assets': total_assets,
            'total_equity': total_equity,
            'total_liabilities': total_liab,
            'cash': cash,
            'ROE': roe,
            'ROA': roa,
            'DER': der,
            'Current Ratio': cr,
            'Net Margin': net_margin
        })
    df = pd.DataFrame(rows).set_index('period')
    return df

# ---------------- Streamlit UI ----------------
uploaded = st.file_uploader("Upload laporan (Excel XBRL .xlsx/.xls atau PDF annual/quarterly)", type=['xls','xlsx','pdf'])
if not uploaded:
    st.info("Unggah file laporan tahunan/kuartalan IDX (PDF) atau file Excel XBRL.")
    st.stop()

if uploaded.name.lower().endswith('.pdf'):
    st.info("Mendeteksi PDF — mengekstrak tabel (camelot → pdfplumber fallback).")
    raw_tables = extract_tables_from_pdf(uploaded.read())
    st.success(f"Menemukan {len(raw_tables)} tabel (ekstraksi kasar).")
    # show small preview of first few
    for i, t in enumerate(raw_tables[:6]):
        st.write(f"Tabel #{i} (preview)")
        st.dataframe(t.head(8))

    st.write("Mendeteksi mana tabel yang mirip Neraca / Laba Rugi / Arus Kas...")
    mapping = detect_statement_tables(raw_tables)
    st.write("Deteksi otomatis (index tabel):", mapping)

    # let user override / choose table indices
    col1,col2,col3 = st.columns(3)
    bal_idx = col1.selectbox("Pilih index tabel Neraca (balance)", options=[None]+list(range(len(raw_tables))), index= mapping['balance'][0] if mapping['balance'] else 0)
    inc_idx = col2.selectbox("Pilih index tabel Laba Rugi (income)", options=[None]+list(range(len(raw_tables))), index= mapping['income'][0] if mapping['income'] else 0)
    cash_idx = col3.selectbox("Pilih index tabel Arus Kas (cashflow)", options=[None]+list(range(len(raw_tables))), index= mapping['cash'][0] if mapping['cash'] else 0)

    income_df = table_to_period_df(raw_tables[inc_idx]) if inc_idx is not None else pd.DataFrame()
    balance_df = table_to_period_df(raw_tables[bal_idx]) if bal_idx is not None else pd.DataFrame()
    cash_df = table_to_period_df(raw_tables[cash_idx]) if cash_idx is not None else pd.DataFrame()

else:
    # Excel XBRL handling (read all sheets and try to detect contexts like previous app)
    st.info("Mendeteksi Excel XBRL — mem-parsing semua sheet.")
    xls = pd.read_excel(uploaded, sheet_name=None, header=None)
    st.write("Sheets:", list(xls.keys()))
    # try to find 'Context' sheet and coded sheets like '1210000','1311000','1410000'
    # reuse earlier logic: build period-indexed for identified sheets
    candidate_balance = None
    candidate_income = None
    candidate_cash = None
    for name, df in xls.items():
        lc = str(name).lower()
        if '1210000' in lc or 'balance' in lc or 'posisi' in lc or 'neraca' in lc:
            candidate_balance = name
        if '1311000' in lc or 'income' in lc or 'laba' in lc:
            candidate_income = name
        if '1410000' in lc or 'cash' in lc or 'arus' in lc:
            candidate_cash = name
    st.write("Deteksi kandidat sheet:", candidate_balance, candidate_income, candidate_cash)
    # convert raw sheets to period tables using heuristic function (reuse table_to_period_df after converting to df)
    def rawsheet_to_df(raw):
        # raw is header=None read; convert to string DF
        return raw.astype(str).fillna('')
    income_df = table_to_period_df(rawsheet_to_df(xls[candidate_income])) if candidate_income else pd.DataFrame()
    balance_df = table_to_period_df(rawsheet_to_df(xls[candidate_balance])) if candidate_balance else pd.DataFrame()
    cash_df = table_to_period_df(rawsheet_to_df(xls[candidate_cash])) if candidate_cash else pd.DataFrame()

# show previews
st.subheader("Preview hasil konversi ke tabel periodik (jika tersedia)")
if not income_df.empty:
    st.write("Income (preview):")
    st.dataframe(income_df.head(6))
if not balance_df.empty:
    st.write("Balance (preview):")
    st.dataframe(balance_df.head(6))
if not cash_df.empty:
    st.write("Cashflow (preview):")
    st.dataframe(cash_df.head(6))

# compute ratios
if income_df.empty or balance_df.empty:
    st.warning("Income atau Balance belum tersedia untuk perhitungan rasio — coba pilih tabel lain atau upload file lain.")
else:
    ratios = compute_basic_ratios(income_df, balance_df, cash_df)
    st.subheader("Rasio dasar (hasil ekstraksi otomatis)")
    st.dataframe(ratios.style.format("{:.2f}", subset=['revenue','net_income','total_assets','total_equity','total_liabilities','cash']).applymap(lambda v: f\"{v:.2%}\" if isinstance(v,(float,)) and abs(v)<=1 else v, subset=['ROE','ROA','Net Margin']).fillna(""))
    # plot
    sel = st.multiselect("Pilih metrik untuk graf", options=['revenue','net_income','ROE','ROA','DER','Current Ratio','Net Margin'], default=['revenue','ROE'])
    if sel:
        plot_df = ratios[sel].reset_index().melt(id_vars='period', var_name='metric', value_name='value')
        fig = px.line(plot_df, x='period', y='value', color='metric', markers=True)
        st.plotly_chart(fig, use_container_width=True)
    st.download_button("Download rasio CSV", data=ratios.reset_index().to_csv(index=False).encode('utf-8'), file_name='ratios.csv', mime='text/csv')

st.caption("Jika tabel PDF sangat berantakan (split cells / multirow headers), pilih tabel yang sesuai secara manual. Aku bisa bantu tuning heuristik parsing untuk file tertentu — kirim file contoh (seperti ADHI yang kamu upload) & aku akan sesuaikan mapping otomatis.")
