import streamlit as st
import pandas as pd
import numpy as np
import re
import tempfile
import os
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

BALANCE_KEYWORDS = ['laporan posisi keuangan','statement of financial position','total assets','jumlah aset','aset']
INCOME_KEYWORDS  = ['laba rugi','income statement','statement of profit','penjualan','revenue','sales','pendapatan']
CASH_KEYWORDS    = ['arus kas','cash flows','statement of cash flows','kas dan setara kas','arus kas dari aktivitas']

def text_contains_any(cell, keywords):
    if not isinstance(cell, str):
        cell = str(cell)
    s = cell.lower()
    return any(kw in s for kw in keywords)

def table_to_period_df(df_raw):
    df = df_raw.copy().replace(r'^\s*$', np.nan, regex=True).dropna(how='all', axis=0).dropna(how='all', axis=1)
    if df.shape[1] < 2:
        return pd.DataFrame()
    header_idx = 0
    for i in range(min(4, df.shape[0])):
        row = df.iloc[i].astype(str).tolist()
        year_like = sum(1 for cell in row if re.search(r"(19|20)\d{2}", str(cell)))
        if year_like >= 1:
            header_idx = i
            break
    header = df.iloc[header_idx].astype(str).tolist()
    body = df.iloc[header_idx+1:].copy()
    body.columns = header
    labels = body.iloc[:,0].astype(str).str.strip().tolist()
    value_cols = list(body.columns[1:])
    rows = []
    for row_idx, lbl in enumerate(labels):
        for col in value_cols:
            rawval = body.iat[row_idx, body.columns.get_loc(col)]
            rows.append((str(col).strip(), lbl.strip(), clean_number(rawval)))
    if not rows:
        return pd.DataFrame()
    df_long = pd.DataFrame(rows, columns=['period','account','value'])
    pivot = df_long.pivot_table(index='period', columns='account', values='value', aggfunc='first')
    try:
        pivot.index = pd.Index([p.strip() for p in pivot.index])
    except:
        pass
    return pivot

def extract_tables_from_pdf(file_bytes):
    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp.write(file_bytes)
    tmp.close()
    tables = []
    try:
        cam_tables = camelot.read_pdf(tmp.name, pages='all', flavor='stream')
        for t in cam_tables:
            tables.append(t.df)
    except Exception as e:
        st.warning(f"camelot extraction error: {e}")
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
    mapping = {'income':[], 'balance':[], 'cash':[]}
    for i, df in enumerate(raw_tables):
        sample = " ".join(df.astype(str).stack().head(30).astype(str).str.lower().tolist())
        if text_contains_any(sample, BALANCE_KEYWORDS):
            mapping['balance'].append(i)
        if text_contains_any(sample, INCOME_KEYWORDS):
            mapping['income'].append(i)
        if text_contains_any(sample, CASH_KEYWORDS):
            mapping['cash'].append(i)
    return mapping

def compute_basic_ratios(income_df, balance_df, cash_df):
    periods = sorted(set(list(income_df.index)+list(balance_df.index)+list(cash_df.index)), key=lambda x: str(x))
    rows=[]
    for p in periods:
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
        revenue = income_df.at[p, rev_col] if rev_col in income_df.columns and p in income_df.index else np.nan
        net_income = income_df.at[p, ni_col] if ni_col in income_df.columns and p in income_df.index else np.nan
        total_assets = balance_df.at[p, ta_col] if ta_col in balance_df.columns and p in balance_df.index else np.nan
        total_equity = balance_df.at[p, te_col] if te_col in balance_df.columns and p in balance_df.index else np.nan
        total_liab = balance_df.at[p, tl_col] if tl_col in balance_df.columns and p in balance_df.index else np.nan
        cash = balance_df.at[p, ca_col] if ca_col in balance_df.columns and p in balance_df.index else np.nan
        current_assets = balance_df.at[p, current_assets_col] if current_assets_col in balance_df.columns and p in balance_df.index else np.nan
        current_liab = balance_df.at[p, current_liab_col] if current_liab_col in balance_df.columns and p in balance_df.index else np.nan
        roe = net_income/total_equity if total_equity not in (0, np.nan) else np.nan
        roa = net_income/total_assets if total_assets not in (0, np.nan) else np.nan
        der = total_liab/total_equity if total_equity not in (0, np.nan) else np.nan
        cr  = current_assets/current_liab if current_liab not in (0, np.nan) else np.nan
        net_margin = net_income/revenue if revenue not in (0, np.nan) else np.nan
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
    return pd.DataFrame(rows).set_index('period')

# ---------------- UI ----------------
uploaded = st.file_uploader("Upload laporan (PDF atau Excel)", type=['pdf','xls','xlsx'])
if not uploaded:
    st.stop()

if uploaded.name.lower().endswith('.pdf'):
    raw_tables = extract_tables_from_pdf(uploaded.read())
    st.write(f"Berhasil ekstrak {len(raw_tables)} tabel.")
    mapping = detect_statement_tables(raw_tables)
    col1,col2,col3 = st.columns(3)
    bal_idx = col1.selectbox("Index Neraca", options=[None]+list(range(len(raw_tables))), index=(mapping['balance'][0] if mapping['balance'] else 0))
    inc_idx = col2.selectbox("Index Laba Rugi", options=[None]+list(range(len(raw_tables))), index=(mapping['income'][0] if mapping['income'] else 0))
    cash_idx = col3.selectbox("Index Arus Kas", options=[None]+list(range(len(raw_tables))), index=(mapping['cash'][0] if mapping['cash'] else 0))
    income_df = table_to_period_df(raw_tables[inc_idx]) if inc_idx is not None else pd.DataFrame()
    balance_df = table_to_period_df(raw_tables[bal_idx]) if bal_idx is not None else pd.DataFrame()
    cash_df = table_to_period_df(raw_tables[cash_idx]) if cash_idx is not None else pd.DataFrame()
else:
    xls = pd.read_excel(uploaded, sheet_name=None, header=None)
    candidate_balance = candidate_income = candidate_cash = None
    for name in xls:
        lc = str(name).lower()
        if '1210000' in lc or 'balance' in lc or 'neraca' in lc:
            candidate_balance = name
        if '1311000' in lc or 'income' in lc or 'laba' in lc:
            candidate_income = name
        if '1410000' in lc or 'cash' in lc or 'arus' in lc:
            candidate_cash = name
    def rawsheet_to_df(raw):
        return raw.astype(str).fillna('')
    income_df = table_to_period_df(rawsheet_to_df(xls[candidate_income])) if candidate_income else pd.DataFrame()
    balance_df = table_to_period_df(rawsheet_to_df(xls[candidate_balance])) if candidate_balance else pd.DataFrame()
    cash_df = table_to_period_df(rawsheet_to_df(xls[candidate_cash])) if candidate_cash else pd.DataFrame()

if income_df.empty or balance_df.empty:
    st.warning("Data belum lengkap untuk hitung rasio.")
else:
    ratios = compute_basic_ratios(income_df, balance_df, cash_df)

    # Styling fix (NO SyntaxError)
    style = ratios.style
    num_cols = ['revenue','net_income','total_assets','total_equity','total_liabilities','cash']
    style = style.format("{:.2f}", subset=num_cols)
    pct_cols = ['ROE', 'ROA', 'Net Margin']
    for col in pct_cols:
        style = style.format({col: "{:.2%}"})
    st.dataframe(style.fillna(""))

    sel = st.multiselect("Pilih metrik untuk grafik", options=['revenue','net_income','ROE','ROA','DER','Current Ratio','Net Margin'], default=['revenue','ROE'])
    if sel:
        plot_df = ratios[sel].reset_index().melt(id_vars='period', var_name='metric', value_name='value')
        fig = px.line(plot_df, x='period', y='value', color='metric', markers=True)
        st.plotly_chart(fig, use_container_width=True)
    st.download_button("Download CSV", data=ratios.reset_index().to_csv(index=False).encode('utf-8'), file_name='ratios.csv', mime='text/csv')
