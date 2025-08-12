"""
Streamlit IDX Financial Analyzer (Manual Mapping + Auto Previous detection)
- Upload IDX Excel (.xls/.xlsx)
- Auto-detect sheets + improved parser
- UI for manual mapping (sheet, account column, numeric columns)
- Auto-detect previous period column from same sheet if present (and allow override)
- Compute basic ratios, show comparisons, fetch stock price (yfinance)
"""

import streamlit as st
import pandas as pd
import numpy as np
import io, re
import plotly.express as px
import yfinance as yf

st.set_page_config(page_title="IDX Analyzer (Manual Mapping + AutoPrev)", layout="wide")
st.title("IDX Financial Analyzer — Manual Mapping + Auto Previous")

# ---------------------
# Utility / Parser
# ---------------------
@st.cache_data
def read_workbook_bytes(bytes_data: bytes):
    try:
        return pd.read_excel(io.BytesIO(bytes_data), sheet_name=None)
    except Exception:
        # fallback for some .xls files
        return pd.read_excel(io.BytesIO(bytes_data), sheet_name=None, engine='xlrd')

def detect_idx_statement_sheets(sheets_dict):
    mapping = {'neraca': None, 'laba_rugi': None, 'arus_kas': None}
    keywords = {
        'neraca': ['laporan posisi keuangan', 'statement of financial position', 'balance sheet', 'posisi keuangan', 'statement of financial'],
        'laba_rugi': ['laba rugi', 'statement of profit', 'income statement', 'profit or loss', 'laporan laba rugi'],
        'arus_kas': ['arus kas', 'cash flows', 'statement of cash flows', 'laporan arus kas']
    }
    for name, df in sheets_dict.items():
        try:
            txt = " ".join(df.fillna("").astype(str).apply(lambda r: " ".join(r.values), axis=1).tolist()).lower()
        except Exception:
            txt = ""
        for k, kws in keywords.items():
            if any(kw in txt for kw in kws):
                if mapping[k] is None:
                    mapping[k] = name
    return mapping

def sheet_preview_df(df, nrows=10):
    df2 = df.copy().fillna("")
    # show top nrows for preview
    return df2.head(nrows)

def guess_account_and_numeric_cols(df):
    """
    Heuristics: account column = column with few numeric cells.
    Numeric columns = columns with many numeric-looking cells.
    Returns (account_col, [numeric_cols])
    """
    working = df.fillna("").astype(str)
    def is_num(s):
        s = s.strip()
        if s == "": return False
        s2 = s.replace('(', '-').replace(')', '').replace('.', '').replace(',', '')
        return bool(re.search(r'\d', s2))
    col_scores = {}
    for c in working.columns:
        col_scores[c] = working[c].apply(is_num).sum()
    # account col = min numeric count
    account_col = min(col_scores, key=col_scores.get)
    numeric_cols = [c for c, count in col_scores.items() if count > max(1, 0.2 * len(working))]
    # ensure account_col not in numeric_cols
    numeric_cols = [c for c in numeric_cols if c != account_col]
    return account_col, numeric_cols

def to_number_cell(x):
    if pd.isna(x): return np.nan
    s = str(x).strip()
    if s == "": return np.nan
    s = s.replace('(', '-').replace(')', '')
    # heuristic: if many dots and no comma, dots are thousand sep
    if s.count('.') > 1 and ',' not in s:
        s = s.replace('.', '')
    s = s.replace(',', '')
    s = re.sub(r'[^\d\-.]', '', s)
    try:
        return float(s)
    except:
        return np.nan

def extract_tidy(df, account_col, numeric_cols, prefer_id_label=True):
    working = df.copy()
    # rename duplicate cols to strings
    working.columns = [str(c) for c in working.columns]
    # pick account col values (clean bilingual)
    accounts = working[account_col].astype(str).apply(lambda s: s.split('  ')[0].strip() if '  ' in s else s.strip())
    tidy = pd.DataFrame(index=accounts)
    for c in numeric_cols:
        tidy[c] = working[c].apply(to_number_cell).values
    tidy.index.name = 'Account'
    # remove empty rows
    tidy = tidy.dropna(how='all')
    return tidy

def compute_key_ratios(neraca_df, laba_df):
    periods = neraca_df.columns.tolist() if not neraca_df.empty else (laba_df.columns.tolist() if not laba_df.empty else [])
    ratios = {}
    for p in periods:
        def find_first(df, keywords):
            if df is None or df.empty: return np.nan
            for idx in df.index:
                for k in keywords:
                    if k in str(idx).lower():
                        try:
                            return df.at[idx, p]
                        except Exception:
                            return np.nan
            return np.nan
        total_assets = find_first(neraca_df, ['total aset','total assets','jumlah aset','total assets'])
        total_equity = find_first(neraca_df, ['total ekuitas','equity','jumlah ekuitas','total equity'])
        total_liab = find_first(neraca_df, ['total kewajiban','total liabilities','kewajiban'])
        cash = find_first(neraca_df, ['kas','cash'])
        current_assets = find_first(neraca_df, ['aktiva lancar','aset lancar','current assets'])
        current_liab = find_first(neraca_df, ['kewajiban lancar','current liabilities','liabilities current'])
        revenue = find_first(laba_df, ['pendapatan','revenue','penjualan','sales'])
        net_income = find_first(laba_df, ['laba bersih','profit','net income','profit (loss)'])
        gross_profit = find_first(laba_df, ['laba bruto','gross profit'])
        def sd(a,b):
            try:
                a=float(a); b=float(b); return a/b if b!=0 else np.nan
            except: return np.nan
        ratios[p] = {
            'Current Ratio': sd(current_assets, current_liab),
            'Quick Ratio': sd(cash, current_liab),
            'Debt to Equity': sd(total_liab, total_equity),
            'ROA': sd(net_income, total_assets),
            'ROE': sd(net_income, total_equity),
            'Gross Margin': sd(gross_profit, revenue),
            'Net Margin': sd(net_income, revenue),
            'Revenue': revenue,
            'Net Income': net_income,
            'Total Assets': total_assets,
            'Total Equity': total_equity,
            'Cash': cash
        }
    return pd.DataFrame(ratios).T if ratios else pd.DataFrame()

# ---------------------
# Streamlit UI
# ---------------------
st.sidebar.header("Upload & Settings")
uploaded = st.sidebar.file_uploader("Upload file laporan keuangan (.xls/.xlsx) — dari IDX", type=['xls','xlsx'])
ticker = st.sidebar.text_input("Ticker (Yahoo format, e.g. AALI.JK)", value="")
price_period = st.sidebar.selectbox("Periode harga (yfinance)", options=['6mo','1y','2y','5y'], index=1)
run = st.sidebar.button("Proses")

if not uploaded:
    st.info("Silakan upload file laporan keuangan (contoh: ADRO / AADI).")
    st.stop()

# Read workbook
wb = read_workbook_bytes(uploaded.read())
sheet_names = list(wb.keys())
detected = detect_idx_statement_sheets(wb)

st.subheader("Step 1 — Pilih sheet (otomatis terdeteksi)")
st.write("Sheet ditemukan dalam workbook:", sheet_names)
st.write("Deteksi otomatis (heuristik):", detected)

col1, col2, col3 = st.columns([1,1,1])
with col1:
    selected_neraca = st.selectbox("Pilih sheet Neraca (Balance)", options=["(auto) "+str(detected.get('neraca'))] + sheet_names, index=0 if detected.get('neraca') else 1)
with col2:
    selected_laba = st.selectbox("Pilih sheet Laba Rugi (Income)", options=["(auto) "+str(detected.get('laba_rugi'))] + sheet_names, index=0 if detected.get('laba_rugi') else 1)
with col3:
    selected_arus = st.selectbox("Pilih sheet Arus Kas (Cashflow)", options=["(auto) "+str(detected.get('arus_kas'))] + sheet_names, index=0 if detected.get('arus_kas') else 1)

def normalize_selection(sel):
    # If user picked the "(auto) X" option, return X or None
    if isinstance(sel, str) and sel.startswith("(auto) "):
        v = sel.replace("(auto) ", "")
        return v if v != "None" else None
    return sel

sel_neraca_sheet = normalize_selection(selected_neraca)
sel_laba_sheet = normalize_selection(selected_laba)
sel_arus_sheet = normalize_selection(selected_arus)

# show previews & allow manual mapping of columns
st.subheader("Step 2 — Preview & Manual Mapping (Jika Perlu)")

def sheet_controls(sheet_name, key_prefix):
    if sheet_name is None:
        st.info(f"No sheet selected for {key_prefix}")
        return None, None, None
    df = wb[sheet_name].copy()
    st.write(f"Preview sheet `{sheet_name}`")
    st.dataframe(sheet_preview_df(df, nrows=8))
    # propose account & numeric cols
    acct_col_guess, numeric_guess = guess_account_and_numeric_cols(df)
    cols = list(df.columns)
    acct_col = st.selectbox(f"{key_prefix}: Pilih kolom akun (label)", options=cols, index=cols.index(acct_col_guess) if acct_col_guess in cols else 0, key=key_prefix+"_acct")
    numeric_cols = st.multiselect(f"{key_prefix}: Pilih kolom angka (urut dari left->right; pilih minimal 1)", options=cols, default=[c for c in numeric_guess if c in cols], key=key_prefix+"_num")
    return df, acct_col, numeric_cols

neraca_df_raw, neraca_acct_col, neraca_numeric_cols = sheet_controls(sel_neraca_sheet, "Neraca")
laba_df_raw, laba_acct_col, laba_numeric_cols = sheet_controls(sel_laba_sheet, "LabaRugi")
arus_df_raw, arus_acct_col, arus_numeric_cols = sheet_controls(sel_arus_sheet, "ArusKas")

# If user left numeric_cols empty, try to auto-guess from sheet content
if neraca_numeric_cols is None: neraca_numeric_cols = []
if laba_numeric_cols is None: laba_numeric_cols = []
if arus_numeric_cols is None: arus_numeric_cols = []

# ---------------------
# Auto detect Current vs Previous inside same sheet
# ---------------------
st.subheader("Step 3 — Auto-detect Current & Previous columns (dari kolom angka yang dipilih)")
def detect_current_prev_from_cols(cols_list):
    # choose rightmost as Current, previous as Previous if >1
    if not cols_list:
        return None, None
    # preserve order as in workbook columns
    # assume the columns appear left->right as in df.columns
    cur = cols_list[-1]
    prev = cols_list[-2] if len(cols_list) >= 2 else None
    return cur, prev

neraca_cur_col, neraca_prev_col = detect_current_prev_from_cols(neraca_numeric_cols)
laba_cur_col, laba_prev_col = detect_current_prev_from_cols(laba_numeric_cols)
arus_cur_col, arus_prev_col = detect_current_prev_from_cols(arus_numeric_cols)

# allow user to override detected mapping
def show_override_selector(sheet_raw, numeric_cols, detected_cur, detected_prev, prefix):
    if sheet_raw is None or not numeric_cols:
        return detected_cur, detected_prev
    cols = numeric_cols
    cur_choice = st.selectbox(f"{prefix} - Pilih kolom CURRENT (biasanya paling kanan)", options=["(auto) " + str(detected_cur)] + cols, key=prefix+"_cur")
    prev_choice = None
    if len(cols) >= 2:
        prev_choice = st.selectbox(f"{prefix} - Pilih kolom PREVIOUS (opsional)", options=["(auto) " + str(detected_prev)] + ["(none)"] + cols, key=prefix+"_prev")
    def norm(choice):
        if choice is None: return None
        if isinstance(choice, str) and choice.startswith("(auto) "):
            v = choice.replace("(auto) ","")
            return v if v != "None" else None
        if choice == "(none)":
            return None
        return choice
    return norm(cur_choice), norm(prev_choice)

neraca_cur_col, neraca_prev_col = show_override_selector(neraca_df_raw, neraca_numeric_cols, neraca_cur_col, neraca_prev_col, "Neraca")
laba_cur_col, laba_prev_col = show_override_selector(laba_df_raw, laba_numeric_cols, laba_cur_col, laba_prev_col, "LabaRugi")
arus_cur_col, arus_prev_col = show_override_selector(arus_df_raw, arus_numeric_cols, arus_cur_col, arus_prev_col, "ArusKas")

# ---------------------
# Extract tidy tables according to mapping
# ---------------------
st.subheader("Step 4 — Ekstraksi data (hasil tidy)")
def build_and_show_tidy(raw_df, acct_col, cur_col, prev_col, name):
    if raw_df is None or acct_col is None or cur_col is None:
        st.info(f"Hilang data untuk {name}. Periksa mapping sheet/kolom.")
        return pd.DataFrame(), pd.DataFrame()
    # use columns order: [prev_col (if exists), cur_col]
    numeric_cols = []
    if prev_col and prev_col in raw_df.columns and prev_col != cur_col:
        numeric_cols.append(prev_col)
    numeric_cols.append(cur_col)
    tidy = extract_tidy(raw_df, acct_col, numeric_cols)
    st.write(f"Tidy {name} (index=Account, cols={numeric_cols}):")
    st.dataframe(tidy.head(40))
    # split into current & previous tidy if prev exists
    tidy_cur = tidy[[c for c in tidy.columns if c==cur_col]].rename(columns={cur_col: 'Current'})
    tidy_prev = tidy[[c for c in tidy.columns if prev_col and c==prev_col]].rename(columns={prev_col: 'Previous'}) if prev_col else pd.DataFrame()
    return tidy_cur, tidy_prev

neraca_tidy_cur, neraca_tidy_prev = build_and_show_tidy(neraca_df_raw, neraca_acct_col, neraca_cur_col, neraca_prev_col, "Neraca")
laba_tidy_cur, laba_tidy_prev = build_and_show_tidy(laba_df_raw, laba_acct_col, laba_cur_col, laba_prev_col, "LabaRugi")
arus_tidy_cur, arus_tidy_prev = build_and_show_tidy(arus_df_raw, arus_acct_col, arus_cur_col, arus_prev_col, "ArusKas")

# ---------------------
# Compute ratios & comparison
# ---------------------
st.subheader("Step 5 — Analisis Rasio & Perbandingan")
ratios_cur = compute_key_ratios(neraca_tidy_cur, laba_tidy_cur) if not neraca_tidy_cur.empty or not laba_tidy_cur.empty else pd.DataFrame()
ratios_prev = compute_key_ratios(neraca_tidy_prev, laba_tidy_prev) if not neraca_tidy_prev.empty or not laba_tidy_prev.empty else pd.DataFrame()

if not ratios_cur.empty:
    st.write("Rasio (Current):")
    st.dataframe(ratios_cur.style.format("{:.4f}"))
else:
    st.info("Rasio Current tidak dapat dihitung (data kurang).")

if not ratios_prev.empty:
    st.write("Rasio (Previous):")
    st.dataframe(ratios_prev.style.format("{:.4f}"))

# Comparison chart for chosen metric
all_metrics = ['Revenue','Net Income','ROA','ROE','Debt to Equity','Current Ratio','Gross Margin','Net Margin']
metric = st.selectbox("Pilih metrik untuk perbandingan grafik", options=all_metrics)
comp_df_list = []
if not ratios_cur.empty:
    tmp = ratios_cur[[metric]].reset_index().rename(columns={'index':'Period', metric: 'Value'})
    tmp['Source'] = 'Current'
    comp_df_list.append(tmp)
if not ratios_prev.empty:
    tmp2 = ratios_prev[[metric]].reset_index().rename(columns={'index':'Period', metric: 'Value'})
    tmp2['Source'] = 'Previous'
    comp_df_list.append(tmp2)
if comp_df_list:
    comp_df = pd.concat(comp_df_list, ignore_index=True)
    fig = px.bar(comp_df, x='Period', y='Value', color='Source', barmode='group', title=f'Perbandingan {metric}')
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Tidak ada data cukup untuk grafik perbandingan.")

# ---------------------
# Price analysis
# ---------------------
st.subheader("Step 6 — Performa Saham (opsional)")
if ticker:
    try:
        hist = yf.Ticker(ticker).history(period=price_period)
        if hist.empty:
            st.warning("Tidak ada data harga untuk ticker ini.")
        else:
            st.line_chart(hist['Close'])
            st.metric("Harga terakhir", f"{hist['Close'][-1]:.2f}")
    except Exception as e:
        st.error("Gagal mengambil data harga: " + str(e))
else:
    st.info("Masukkan ticker di sidebar untuk melihat grafik harga.")

# ---------------------
# Save / Export options (simple)
# ---------------------
st.subheader("Step 7 — Export / Save")
if st.button("Download Tidy (Current) as Excel"):
    # combine tidy tables into a single excel in-memory
    with io.BytesIO() as towrite:
        with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
            if not neraca_tidy_cur.empty: neraca_tidy_cur.to_excel(writer, sheet_name='Neraca_Current')
            if not laba_tidy_cur.empty: laba_tidy_cur.to_excel(writer, sheet_name='Laba_Current')
            if not arus_tidy_cur.empty: arus_tidy_cur.to_excel(writer, sheet_name='Arus_Current')
        towrite.seek(0)
        st.download_button("Download file Excel", data=towrite.read(), file_name="tidy_current.xlsx")

st.markdown("""
**Catatan & Tips**
- Jika parsing otomatis gagal, gunakan dropdown di Step 1 & 2 untuk memilih sheet dan kolom yang benar.
- Auto-previous mengambil kolom angka paling kanan sebagai CURRENT dan kolom sebelah kirinya sebagai PREVIOUS (jika ada). Anda bisa override di Step 3.
- Jika workbook Anda menyimpan tahun/period di sheet terpisah (mis. `1210000_2023`, `1210000_2022`), pilih sheet previous di Step 1 untuk perbandingan.
""")
