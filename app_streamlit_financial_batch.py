# app_streamlit_financial_batch.py
"""
Streamlit IDX Financial Analyzer - Single / Compare (2 files) / Batch Mode
Features:
- Upload single IDX Excel (.xls/.xlsx) and analyze
- Upload two files for direct comparison (auto detect current vs previous)
- Batch mode: upload many .xls/.xlsx files or a ZIP containing them -> produce screener Excel
- Manual mapping UI for sheet & columns if automatic parsing fails
- Export tidy/current screener results to Excel
"""

import streamlit as st
import pandas as pd, numpy as np
import io, re, zipfile, os, tempfile
import plotly.express as px
import yfinance as yf

st.set_page_config(page_title="IDX Financial Analyzer (Batch)", layout="wide")
st.title("IDX Financial Analyzer — Single / Compare / Batch")

# --------------------
# Parser utilities (reusable)
# --------------------
@st.cache_data
def read_workbook_bytes(bytes_data: bytes):
    try:
        return pd.read_excel(io.BytesIO(bytes_data), sheet_name=None)
    except Exception:
        # fallback for some .xls files
        return pd.read_excel(io.BytesIO(bytes_data), sheet_name=None, engine='xlrd')

def read_workbook_path(path):
    try:
        return pd.read_excel(path, sheet_name=None)
    except Exception:
        return pd.read_excel(path, sheet_name=None, engine='xlrd')

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

def guess_account_and_numeric_cols(df):
    working = df.fillna("").astype(str)
    def is_num(s):
        s = s.strip()
        if s == "": return False
        s2 = s.replace('(', '-').replace(')', '').replace('.', '').replace(',', '')
        return bool(re.search(r'\d', s2))
    col_scores = {}
    for c in working.columns:
        col_scores[c] = working[c].apply(is_num).sum()
    account_col = min(col_scores, key=col_scores.get)
    numeric_cols = [c for c, count in col_scores.items() if count > max(1, 0.2 * len(working))]
    numeric_cols = [c for c in numeric_cols if c != account_col]
    return account_col, numeric_cols

def to_number_cell(x):
    if pd.isna(x): return np.nan
    s = str(x).strip()
    if s == "": return np.nan
    s = s.replace('(', '-').replace(')', '')
    if s.count('.') > 1 and ',' not in s:
        s = s.replace('.', '')
    s = s.replace(',', '')
    s = re.sub(r'[^\d\-.]', '', s)
    try:
        return float(s)
    except:
        return np.nan

def extract_tidy(df, account_col, numeric_cols):
    working = df.copy()
    working.columns = [str(c) for c in working.columns]
    accounts = working[account_col].astype(str).apply(lambda s: s.split('  ')[0].strip() if '  ' in s else s.strip())
    tidy = pd.DataFrame(index=accounts)
    for c in numeric_cols:
        tidy[c] = working[c].apply(to_number_cell).values
    tidy.index.name = 'Account'
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
                        except:
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

# -------------------------
# Mode selection & Upload
# -------------------------
st.sidebar.header("Mode & Upload")
mode = st.sidebar.selectbox("Mode", options=["Single file", "Compare 2 files", "Batch (many files or ZIP)"])
st.sidebar.markdown("Tips: For Compare mode upload 2 files (newer and older). For Batch, upload many files or a ZIP.")

uploaded_files = st.sidebar.file_uploader("Upload file(s) (.xls .xlsx .zip)", accept_multiple_files=True, type=['xls','xlsx','zip'])
ticker = st.sidebar.text_input("Ticker (optional, Yahoo format, e.g. AALI.JK)")
price_period = st.sidebar.selectbox("Price period (yfinance)", options=['6mo','1y','2y','5y'], index=1)

if not uploaded_files:
    st.info("Silakan upload file (single/2 files/batch zip) di sidebar.")
    st.stop()

# Utility: extract files from uploaded list (including zip)
def expand_uploaded_files(uploaded_list):
    files = []  # tuples (filename, bytes)
    for upl in uploaded_list:
        name = upl.name
        data = upl.read()
        if name.lower().endswith('.zip'):
            # extract zip contents to temp and add xls/xlsx files
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for fname in z.namelist():
                    if fname.lower().endswith(('.xls','.xlsx')) and not fname.startswith('__MACOSX'):
                        files.append((os.path.basename(fname), z.read(fname)))
        else:
            files.append((name, data))
    return files

expanded = expand_uploaded_files(uploaded_files)
st.write(f"Files to process: {[n for n,_ in expanded]}")

# -------------------------
# Helper to process one workbook bytes and return tidy tables + ratios summary
# -------------------------
def process_workbook_bytes(name, bytes_data, allow_manual=False):
    result = {'name': name, 'sheets': None, 'neraca_tidy': None, 'laba_tidy': None, 'arus_tidy': None, 'ratios': None}
    try:
        wb = read_workbook_bytes(bytes_data)
    except Exception as e:
        result['error'] = str(e)
        return result
    result['sheets'] = list(wb.keys())
    detected = detect_idx_statement_sheets(wb)
    # default choose detected or first sheets
    neraca_sheet = detected.get('neraca') or (result['sheets'][0] if result['sheets'] else None)
    laba_sheet = detected.get('laba_rugi') or (result['sheets'][1] if len(result['sheets'])>1 else neraca_sheet)
    arus_sheet = detected.get('arus_kas') or (result['sheets'][2] if len(result['sheets'])>2 else None)
    # allow manual override? UI-level, not here in batch
    try:
        neraca_df = wb[neraca_sheet] if neraca_sheet in wb else pd.DataFrame()
        laba_df = wb[laba_sheet] if laba_sheet in wb else pd.DataFrame()
        arus_df = wb[arus_sheet] if arus_sheet in wb else pd.DataFrame()
    except Exception as e:
        neraca_df = pd.DataFrame(); laba_df=pd.DataFrame(); arus_df=pd.DataFrame()
    # guess account & numeric columns
    try:
        neraca_acct, neraca_nums = guess_account_and_numeric_cols(neraca_df) if not neraca_df.empty else (None, [])
    except Exception:
        neraca_acct, neraca_nums = (None, [])
    try:
        laba_acct, laba_nums = guess_account_and_numeric_cols(laba_df) if not laba_df.empty else (None, [])
    except Exception:
        laba_acct, laba_nums = (None, [])
    try:
        arus_acct, arus_nums = guess_account_and_numeric_cols(arus_df) if not arus_df.empty else (None, [])
    except Exception:
        arus_acct, arus_nums = (None, [])
    # pick current and previous as rightmost numeric cols
    def pick_cur_prev(nums):
        if not nums: return (None, None)
        cur = nums[-1]
        prev = nums[-2] if len(nums)>=2 else None
        return cur, prev
    neraca_cur, neraca_prev = pick_cur_prev(neraca_nums)
    laba_cur, laba_prev = pick_cur_prev(laba_nums)
    arus_cur, arus_prev = pick_cur_prev(arus_nums)
    # extract tidy
    neraca_tidy = extract_tidy(neraca_df, neraca_acct, [c for c in ([neraca_prev] if neraca_prev else []) + ([neraca_cur] if neraca_cur else [])]) if neraca_acct else pd.DataFrame()
    laba_tidy = extract_tidy(laba_df, laba_acct, [c for c in ([laba_prev] if laba_prev else []) + ([laba_cur] if laba_cur else [])]) if laba_acct else pd.DataFrame()
    arus_tidy = extract_tidy(arus_df, arus_acct, [c for c in ([arus_prev] if arus_prev else []) + ([arus_cur] if arus_cur else [])]) if arus_acct else pd.DataFrame()
    # compute ratios using Current column name if present
    # rename columns to 'Previous' and 'Current' for consistency if both exist
    def rename_cols(tidy, prev, cur):
        if tidy.empty: return tidy
        cols = list(tidy.columns)
        mapping = {}
        if prev and prev in cols:
            mapping[prev] = 'Previous'
        if cur and cur in cols:
            mapping[cur] = 'Current'
        return tidy.rename(columns=mapping)
    neraca_tidy = rename_cols(neraca_tidy, neraca_prev, neraca_cur)
    laba_tidy = rename_cols(laba_tidy, laba_prev, laba_cur)
    arus_tidy = rename_cols(arus_tidy, arus_prev, arus_cur)
    # compute ratios using Current if available, else available columns
    ratios = compute_key_ratios(neraca_tidy, laba_tidy)
    result.update({'neraca_tidy': neraca_tidy, 'laba_tidy': laba_tidy, 'arus_tidy': arus_tidy, 'ratios': ratios})
    return result

# -------------------------
# Mode: Single file
# -------------------------
if mode == "Single file":
    # take the first uploaded file only
    name, data = expanded[0]
    st.header("Single-file analysis: " + name)
    res = process_workbook_bytes(name, data)
    st.subheader("Detected sheets & brief info")
    st.write(res.get('sheets', []))
    # show previews and allow manual mapping if necessary
    # Reuse UI from previous app for manual mapping: show sheet selection and column pickers
    wb = read_workbook_bytes(data)
    sheets = list(wb.keys())
    detected = detect_idx_statement_sheets(wb)
    st.write("Auto-detected:", detected)
    st.markdown("If auto-detect wrong, select sheet & columns below.")
    col1, col2, col3 = st.columns(3)
    sel_neraca = col1.selectbox("Neraca sheet", options=["(auto) "+str(detected.get('neraca'))] + sheets, index=0)
    sel_laba = col2.selectbox("Laba sheet", options=["(auto) "+str(detected.get('laba_rugi'))] + sheets, index=0)
    sel_arus = col3.selectbox("Arus sheet", options=["(auto) "+str(detected.get('arus_kas'))] + sheets, index=0)
    def norm(sel):
        if sel and isinstance(sel,str) and sel.startswith("(auto) "): v=sel.replace("(auto) ",""); return v if v!="None" else None
        return sel
    sel_neraca = norm(sel_neraca); sel_laba = norm(sel_laba); sel_arus = norm(sel_arus)
    # show previews & let user pick account & numeric cols
    def sheet_controls_ui(sheet_name, label):
        if not sheet_name:
            st.info(f"No sheet for {label}")
            return None, None, None
        df = wb[sheet_name]
        st.write(f"Preview `{sheet_name}`")
        st.dataframe(df.head(8))
        acct_guess, num_guess = guess_account_and_numeric_cols(df)
        acct = st.selectbox(f"{label} - account column", options=list(df.columns), index=list(df.columns).index(acct_guess) if acct_guess in df.columns else 0)
        nums = st.multiselect(f"{label} - numeric columns (order left->right)", options=list(df.columns), default=[c for c in num_guess if c in df.columns])
        return df, acct, nums
    neraca_df_raw, neraca_acct_col, neraca_nums = sheet_controls_ui(sel_neraca, "Neraca")
    laba_df_raw, laba_acct_col, laba_nums = sheet_controls_ui(sel_laba, "Laba Rugi")
    arus_df_raw, arus_acct_col, arus_nums = sheet_controls_ui(sel_arus, "Arus Kas")
    # detect current/previous from numeric picks
    def detect_cur_prev(nums):
        if not nums: return (None, None)
        cur = nums[-1]; prev = nums[-2] if len(nums)>=2 else None
        return cur, prev
    neraca_cur, neraca_prev = detect_cur_prev(neraca_nums)
    laba_cur, laba_prev = detect_cur_prev(laba_nums)
    arus_cur, arus_prev = detect_cur_prev(arus_nums)
    # allow override
    neraca_cur = col1.selectbox("Neraca - Current col", options=["(auto) "+str(neraca_cur)] + (neraca_nums or []), index=0)
    neraca_prev = col1.selectbox("Neraca - Previous col (optional)", options=["(auto) "+str(neraca_prev), "(none)"] + (neraca_nums or []), index=0)
    def normcol(x):
        if isinstance(x,str) and x.startswith("(auto) "): v=x.replace("(auto) ",""); return None if v=="None" else v
        if x=="(none)": return None
        return x
    neraca_cur = normcol(neraca_cur); neraca_prev = normcol(neraca_prev)
    # extract tidy tables
    try:
        neraca_tidy = extract_tidy(neraca_df_raw, neraca_acct_col, [c for c in ([neraca_prev] if neraca_prev else []) + ([neraca_cur] if neraca_cur else [])]) if neraca_df_raw is not None else pd.DataFrame()
    except:
        neraca_tidy = pd.DataFrame()
    try:
        laba_tidy = extract_tidy(laba_df_raw, laba_acct_col, [c for c in ([laba_prev] if laba_prev else []) + ([laba_cur] if laba_cur else [])]) if laba_df_raw is not None else pd.DataFrame()
    except:
        laba_tidy = pd.DataFrame()
    st.subheader("Tidy Preview (Current)")
    st.dataframe(neraca_tidy.head(40))
    ratios = compute_key_ratios(neraca_tidy, laba_tidy)
    st.subheader("Rasio (Current)")
    st.dataframe(ratios.style.format("{:.4f}"))
    # price chart
    if ticker:
        try:
            hist = yf.Ticker(ticker).history(period=price_period)
            if not hist.empty:
                st.line_chart(hist['Close'])
        except Exception as e:
            st.write("Gagal ambil harga:", e)
    # option to download tidy
    if st.button("Download tidy current as Excel"):
        with io.BytesIO() as towrite:
            with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
                if not neraca_tidy.empty: neraca_tidy.to_excel(writer, sheet_name='Neraca_Current')
                if not laba_tidy.empty: laba_tidy.to_excel(writer, sheet_name='Laba_Current')
            towrite.seek(0)
            st.download_button("Download Excel", data=towrite.read(), file_name=f"tidy_{name}.xlsx")
# -------------------------
# Mode: Compare 2 files
# -------------------------
elif mode == "Compare 2 files":
    if len(expanded) < 2:
        st.error("Silakan upload 2 file untuk mode Compare.")
        st.stop()
    # pick exactly two files (allow user to select)
    names = [n for n,_ in expanded]
    sel1 = st.selectbox("Pilih file NEWER (current)", options=names, index=0)
    sel2 = st.selectbox("Pilih file OLDER (previous)", options=names, index=1)
    data1 = dict(expanded)[sel1]
    data2 = dict(expanded)[sel2]
    st.header(f"Compare: {sel1} vs {sel2}")
    r1 = process_workbook_bytes(sel1, data1)
    r2 = process_workbook_bytes(sel2, data2)
    st.subheader("Detected sheets (newer)"); st.write(r1.get('sheets'))
    st.subheader("Detected sheets (older)"); st.write(r2.get('sheets'))
    # present tidies and ratios side by side
    c1,c2 = st.columns(2)
    with c1:
        st.write("Newer - Neraca (Current)"); st.dataframe(r1['neraca_tidy'].head(30))
        st.write("Newer - Laba (Current)"); st.dataframe(r1['laba_tidy'].head(30))
        st.write("Newer - Ratios"); st.dataframe(r1['ratios'].style.format("{:.4f}"))
    with c2:
        st.write("Older - Neraca (Current)"); st.dataframe(r2['neraca_tidy'].head(30))
        st.write("Older - Laba (Current)"); st.dataframe(r2['laba_tidy'].head(30))
        st.write("Older - Ratios"); st.dataframe(r2['ratios'].style.format("{:.4f}"))
    # comparison charts for a metric
    metric = st.selectbox("Pilih metrik untuk perbandingan", options=['Revenue','Net Income','ROA','ROE','Debt to Equity','Current Ratio','Gross Margin','Net Margin'])
    comp_df = []
    if not r1['ratios'].empty:
        tmp = r1['ratios'][[metric]].reset_index().rename(columns={'index':'Period', metric:'Value'})
        tmp['Source'] = 'Newer'
        comp_df.append(tmp)
    if not r2['ratios'].empty:
        tmp2 = r2['ratios'][[metric]].reset_index().rename(columns={'index':'Period', metric:'Value'})
        tmp2['Source'] = 'Older'
        comp_df.append(tmp2)
    if comp_df:
        comp_df = pd.concat(comp_df, ignore_index=True)
        fig = px.bar(comp_df, x='Period', y='Value', color='Source', barmode='group', title=f'Perbandingan {metric}')
        st.plotly_chart(fig, use_container_width=True)
    # download comparison summary
    if st.button("Download comparison summary Excel"):
        with io.BytesIO() as towrite:
            with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
                if not r1['neraca_tidy'].empty: r1['neraca_tidy'].to_excel(writer, sheet_name='Newer_Neraca')
                if not r1['laba_tidy'].empty: r1['laba_tidy'].to_excel(writer, sheet_name='Newer_Laba')
                if not r2['neraca_tidy'].empty: r2['neraca_tidy'].to_excel(writer, sheet_name='Older_Neraca')
                if not r2['laba_tidy'].empty: r2['laba_tidy'].to_excel(writer, sheet_name='Older_Laba')
            towrite.seek(0)
            st.download_button("Download Excel", data=towrite.read(), file_name=f"compare_{sel1}_vs_{sel2}.xlsx")
# -------------------------
# Mode: Batch
# -------------------------
elif mode == "Batch (many files or ZIP)":
    st.header("Batch processing — generate screener")
    # process all expanded files
    results = []
    progress = st.progress(0)
    total = len(expanded)
    for i, (name, data) in enumerate(expanded):
        progress.progress(int((i+1)/total*100))
        res = process_workbook_bytes(name, data)
        results.append(res)
    progress.empty()
    st.subheader("Processed files summary")
    summary_rows = []
    for r in results:
        summary_rows.append({'file': r['name'], 'sheets': ", ".join(r['sheets'] or []), 'ratios_available': not r['ratios'].empty if r.get('ratios') is not None else False})
    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(summary_df)
    # build screener table: pick some key metrics from ratios 'Current' if present
    screener_rows = []
    for r in results:
        name = r['name']
        ratios = r.get('ratios')
        if ratios is None or ratios.empty:
            screener_rows.append({'file': name, 'Revenue': np.nan, 'Net Income': np.nan, 'ROA': np.nan, 'ROE': np.nan, 'Debt to Equity': np.nan})
        else:
            # pick latest period row (index) if multiple periods
            try:
                latest = ratios.index.max()
                row = ratios.loc[latest]
                screener_rows.append({'file': name, 'Period': latest, 'Revenue': row.get('Revenue'), 'Net Income': row.get('Net Income'), 'ROA': row.get('ROA'), 'ROE': row.get('ROE'), 'Debt to Equity': row.get('Debt to Equity')})
            except Exception:
                screener_rows.append({'file': name, 'Revenue': np.nan, 'Net Income': np.nan, 'ROA': np.nan, 'ROE': np.nan, 'Debt to Equity': np.nan})
    screener_df = pd.DataFrame(screener_rows).set_index('file')
    st.subheader("Screener (basic)")
    st.dataframe(screener_df.style.format("{:.4f}"))
    # ranking example: sort by ROE desc then Debt to Equity asc
    st.subheader("Ranking contoh (ROE desc, DebtToEquity asc)")
    ranked = screener_df.reset_index().sort_values(by=['ROE','Debt to Equity'], ascending=[False, True])
    st.dataframe(ranked)
    # allow download screener as Excel
    if st.button("Download screener as Excel"):
        with io.BytesIO() as towrite:
            with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
                screener_df.to_excel(writer, sheet_name='Screener')
                # also save individual tidies
                for r in results:
                    base = os.path.splitext(r['name'])[0][:30]
                    if r.get('neraca_tidy') is not None and not r['neraca_tidy'].empty:
                        r['neraca_tidy'].to_excel(writer, sheet_name=f"{base}_neraca")
                    if r.get('laba_tidy') is not None and not r['laba_tidy'].empty:
                        r['laba_tidy'].to_excel(writer, sheet_name=f"{base}_laba")
            towrite.seek(0)
            st.download_button("Download ZIP Excel", data=towrite.read(), file_name="screener_results.xlsx")
    # optional: fetch price for all tickers if user provided tickers mapping (not implemented)
    st.info("Batch processing selesai. Untuk integrasi harga (yfinance) tambahkan mapping file->ticker.")
else:
    st.error("Mode tidak dikenali.")
