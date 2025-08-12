import streamlit as st
import pandas as pd
import numpy as np
import re
import tempfile
import os
import camelot
import pdfplumber
import plotly.express as px

# Konfigurasi halaman Streamlit
st.set_page_config(page_title="IDX Financial Analyzer (PDF + Excel)", layout="wide")
st.title("IDX Financial Analyzer — PDF & Excel (FIXED)")

# ==================== FUNGSI UTILITAS ====================
def clean_number(x):
    """Membersihkan dan mengonversi string angka ke float"""
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float)):
        return x
    
    s = str(x).strip()
    # Handle nilai kosong
    if s in ["-", "—", ""]:
        return np.nan
    
    # Deteksi angka negatif dalam tanda kurung
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    
    # Hapus karakter non-numerik
    s = re.sub(r"[^\d\.\-]", "", s)
    
    try:
        val = float(s)
        return -val if neg else val
    except:
        return np.nan

# Keyword untuk deteksi jenis laporan keuangan
BALANCE_KEYWORDS = ['laporan posisi keuangan', 'statement of financial position', 'total assets', 'jumlah aset', 'aset']
INCOME_KEYWORDS = ['laba rugi', 'income statement', 'statement of profit', 'penjualan', 'revenue', 'sales', 'pendapatan']
CASH_KEYWORDS = ['arus kas', 'cash flows', 'statement of cash flows', 'kas dan setara kas', 'arus kas dari aktivitas']

def text_contains_any(cell, keywords):
    """Cek apakah teks mengandung salah satu keyword"""
    if not isinstance(cell, str):
        cell = str(cell)
    s = cell.lower()
    return any(kw in s for kw in keywords)

# ==================== EKSTRAKSI TABEL ====================
def table_to_period_df(df_raw):
    """
    Konversi tabel mentah menjadi DataFrame periodik
    Langkah:
    1. Bersihkan dataframe dari sel kosong
    2. Deteksi baris header
    3. Ekstrak label periode
    4. Transformasi ke format long
    5. Pivoting ke format period x account
    """
    # Normalisasi dan hapus baris/kolom kosong
    df = df_raw.copy()
    df = df.replace(r'^\s*$', np.nan, regex=True)
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    
    if df.empty or df.shape[1] < 2:
        return pd.DataFrame()
    
    # Deteksi header (baris pertama yang mengandung tahun)
    header_idx = 0
    for i in range(min(5, df.shape[0])):
        if any(re.search(r"(19|20)\d{2}", str(x)) for x in df.iloc[i]):
            header_idx = i
            break
    
    # Ekstrak body tabel
    header = df.iloc[header_idx].astype(str).tolist()
    body = df.iloc[header_idx+1:].reset_index(drop=True)
    body.columns = [str(h).strip() for h in header]
    
    # Hapus kolom kosong dan reset index
    body = body.dropna(how='all', axis=1)
    if body.shape[1] < 2:
        return pd.DataFrame()
    
    # Transformasi ke format long
    rows = []
    account_col = body.columns[0]
    
    for _, row in body.iterrows():
        account = str(row[account_col]).strip()
        if not account:
            continue
            
        for col in body.columns[1:]:
            period = str(col).strip()
            raw_value = row[col]
            value = clean_number(raw_value)
            
            # Skip nilai yang tidak valid
            if pd.isna(value) or period == 'nan':
                continue
                
            rows.append({
                'period': period,
                'account': account,
                'value': value
            })
    
    if not rows:
        return pd.DataFrame()
    
    # Buat pivot table
    pivot = pd.pivot_table(
        pd.DataFrame(rows),
        index='period',
        columns='account',
        values='value',
        aggfunc='first'
    )
    
    return pivot

def extract_tables_from_pdf(file_bytes):
    """Ekstrak tabel dari PDF menggunakan Camelot dan PDFPlumber"""
    tables = []
    
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    
    try:
        # Coba ekstrak dengan Camelot
        cam_tables = camelot.read_pdf(tmp_path, pages='all', flavor='stream')
        tables = [t.df for t in cam_tables]
    except Exception as e:
        st.warning(f"Camelot error: {str(e)}")
        try:
            # Fallback ke PDFPlumber
            with pdfplumber.open(tmp_path) as pdf:
                for page in pdf.pages:
                    for table in page.extract_tables():
                        if table:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            tables.append(df)
        except Exception as e2:
            st.error(f"PDFPlumber error: {str(e2)}")
    finally:
        os.unlink(tmp_path)
    
    return tables

def detect_statement_tables(raw_tables):
    """Identifikasi jenis laporan keuangan berdasarkan keyword"""
    mapping = {'income': [], 'balance': [], 'cash': []}
    
    for idx, df in enumerate(raw_tables):
        # Gabungkan konten tabel untuk deteksi keyword
        sample_text = " ".join(df.head(10).astype(str).values.flatten()).lower()
        
        if any(kw in sample_text for kw in BALANCE_KEYWORDS):
            mapping['balance'].append(idx)
        if any(kw in sample_text for kw in INCOME_KEYWORDS):
            mapping['income'].append(idx)
        if any(kw in sample_text for kw in CASH_KEYWORDS):
            mapping['cash'].append(idx)
    
    return mapping

# ==================== PERHITUNGAN RASIO ====================
def compute_basic_ratios(income_df, balance_df, cash_df):
    """Hitung rasio keuangan utama dari data laporan"""
    # Gabungkan semua periode yang tersedia
    periods = sorted(set(income_df.index) | set(balance_df.index) | set(cash_df), key=str)
    rows = []
    
    for period in periods:
        # Fungsi pencarian kolom dengan keyword
        def find_account(df, keywords):
            if df is None or df.empty:
                return None
            for col in df.columns:
                if any(kw in str(col).lower() for kw in keywords):
                    return col
            return None
        
        # Cari akun-akun utama
        revenue_acc = find_account(income_df, ['pendapatan', 'revenue', 'sales'])
        net_income_acc = find_account(income_df, ['laba bersih', 'net income'])
        total_assets_acc = find_account(balance_df, ['total aset', 'total assets'])
        total_equity_acc = find_account(balance_df, ['total ekuitas', 'total equity'])
        total_liab_acc = find_account(balance_df, ['total liabilitas', 'total liabilities'])
        cash_acc = find_account(balance_df, ['kas', 'cash'])
        current_assets_acc = find_account(balance_df, ['aset lancar', 'current assets'])
        current_liab_acc = find_account(balance_df, ['liabilitas jangka pendek', 'current liabilities'])
        
        # Ekstrak nilai dengan penanganan error
        def safe_get(df, period, account):
            try:
                if account and period in df.index:
                    return df.at[period, account]
            except:
                pass
            return np.nan
        
        revenue = safe_get(income_df, period, revenue_acc)
        net_income = safe_get(income_df, period, net_income_acc)
        total_assets = safe_get(balance_df, period, total_assets_acc)
        total_equity = safe_get(balance_df, period, total_equity_acc)
        total_liab = safe_get(balance_df, period, total_liab_acc)
        cash = safe_get(balance_df, period, cash_acc)
        current_assets = safe_get(balance_df, period, current_assets_acc)
        current_liab = safe_get(balance_df, period, current_liab_acc)
        
        # Hitung rasio
        roe = net_income / total_equity if all(pd.notna([net_income, total_equity])) and total_equity != 0 else np.nan
        roa = net_income / total_assets if all(pd.notna([net_income, total_assets])) and total_assets != 0 else np.nan
        der = total_liab / total_equity if all(pd.notna([total_liab, total_equity])) and total_equity != 0 else np.nan
        cr = current_assets / current_liab if all(pd.notna([current_assets, current_liab])) and current_liab != 0 else np.nan
        net_margin = net_income / revenue if all(pd.notna([net_income, revenue])) and revenue != 0 else np.nan
        
        rows.append({
            'period': period,
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

# ==================== ANTARMUKA PENGGUNA ====================
uploaded = st.file_uploader("Upload Laporan Keuangan (PDF atau Excel)", type=['pdf', 'xlsx', 'xls'])

if not uploaded:
    st.info("Silakan upload file PDF atau Excel laporan keuangan")
    st.stop()

# Pemrosesan file PDF
if uploaded.name.lower().endswith('.pdf'):
    with st.spinner("Memproses PDF..."):
        raw_tables = extract_tables_from_pdf(uploaded.read())
    
    if not raw_tables:
        st.error("Gagal mengekstrak tabel dari PDF")
        st.stop()
    
    st.success(f"Berhasil ekstrak {len(raw_tables)} tabel")
    mapping = detect_statement_tables(raw_tables)
    
    # Tampilkan pilihan tabel
    col1, col2, col3 = st.columns(3)
    options = ["Pilih otomatis"] + list(range(len(raw_tables)))
    
    # Fungsi pilihan default
    def get_default_idx(mapping_key):
        return mapping[mapping_key][0] + 1 if mapping.get(mapping_key) else 0
    
    bal_idx = col1.selectbox(
        "Tabel Neraca",
        options=options,
        index=get_default_idx('balance'),
        format_func=lambda x: "Otomatis" if x == "Pilih otomatis" else f"Tabel {x}"
    )
    
    inc_idx = col2.selectbox(
        "Tabel Laba Rugi",
        options=options,
        index=get_default_idx('income')
    )
    
    cash_idx = col3.selectbox(
        "Tabel Arus Kas",
        options=options,
        index=get_default_idx('cash')
    )
    
    # Konversi pilihan user
    def get_table(idx):
        if idx == "Pilih otomatis" or idx is None:
            return pd.DataFrame()
        return raw_tables[idx]
    
    income_df = table_to_period_df(get_table(inc_idx))
    balance_df = table_to_period_df(get_table(bal_idx))
    cash_df = table_to_period_df(get_table(cash_idx))

# Pemrosesan file Excel
else:
    with st.spinner("Memproses Excel..."):
        try:
            xls = pd.read_excel(uploaded, sheet_name=None, header=None)
            sheet_names = list(xls.keys())
            
            # Deteksi sheet berdasarkan pola nama
            balance_sheet = next((s for s in sheet_names if 'neraca' in s.lower() or 'balance' in s.lower()), None)
            income_sheet = next((s for s in sheet_names if 'laba' in s.lower() or 'income' in s.lower()), None)
            cash_sheet = next((s for s in sheet_names if 'arus' in s.lower() or 'cash' in s.lower()), None)
            
            # Konversi sheet yang terdeteksi
            income_df = table_to_period_df(xls[income_sheet]) if income_sheet else pd.DataFrame()
            balance_df = table_to_period_df(xls[balance_sheet]) if balance_sheet else pd.DataFrame()
            cash_df = table_to_period_df(xls[cash_sheet]) if cash_sheet else pd.DataFrame()
            
        except Exception as e:
            st.error(f"Error processing Excel: {str(e)}")
            st.stop()

# ==================== TAMPILAN HASIL ====================
if income_df.empty or balance_df.empty:
    st.warning("Data tidak lengkap untuk menghitung rasio. Periksa pemilihan tabel.")
    st.subheader("Pratinjau Tabel Terdeteksi")
    
    if 'raw_tables' in locals():
        table_idx = st.selectbox("Pilih tabel untuk preview", range(len(raw_tables)))
        st.dataframe(raw_tables[table_idx].head(10))
    
else:
    with st.spinner("Menghitung rasio..."):
        ratios = compute_basic_ratios(income_df, balance_df, cash_df)
    
    # Format tampilan tabel
    fmt_dict = {
        'revenue': '{:,.2f}',
        'net_income': '{:,.2f}',
        'total_assets': '{:,.2f}',
        'ROE': '{:.2%}',
        'ROA': '{:.2%}',
        'Net Margin': '{:.2%}'
    }
    
    st.subheader("Analisis Rasio Keuangan")
    styled_ratios = ratios.style.format(fmt_dict).background_gradient(cmap='Blues')
    st.dataframe(styled_ratios)
    
    # Visualisasi data
    st.subheader("Visualisasi Tren Keuangan")
    metrics = st.multiselect(
        "Pilih metrik untuk grafik",
        options=['revenue', 'net_income', 'ROE', 'ROA', 'DER', 'Current Ratio', 'Net Margin'],
        default=['revenue', 'ROE']
    )
    
    if metrics:
        plot_df = ratios[metrics].reset_index().melt('period', var_name='metric', value_name='value')
        fig = px.line(
            plot_df, 
            x='period', 
            y='value', 
            color='metric',
            markers=True,
            title="Tren Rasio Keuangan",
            labels={'value': 'Nilai', 'period': 'Periode'}
        )
        fig.update_layout(hovermode='x unified')
        st.plotly_chart(fig, use_container_width=True)
    
    # Ekspor hasil
    st.subheader("Ekspor Data")
    csv = ratios.reset_index().to_csv(index=False).encode('utf-8')
    st.download_button(
        "Download CSV",
        data=csv,
        file_name='financial_ratios.csv',
        mime='text/csv'
    )
