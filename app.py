import streamlit as st
import pandas as pd
import numpy as np
import re
import tempfile
import os
import camelot
import pdfplumber
import plotly.express as px
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Konfigurasi halaman Streamlit
st.set_page_config(page_title="IDX Financial Analyzer", layout="wide")
st.title("📊 IDX Financial Analyzer")
st.caption("Analisis laporan keuangan seperti Stockbit - Unggah PDF atau Excel")

# ==================== FUNGSI UTILITAS ====================
def clean_number(x):
    """Membersihkan dan mengonversi string angka ke float"""
    # Handle pandas Series
    if isinstance(x, pd.Series):
        return x.map(clean_number)
    
    # Handle missing values
    if pd.isna(x) or x is None:
        return np.nan
    
    # Handle numeric types
    if isinstance(x, (int, float)):
        return x
    
    s = str(x).strip()
    
    # Handle empty values
    if s in ["-", "—", "", "n/a", "N/A", "NA"]:
        return np.nan
    
    # Handle negative numbers in parentheses
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    
    # Remove non-numeric characters
    s = re.sub(r"[^\d\.\-]", "", s)
    
    # Handle cases like "1.000.000" -> "1000000"
    if s.count('.') > 1:
        s = s.replace('.', '')
    
    try:
        val = float(s)
        return -val if neg else val
    except Exception as e:
        logger.warning(f"Gagal mengonversi '{x}' ke angka: {e}")
        return np.nan

# Keyword untuk deteksi jenis laporan keuangan
BALANCE_KEYWORDS = ['laporan posisi keuangan', 'neraca', 'statement of financial position', 
                    'total assets', 'jumlah aset', 'aset', 'aktiva', 'kewajiban', 'ekuitas']
INCOME_KEYWORDS = ['laba rugi', 'income statement', 'statement of profit', 
                   'penjualan', 'revenue', 'sales', 'pendapatan', 'laba bersih', 'beban']
CASH_KEYWORDS = ['arus kas', 'cash flows', 'statement of cash flows', 
                 'kas dan setara kas', 'arus kas dari aktivitas', 'aliran kas']

def text_contains_any(text, keywords):
    """Cek apakah teks mengandung salah satu keyword"""
    if not isinstance(text, str):
        text = str(text)
    text = text.lower()
    return any(kw in text for kw in keywords)

# ==================== EKSTRAKSI TABEL ====================
def table_to_period_df(df_raw):
    """Konversi tabel mentah menjadi DataFrame periodik"""
    try:
        # Normalisasi dan hapus baris/kolom kosong
        df = df_raw.copy()
        df = df.map(lambda x: str(x).strip() if pd.notna(x) else x)
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
        
        # Hapus kolom yang seluruhnya kosong
        body = body.dropna(how='all', axis=1)
        if body.shape[1] < 2:
            return pd.DataFrame()
        
        # Transformasi ke format long
        rows = []
        account_col = body.columns[0]
        
        for idx, row in body.iterrows():
            account = str(row[account_col]).strip()
            if not account or account.lower() in ['', 'nan', 'none']:
                continue
                
            for col in body.columns[1:]:
                period = str(col).strip()
                if not period or period.lower() in ['', 'nan', 'none']:
                    continue
                    
                raw_value = row[col]
                # Pastikan kita hanya mengambil nilai tunggal
                if isinstance(raw_value, pd.Series):
                    raw_value = raw_value.iloc[0] if not raw_value.empty else np.nan
                
                value = clean_number(raw_value)
                
                if pd.notna(value):
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
    
    except Exception as e:
        logger.error(f"Error dalam table_to_period_df: {str(e)}")
        return pd.DataFrame()

def extract_tables_from_pdf(file_bytes):
    """Ekstrak tabel dari PDF menggunakan Camelot dan PDFPlumber"""
    tables = []
    
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        
        # Coba ekstrak dengan Camelot
        try:
            cam_tables = camelot.read_pdf(tmp_path, pages='all', flavor='stream', suppress_stdout=True)
            tables = [t.df for t in cam_tables]
            logger.info(f"Camelot berhasil mengekstrak {len(tables)} tabel")
        except Exception as e:
            logger.warning(f"Camelot error: {str(e)}")
        
        # Jika Camelot gagal, coba dengan PDFPlumber
        if not tables:
            try:
                with pdfplumber.open(tmp_path) as pdf:
                    for page in pdf.pages:
                        for table in page.extract_tables():
                            if table and len(table) > 1:
                                df = pd.DataFrame(table[1:], columns=table[0])
                                tables.append(df)
                logger.info(f"PDFPlumber berhasil mengekstrak {len(tables)} tabel")
            except Exception as e:
                logger.error(f"PDFPlumber error: {str(e)}")
    except Exception as e:
        logger.error(f"Error ekstraksi PDF: {str(e)}")
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
    
    return tables

def detect_statement_tables(raw_tables):
    """Identifikasi jenis laporan keuangan berdasarkan keyword"""
    mapping = {'income': [], 'balance': [], 'cash': []}
    
    for idx, df in enumerate(raw_tables):
        try:
            # Gabungkan konten seluruh tabel untuk deteksi keyword
            sample_text = " ".join(df.astype(str).values.flatten()).lower()
            
            # Deteksi neraca
            if any(kw in sample_text for kw in BALANCE_KEYWORDS):
                mapping['balance'].append(idx)
            
            # Deteksi laporan laba rugi
            if any(kw in sample_text for kw in INCOME_KEYWORDS):
                mapping['income'].append(idx)
            
            # Deteksi laporan arus kas
            if any(kw in sample_text for kw in CASH_KEYWORDS):
                mapping['cash'].append(idx)
        except:
            continue
    
    return mapping

# ==================== PERHITUNGAN RASIO ====================
def compute_basic_ratios(income_df, balance_df):
    """Hitung rasio keuangan utama dari data laporan"""
    try:
        # Gabungkan semua periode yang tersedia
        periods = sorted(set(income_df.index) | set(balance_df.index), key=str)
        rows = []
        
        for period in periods:
            # Fungsi pencarian kolom dengan keyword
            def find_account(df, keywords):
                if df is None or df.empty:
                    return None
                for col in df.columns:
                    col_str = str(col).lower()
                    if any(kw in col_str for kw in keywords):
                        return col
                return None
            
            # Cari akun-akun utama
            revenue_acc = find_account(income_df, ['pendapatan', 'revenue', 'sales', 'penjualan'])
            net_income_acc = find_account(income_df, ['laba bersih', 'net income', 'laba tahun berjalan', 'profit for the year'])
            total_assets_acc = find_account(balance_df, ['total aset', 'total assets', 'jumlah aset', 'aktiva'])
            total_equity_acc = find_account(balance_df, ['total ekuitas', 'total equity', 'jumlah ekuitas', 'ekuitas'])
            total_liab_acc = find_account(balance_df, ['total liabilitas', 'total liabilities', 'jumlah kewajiban', 'liabilitas'])
            current_assets_acc = find_account(balance_df, ['aset lancar', 'current assets'])
            current_liab_acc = find_account(balance_df, ['liabilitas jangka pendek', 'current liabilities'])
            
            # Ekstrak nilai dengan penanganan error
            def safe_get(df, period, account):
                try:
                    if account and period in df.index:
                        value = df.at[period, account]
                        if isinstance(value, pd.Series):
                            return value.iloc[0]
                        return value
                except:
                    pass
                return np.nan
            
            revenue = safe_get(income_df, period, revenue_acc)
            net_income = safe_get(income_df, period, net_income_acc)
            total_assets = safe_get(balance_df, period, total_assets_acc)
            total_equity = safe_get(balance_df, period, total_equity_acc)
            total_liab = safe_get(balance_df, period, total_liab_acc)
            current_assets = safe_get(balance_df, period, current_assets_acc)
            current_liab = safe_get(balance_df, period, current_liab_acc)
            
            # Hitung rasio keuangan
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
                'current_ratio': cr,
                'ROE': roe,
                'ROA': roa,
                'DER': der,
                'net_margin': net_margin
            })
        
        return pd.DataFrame(rows).set_index('period')
    
    except Exception as e:
        logger.error(f"Error menghitung rasio: {str(e)}")
        return pd.DataFrame()

# ==================== ANTARMUKA PENGGUNA ====================
uploaded = st.file_uploader("Unggah Laporan Keuangan (PDF atau Excel)", type=['pdf', 'xlsx', 'xls'])

if not uploaded:
    st.info("Silakan unggah file PDF atau Excel laporan keuangan")
    st.stop()

# Pemrosesan file PDF
if uploaded.name.lower().endswith('.pdf'):
    with st.spinner("Memproses PDF..."):
        raw_tables = extract_tables_from_pdf(uploaded.read())
    
    if not raw_tables:
        st.error("⚠️ Gagal mengekstrak tabel dari PDF. Format mungkin tidak didukung.")
        st.stop()
    
    st.success(f"✅ Berhasil ekstrak {len(raw_tables)} tabel")
    mapping = detect_statement_tables(raw_tables)
    
    # Tampilkan preview tabel untuk membantu pemilihan
    st.subheader("Pilih Tabel yang Sesuai")
    st.info("Pilih tabel yang sesuai berdasarkan preview di bawah. Sistem telah mencoba mendeteksi otomatis.")
    
    # Tampilkan preview tabel
    preview_idx = st.selectbox("Preview Tabel", range(len(raw_tables)), format_func=lambda x: f"Tabel {x+1}")
    st.dataframe(raw_tables[preview_idx].head(10))
    
    # Tampilkan pilihan tabel
    col1, col2, col3 = st.columns(3)
    
    # Fungsi untuk mendapatkan indeks default
    def get_default_idx(mapping_key):
        return mapping[mapping_key][0] if mapping.get(mapping_key) else 0
    
    bal_idx = col1.selectbox(
        "Tabel Neraca",
        options=range(len(raw_tables)),
        index=get_default_idx('balance'),
        format_func=lambda x: f"Tabel {x+1}"
    )
    
    inc_idx = col2.selectbox(
        "Tabel Laba Rugi",
        options=range(len(raw_tables)),
        index=get_default_idx('income')
    )
    
    cash_idx = col3.selectbox(
        "Tabel Arus Kas",
        options=range(len(raw_tables)),
        index=get_default_idx('cash')
    )
    
    # Konversi tabel yang dipilih
    income_df = table_to_period_df(raw_tables[inc_idx])
    balance_df = table_to_period_df(raw_tables[bal_idx])
    cash_df = table_to_period_df(raw_tables[cash_idx])

# Pemrosesan file Excel
else:
    with st.spinner("Memproses Excel..."):
        try:
            xls = pd.read_excel(uploaded, sheet_name=None, header=None)
            sheet_names = list(xls.keys())
            
            # Deteksi sheet berdasarkan pola nama
            balance_sheet = next((s for s in sheet_names if any(kw in s.lower() for kw in BALANCE_KEYWORDS)), None)
            income_sheet = next((s for s in sheet_names if any(kw in s.lower() for kw in INCOME_KEYWORDS)), None)
            cash_sheet = next((s for s in sheet_names if any(kw in s.lower() for kw in CASH_KEYWORDS)), None)
            
            # Konversi sheet yang terdeteksi
            income_df = table_to_period_df(xls[income_sheet]) if income_sheet else pd.DataFrame()
            balance_df = table_to_period_df(xls[balance_sheet]) if balance_sheet else pd.DataFrame()
            cash_df = table_to_period_df(xls[cash_sheet]) if cash_sheet else pd.DataFrame()
            
            st.success(f"✅ Berhasil memproses {len(sheet_names)} sheet Excel")
            
        except Exception as e:
            st.error(f"❌ Error processing Excel: {str(e)}")
            st.stop()

# ==================== TAMPILAN HASIL ====================
if income_df.empty or balance_df.empty:
    st.warning("📋 Data tidak lengkap untuk menghitung rasio. Periksa pemilihan tabel.")
    
    if not income_df.empty:
        st.subheader("Preview Tabel Laba Rugi")
        st.dataframe(income_df.head())
    
    if not balance_df.empty:
        st.subheader("Preview Tabel Neraca")
        st.dataframe(balance_df.head())
    
else:
    with st.spinner("🧮 Menghitung rasio keuangan..."):
        ratios = compute_basic_ratios(income_df, balance_df)
    
    if ratios.empty:
        st.error("❌ Gagal menghitung rasio. Format laporan mungkin tidak sesuai.")
    else:
        # Format tampilan tabel
        st.subheader("📈 Analisis Rasio Keuangan")
        
        # Buat tab untuk berbagai jenis rasio
        tab1, tab2, tab3 = st.tabs(["Profitabilitas", "Solvabilitas", "Likuiditas"])
        
        with tab1:
            st.markdown("### Rasio Profitabilitas")
            profit_cols = ['revenue', 'net_income', 'ROE', 'ROA', 'net_margin']
            st.dataframe(ratios[profit_cols].style.format({
                'revenue': 'Rp {:,.2f}',
                'net_income': 'Rp {:,.2f}',
                'ROE': '{:.2%}',
                'ROA': '{:.2%}',
                'net_margin': '{:.2%}'
            }))
        
        with tab2:
            st.markdown("### Rasio Solvabilitas")
            solv_cols = ['total_assets', 'total_equity', 'total_liabilities', 'DER']
            st.dataframe(ratios[solv_cols].style.format({
                'total_assets': 'Rp {:,.2f}',
                'total_equity': 'Rp {:,.2f}',
                'total_liabilities': 'Rp {:,.2f}',
                'DER': '{:.2f}x'
            }))
        
        with tab3:
            st.markdown("### Rasio Likuiditas")
            liq_cols = ['current_ratio']
            st.dataframe(ratios[liq_cols].style.format({
                'current_ratio': '{:.2f}x'
            }))
        
        # Visualisasi data
        st.subheader("📊 Tren Rasio Keuangan")
        
        # Pilih metrik untuk visualisasi
        default_metrics = ['revenue', 'net_income', 'ROE']
        available_metrics = [col for col in ratios.columns if col not in ['period']]
        
        metrics = st.multiselect(
            "Pilih metrik untuk ditampilkan",
            options=available_metrics,
            default=default_metrics
        )
        
        if metrics:
            plot_df = ratios[metrics].reset_index().melt('period', var_name='metric', value_name='value')
            
            # Buat dua jenis visualisasi: line chart dan bar chart
            fig_line = px.line(
                plot_df, 
                x='period', 
                y='value', 
                color='metric',
                markers=True,
                title="Tren Rasio Keuangan",
                labels={'value': 'Nilai', 'period': 'Periode'}
            )
            
            fig_bar = px.bar(
                plot_df,
                x='period',
                y='value',
                color='metric',
                barmode='group',
                title="Perbandingan Rasio per Periode"
            )
            
            st.plotly_chart(fig_line, use_container_width=True)
            st.plotly_chart(fig_bar, use_container_width=True)
        
        # Analisis tambahan seperti Stockbit
        st.subheader("💡 Analisis Fundamental")
        
        if 'ROE' in ratios.columns and 'DER' in ratios.columns:
            latest_roe = ratios['ROE'].iloc[-1]
            latest_der = ratios['DER'].iloc[-1]
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("Return on Equity (ROE)", f"{latest_roe:.2%}", 
                         help="Mengukur efisiensi penggunaan ekuitas pemegang saham")
                
                if latest_roe > 0.15:
                    st.success("✅ ROE sangat baik (>15%)")
                elif latest_roe > 0.10:
                    st.info("ℹ️ ROE cukup baik (10-15%)")
                else:
                    st.warning("⚠️ ROE di bawah standar (<10%)")
            
            with col2:
                st.metric("Debt to Equity Ratio (DER)", f"{latest_der:.2f}x", 
                         help="Mengukur tingkat leverage perusahaan")
                
                if latest_der < 1.0:
                    st.success("✅ DER rendah (risiko kecil)")
                elif latest_der < 2.0:
                    st.info("ℹ️ DER moderat")
                else:
                    st.warning("⚠️ DER tinggi (risiko besar)")
        
        # Ekspor hasil
        st.subheader("💾 Ekspor Data")
        csv = ratios.reset_index().to_csv(index=False).encode('utf-8')
        st.download_button(
            "Download CSV",
            data=csv,
            file_name='financial_analysis.csv',
            mime='text/csv'
        )
