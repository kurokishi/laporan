"""
Streamlit Financial Statement Analyzer (IDX-oriented)
Features:
- Upload current financial Excel (.xls/.xlsx) from IDX
- Optionally upload previous financial Excel for comparison
- Parse sheets (tries to detect Neraca, Laba Rugi, Arus Kas)
- Compute common financial ratios
- Show comparison charts between uploaded file(s)
- Fetch stock price via yfinance for performance charts (user supplies ticker)
- Ready to deploy to Streamlit
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime, timedelta
import plotly.express as px
import yfinance as yf

# -------------------------
# Helpers: parsing utilities
# -------------------------
@st.cache_data
def read_excel_file(file_bytes: bytes):
    """
    Try to read excel into a dict of dataframes keyed by sheet name.
    Handles both xls and xlsx as best-effort.
    """
    try:
        xl = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
        return xl
    except Exception as e:
        # Try engine fallback
        try:
            xl = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, engine='xlrd')
            return xl
        except Exception as e2:
            st.error("Gagal membaca file Excel: " + str(e))
            return {}

def guess_sheet_names(sheet_dict):
    """
    Return mapping for 'neraca', 'laba_rugi', 'arus_kas' to actual sheet names if found.
    """
    name_map = {'neraca': None, 'laba_rugi': None, 'arus_kas': None}
    for name in sheet_dict.keys():
        n = name.lower()
        if 'neraca' in n or 'balance' in n:
            name_map['neraca'] = name
        if 'laba' in n or 'laba rugi' in n or 'income' in n or 'profit' in n:
            name_map['laba_rugi'] = name
        if 'arus kas' in n or 'cash' in n:
            name_map['arus_kas'] = name
    # fallback: try to assign first three sheets if none found
    if not any(name_map.values()):
        keys = list(sheet_dict.keys())
        for i, key in enumerate(keys[:3]):
            if i == 0: name_map['neraca'] = key
            if i == 1: name_map['laba_rugi'] = key
            if i == 2: name_map['arus_kas'] = key
    return name_map

def tidy_financial_df(df):
    """
    Normalize common Excel-structured financial sheet into a tidy form:
    - rows: account/line items
    - columns: years (or periods)
    Returns DataFrame with index = line item, columns = periods (string)
    """
    # drop completely empty rows/cols
    df = df.dropna(how='all').dropna(axis=1, how='all')
    # reset header if header is multi-row: find first row that contains year-like values
    header_row = None
    for i in range(min(5, len(df))):
        row = df.iloc[i].astype(str).str.lower().tolist()
        if any(re.search(r'20\d{2}', r) for r in row):
            header_row = i
            break
    if header_row is not None:
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row+1:]
    # Make first column the account name
    if df.columns.size >= 1:
        first_col = df.columns[0]
        df = df.rename(columns={first_col: 'Account'})
        df = df.set_index('Account', drop=True)
    # Try to convert other columns to numeric (remove commas, parentheses)
    def to_num(x):
        if pd.isna(x): return np.nan
        s = str(x).strip()
        s = s.replace('.', '') if s.count('.') > 1 and ',' in s else s  # heuristic
        s = s.replace(',', '')
        s = s.replace('(', '-').replace(')', '')
        s = s.replace('%', '')
        try:
            return float(s)
        except:
            return np.nan
    df = df.applymap(to_num)
    return df

# -------------------------
# Financial Ratios
# -------------------------
def compute_key_ratios(neraca_df, laba_df):
    """
    Compute a set of common ratios. Expects tidy dataframes with columns = periods (years).
    Returns DataFrame with ratios x periods.
    """
    periods = neraca_df.columns.tolist()
    ratios = {}
    for p in periods:
        try:
            total_assets = neraca_df.loc[[i for i in neraca_df.index if 'total assets' in i.lower() or 'jumlah aset' in i.lower() or 'total kekayaan' in i.lower()]].iloc[0].get(p)
        except:
            # fallback: take largest numeric in assets section
            total_assets = neraca_df[p].dropna().abs().max()
        # common lookups
        def find_first(df, keywords):
            for k in keywords:
                for idx in df.index:
                    if k in idx.lower():
                        return df.at[idx, p] if p in df.columns else np.nan
            return np.nan
        cash = find_first(neraca_df, ['kas', 'cash'])
        inventories = find_first(neraca_df, ['persediaan', 'inventory'])
        total_liab = find_first(neraca_df, ['total liabilities', 'jumlah kewajiban', 'total kewajiban', 'kewajiban']) 
        total_equity = find_first(neraca_df, ['total equity', 'jumlah ekuitas', 'ekuitas', 'total pemegang saham'])
        net_income = find_first(laba_df, ['profit for the year', 'net profit', 'laba bersih', 'net income', 'laba (rugi)'])
        revenue = find_first(laba_df, ['pendapatan usaha', 'revenue', 'penjualan', 'net sales'])
        gross_profit = find_first(laba_df, ['laba kotor', 'gross profit'])
        # basic ratios
        current_ratio = np.nan
        quick_ratio = np.nan
        debt_to_equity = np.nan
        roa = np.nan
        roe = np.nan
        gross_margin = np.nan
        net_margin = np.nan
        try:
            # crude current assets / current liabilities if present
            current_assets = find_first(neraca_df, ['current assets', 'aktiva lancar', 'aset lancar'])
            current_liab = find_first(neraca_df, ['current liabilities', 'kewajiban lancar', 'liabilities current'])
            if not pd.isna(current_assets) and not pd.isna(current_liab) and current_liab != 0:
                current_ratio = current_assets / current_liab
            if not pd.isna(cash) and not pd.isna(current_liab) and current_liab != 0:
                quick_ratio = cash / current_liab
            if not pd.isna(total_liab) and not pd.isna(total_equity) and total_equity != 0:
                debt_to_equity = total_liab / total_equity
            if not pd.isna(total_assets) and not pd.isna(net_income) and total_assets != 0:
                roa = net_income / total_assets
            if not pd.isna(total_equity) and not pd.isna(net_income) and total_equity != 0:
                roe = net_income / total_equity
            if not pd.isna(gross_profit) and not pd.isna(revenue) and revenue != 0:
                gross_margin = gross_profit / revenue
            if not pd.isna(net_income) and not pd.isna(revenue) and revenue != 0:
                net_margin = net_income / revenue
        except Exception as e:
            pass
        ratios[p] = {
            'Current Ratio': current_ratio,
            'Quick Ratio': quick_ratio,
            'Debt to Equity': debt_to_equity,
            'ROA': roa,
            'ROE': roe,
            'Gross Margin': gross_margin,
            'Net Margin': net_margin,
            'Revenue': revenue,
            'Net Income': net_income,
            'Total Assets': total_assets,
            'Total Equity': total_equity,
            'Cash': cash
        }
    ratios_df = pd.DataFrame(ratios).T
    return ratios_df

# -------------------------
# Main Streamlit App
# -------------------------
st.set_page_config(page_title="IDX Financial Analyzer", layout="wide")
st.title("IDX Financial Statement Analyzer — Streamlit")

st.sidebar.header("Upload & Settings")
uploaded_current = st.sidebar.file_uploader("Upload laporan keuangan (current) — .xls/.xlsx", type=['xls', 'xlsx'], key='cur')
uploaded_prev = st.sidebar.file_uploader("Upload laporan keuangan (previous) — optional", type=['xls','xlsx'], key='prev')
ticker = st.sidebar.text_input("Ticker saham (untuk price data, contoh 'AALI.JK')", value="")
price_period = st.sidebar.selectbox("Periode harga saham", options=['1mo','3mo','6mo','1y','2y','5y'], index=3)
price_interval = st.sidebar.selectbox("Interval harga", options=['1d','1wk','1mo'], index=0)
run_btn = st.sidebar.button("Analyze")

st.markdown("""
**Petunjuk singkat:** Unggah file laporan keuangan tahunan yang Anda download dari idx.com. 
Jika tersedia, unggah file laporan keuangan sebelumnya untuk grafik perbandingan. Isi ticker jika ingin melihat performa saham.
""")

if run_btn:
    if not uploaded_current:
        st.warning("Silakan upload file laporan keuangan current di sidebar terlebih dahulu.")
    else:
        # read current
        cur_bytes = uploaded_current.read()
        cur_sheets = read_excel_file(cur_bytes)
        if not cur_sheets:
            st.error("Tidak dapat membaca file saat ini. Pastikan file benar dan di-download dari idx.")
        else:
            st.success("File current terbaca. Menganalisis...")
            cur_map = guess_sheet_names(cur_sheets)
            st.write("Sheet terdeteksi:", cur_map)
            # tidy
            cur_neraca = tidy_financial_df(cur_sheets[cur_map['neraca']]) if cur_map['neraca'] in cur_sheets else pd.DataFrame()
            cur_laba = tidy_financial_df(cur_sheets[cur_map['laba_rugi']]) if cur_map['laba_rugi'] in cur_sheets else pd.DataFrame()
            cur_arus = tidy_financial_df(cur_sheets[cur_map['arus_kas']]) if cur_map['arus_kas'] in cur_sheets else pd.DataFrame()

            st.header("Preview: Current Financials")
            cols = st.columns(3)
            with cols[0]:
                st.subheader("Neraca (preview)")
                st.dataframe(cur_neraca.head(20))
            with cols[1]:
                st.subheader("Laba Rugi (preview)")
                st.dataframe(cur_laba.head(20))
            with cols[2]:
                st.subheader("Arus Kas (preview)")
                st.dataframe(cur_arus.head(20))

            # compute ratios
            ratios_cur = compute_key_ratios(cur_neraca, cur_laba)
            st.header("Rasio Keuangan (Current)")
            st.dataframe(ratios_cur.style.format("{:.4f}").applymap(lambda v: f"{v:.4f}" if pd.notna(v) else ""))

            # previous comparison if uploaded
            if uploaded_prev:
                prev_bytes = uploaded_prev.read()
                prev_sheets = read_excel_file(prev_bytes)
                prev_map = guess_sheet_names(prev_sheets)
                prev_neraca = tidy_financial_df(prev_sheets[prev_map['neraca']]) if prev_map['neraca'] in prev_sheets else pd.DataFrame()
                prev_laba = tidy_financial_df(prev_sheets[prev_map['laba_rugi']]) if prev_map['laba_rugi'] in prev_sheets else pd.DataFrame()
                ratios_prev = compute_key_ratios(prev_neraca, prev_laba)
                st.header("Perbandingan Rasio (Current vs Previous)")
                # align columns (periods)
                common_periods = list(dict.fromkeys(list(ratios_cur.index) + list(ratios_prev.index)))
                comp = pd.concat([
                    ratios_cur.assign(Source='Current'),
                    ratios_prev.assign(Source='Previous')
                ])
                comp = comp.reset_index().rename(columns={'index':'Period'}).set_index(['Source','Period'])
                st.dataframe(comp)

                # Plot comparisons for selected metrics
                st.subheader("Grafik Perbandingan Metrik")
                metric = st.selectbox("Pilih metrik untuk dibandingan", options=['Revenue','Net Income','ROA','ROE','Debt to Equity','Current Ratio','Gross Margin','Net Margin'])
                # prepare chart frame
                chart_df = []
                for s, df in [('Current', ratios_cur), ('Previous', ratios_prev)]:
                    tmp = df[[metric]].copy()
                    tmp = tmp.reset_index().rename(columns={'index':'Period'})
                    tmp['Source'] = s
                    chart_df.append(tmp)
                chart_df = pd.concat(chart_df, ignore_index=True)
                # convert Period strings to sort order if year-like
                try:
                    chart_df['Period_sort'] = chart_df['Period'].str.extract(r'(\d{4})')[0].astype(float)
                    chart_df = chart_df.sort_values(['Period_sort','Source'])
                except:
                    pass
                fig = px.bar(chart_df, x='Period', y=metric, color='Source', barmode='group', title=f'Perbandingan {metric}')
                st.plotly_chart(fig, use_container_width=True)

                # show side-by-side table difference percent
                try:
                    # pick matching most recent period name from current and previous
                    cur_latest = ratios_cur.index.max()
                    prev_latest = ratios_prev.index.max()
                    summary_diff = pd.DataFrame({
                        'Metric': ratios_cur.columns,
                        'Current': ratios_cur.loc[cur_latest].values,
                        'Previous': ratios_prev.loc[prev_latest].values
                    })
                    summary_diff['Pct Change'] = (summary_diff['Current'] - summary_diff['Previous']) / summary_diff['Previous'].replace({0:np.nan})
                    st.subheader(f"Perubahan antar periode (Current: {cur_latest} vs Previous: {prev_latest})")
                    st.dataframe(summary_diff.set_index('Metric').style.format("{:.4f}"))
                except Exception as e:
                    st.info("Tidak dapat menghitung ringkasan perbedaan otomatis: " + str(e))

            # stock price performance
            if ticker:
                st.header(f"Performa Harga: {ticker}")
                try:
                    yf_tkr = yf.Ticker(ticker)
                    hist = yf_tkr.history(period=price_period, interval=price_interval)
                    if hist.empty:
                        st.warning("Tidak ada data harga untuk ticker ini atau ticker bukan format Yahoo (contoh: 'AALI.JK').")
                    else:
                        st.line_chart(hist['Close'])
                        # small stats
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Harga sekarang", f"{hist['Close'][-1]:.2f}")
                        col2.metric("Return periode", f"{(hist['Close'][-1]/hist['Close'][0]-1)*100:.2f}%")
                        col3.metric("Volatilitas (std dev)", f"{hist['Close'].pct_change().std()*100:.2f}%")
                        # overlay moving averages
                        ma_df = hist[['Close']].copy()
                        ma_df['MA20'] = ma_df['Close'].rolling(20).mean()
                        ma_df['MA50'] = ma_df['Close'].rolling(50).mean()
                        fig2 = px.line(ma_df, y=['Close','MA20','MA50'], title='Harga dan Moving Averages')
                        st.plotly_chart(fig2, use_container_width=True)
                        # Basic valuation: P/E using last net income (EPS approximation)
                        try:
                            # EPS approx = Net Income / outstanding shares - we don't have shares so skip or ask user
                            st.info("Perhitungan valuasi (P/E) memerlukan jumlah saham beredar / EPS. Anda bisa tambahkan nilai EPS di sidebar.")
                            eps_input = st.sidebar.number_input("Masukkan EPS terakhir (opsional)", value=float(0.0), format="%.6f")
                            if eps_input and eps_input > 0:
                                pe = hist['Close'][-1] / eps_input
                                st.write(f"Estimated P/E (price / EPS): {pe:.2f}")
                        except Exception as e:
                            st.write("Valuasi dasar gagal dihitung:", e)
                except Exception as e:
                    st.error("Gagal mengambil data harga: " + str(e))

            # final suggestions
            st.header("Kesimpulan & Saran Fitur Tambahan")
            st.markdown("""
            - Cek preview sheet hasil parsing. Struktur file IDX kadang berbeda — jika parsing salah, unggah file contoh lain atau koreksi manual.
            - Fitur tambahan yang disarankan (lihat juga bagian terpisah di bawah):
                1. Otomatis deteksi dan mapping akun (mapping labels IDX ke standard accounts).
                2. Ambil laporan konsolidasi & non-konsolidasi otomatis bila tersedia.
                3. Integrasi database saham Indonesia (untuk mengambil jumlah saham beredar, EPS, dividen historis).
                4. Screener batch: upload banyak ticker & file, generate ranking berdasarkan rasio custom.
                5. Export otomatis ke PDF / PPT laporan analisis untuk presentasi.
                6. Backtest strategi investasi sederhana berbasis rasio fundamental + momentum.
            """)
            st.success("Analisis selesai. Sesuaikan parsing jika ada baris/kolom yang tidak terbaca dengan benar.")
