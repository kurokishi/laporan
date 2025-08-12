import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
from pathlib import Path
import warnings
import streamlit as st
import io
import base64

warnings.filterwarnings('ignore')

class IDXFinancialAnalyzer:
    def __init__(self, file_path, company_name=None):
        self.file_path = Path(file_path)
        self.company_name = company_name or self.file_path.stem
        self.data = None
        self.processed_data = None
        self.df_analysis = None
        self.financial_ratios = None
        
    def load_data(self, sheet_name=0):
        try:
            if self.file_path.suffix.lower() == '.xls':
                self.data = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='xlrd')
            else:
                self.data = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='openpyxl')
            
            st.success(f"✅ Data berhasil dimuat: {self.data.shape[0]} baris, {self.data.shape[1]} kolom")
            return True
            
        except Exception as e:
            st.error(f"❌ Error loading file: {e}")
            return False
    
    def preview_data(self, rows=10):
        if self.data is None:
            st.warning("❌ Data belum dimuat. Jalankan load_data() terlebih dahulu.")
            return
        
        st.subheader("📊 Preview Data")
        st.write(f"**Shape:** {self.data.shape}")
        st.write(f"**Columns:** {list(self.data.columns)}")
        
        st.write("\n**Data Preview:**")
        st.dataframe(self.data.head(rows))
        
        year_cols = [col for col in self.data.columns if str(col).isdigit() and len(str(col)) == 4]
        st.info(f"📅 Detected year columns: {year_cols}")
    
    def standardize_data(self, metric_column='Keterangan', auto_detect=True):
        if self.data is None:
            st.warning("❌ Data belum dimuat.")
            return False
        
        df = self.data.copy()
        
        if metric_column not in df.columns:
            possible_cols = [col for col in df.columns if any(keyword in str(col).lower() 
                           for keyword in ['keterangan', 'description', 'item', 'metric'])]
            if possible_cols:
                metric_column = possible_cols[0]
                st.info(f"📝 Using metric column: {metric_column}")
            else:
                metric_column = df.columns[0]
                st.warning(f"⚠️ Using first column as metric: {metric_column}")
        
        if auto_detect:
            year_cols = [col for col in df.columns if str(col).isdigit() and len(str(col)) == 4]
            year_cols.sort()
        else:
            year_cols = [col for col in df.columns if col != metric_column]
        
        if len(year_cols) < 2:
            st.warning(f"⚠️ Hanya ditemukan {len(year_cols)} kolom tahun. Minimal 2 tahun diperlukan.")
            return False
        
        for col in year_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        df = df.dropna(subset=year_cols, how='all')
        
        metric_mapping = {
            'total aset': 'Total Assets',
            'total aktiva': 'Total Assets', 
            'kas dan setara kas': 'Cash & Equivalents',
            'kas dan bank': 'Cash & Equivalents',
            'piutang usaha': 'Accounts Receivable',
            'persediaan': 'Inventory',
            'aset lancar': 'Current Assets',
            'total liabilitas': 'Total Liabilities',
            'total kewajiban': 'Total Liabilities',
            'utang usaha': 'Accounts Payable',
            'liabilitas jangka pendek': 'Current Liabilities',
            'kewajiban lancar': 'Current Liabilities',
            'utang bank jangka panjang': 'Long-term Debt',
            'liabilitas jangka panjang': 'Long-term Debt',
            'total ekuitas': 'Total Equity',
            'modal disetor': 'Paid-in Capital',
            'pendapatan': 'Revenue',
            'penjualan': 'Revenue', 
            'beban pokok penjualan': 'Cost of Goods Sold',
            'laba kotor': 'Gross Profit',
            'laba usaha': 'Operating Profit',
            'laba sebelum pajak': 'Profit Before Tax',
            'laba bersih': 'Net Profit',
            'laba tahun berjalan': 'Net Profit',
            'dividen': 'Dividends Paid',
            'dividen dibayar': 'Dividends Paid'
        }
        
        df['Standard_Metric'] = df[metric_column].str.lower().str.strip()
        for key, value in metric_mapping.items():
            df.loc[df['Standard_Metric'].str.contains(key, na=False), 'Standard_Metric'] = value
        
        self.processed_data = {
            'metrics': df['Standard_Metric'].tolist(),
            'original_metrics': df[metric_column].tolist(),
            'years': year_cols,
            'data': df[year_cols].values
        }
        
        structured_data = []
        for i, metric in enumerate(self.processed_data['metrics']):
            row = {'Metric': metric, 'Original': self.processed_data['original_metrics'][i]}
            for j, year in enumerate(year_cols):
                row[str(year)] = self.processed_data['data'][i][j]
            structured_data.append(row)
        
        self.df_analysis = pd.DataFrame(structured_data)
        
        st.success(f"✅ Data terstandarisasi: {len(self.df_analysis)} metrics, {len(year_cols)} years")
        return True
    
    def calculate_financial_ratios(self):
        if self.df_analysis is None:
            st.warning("❌ Data belum diproses.")
            return None
        
        df = self.df_analysis.copy()
        years = [col for col in df.columns if str(col).isdigit()]
        
        ratios = {}
        
        for year in years:
            year_ratios = {}
            
            def get_value(metric_name):
                mask = df['Metric'].str.contains(metric_name, case=False, na=False)
                values = df.loc[mask, year]
                return values.iloc[0] if len(values) > 0 and pd.notna(values.iloc[0]) else 0
            
            current_assets = get_value('Current Assets')
            cash = get_value('Cash & Equivalents') 
            current_liab = get_value('Current Liabilities')
            
            if current_liab != 0:
                year_ratios['Current Ratio'] = current_assets / current_liab
                year_ratios['Quick Ratio'] = (cash) / current_liab
            
            revenue = get_value('Revenue')
            gross_profit = get_value('Gross Profit')
            net_profit = get_value('Net Profit')
            total_assets = get_value('Total Assets')
            total_equity = get_value('Total Equity')
            
            if revenue != 0:
                year_ratios['Gross Margin %'] = (gross_profit / revenue) * 100
                year_ratios['Net Margin %'] = (net_profit / revenue) * 100
            
            if total_assets != 0:
                year_ratios['ROA %'] = (net_profit / total_assets) * 100
            
            if total_equity != 0:
                year_ratios['ROE %'] = (net_profit / total_equity) * 100
            
            total_liab = get_value('Total Liabilities')
            long_term_debt = get_value('Long-term Debt')
            
            if total_assets != 0:
                year_ratios['Debt to Asset %'] = (total_liab / total_assets) * 100
            
            if total_equity != 0:
                year_ratios['Debt to Equity %'] = (total_liab / total_equity) * 100
            
            ratios[year] = year_ratios
        
        self.financial_ratios = pd.DataFrame(ratios).T
        return self.financial_ratios
    
    def create_visualizations(self):
        if self.df_analysis is None:
            st.warning("❌ Data belum diproses.")
            return
        
        plt.style.use('seaborn-v0_8')
        sns.set_palette("husl")
        
        df = self.df_analysis.copy()
        years = [col for col in df.columns if str(col).isdigit()]
        
        st.subheader("📈 Financial Visualizations")
        
        # Tab layout
        tab1, tab2, tab3, tab4 = st.tabs([
            "Key Metrics", 
            "Profitability", 
            "Liquidity & Debt", 
            "Ratios"
        ])
        
        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                st.write("### Key Financial Metrics")
                key_metrics = ['Total Assets', 'Revenue', 'Net Profit']
                key_data = df[df['Metric'].isin(key_metrics)]
                if not key_data.empty:
                    key_data_pivot = key_data.set_index('Metric')[years]
                    st.bar_chart(key_data_pivot.T)
            
            with col2:
                st.write("### Asset Composition")
                latest_year = years[-1]
                asset_metrics = ['Current Assets', 'Total Assets']
                asset_data = df[df['Metric'].isin(asset_metrics)]
                if not asset_data.empty and len(asset_data) > 0:
                    asset_values = asset_data[latest_year].values
                    asset_labels = asset_data['Metric'].values
                    fig, ax = plt.subplots(figsize=(5, 5))
                    ax.pie(asset_values, labels=asset_labels, autopct='%1.1f%%')
                    st.pyplot(fig)
        
        with tab2:
            col1, col2 = st.columns(2)
            with col1:
                st.write("### Profitability Trend")
                profit_metrics = ['Gross Profit', 'Operating Profit', 'Net Profit']
                profit_data = df[df['Metric'].isin(profit_metrics)]
                if not profit_data.empty:
                    st.line_chart(profit_data.set_index('Metric')[years].T)
            
            with col2:
                st.write("### Profitability Growth")
                if len(years) >= 2:
                    growth_metrics = ['Revenue', 'Net Profit', 'Total Assets']
                    growth_data = []
                    for metric in growth_metrics:
                        metric_row = df[df['Metric'] == metric]
                        if not metric_row.empty:
                            values = metric_row[years].iloc[0]
                            if len(values) >= 2 and values.iloc[0] != 0:
                                growth_rate = ((values.iloc[-1] - values.iloc[0]) / values.iloc[0]) * 100
                                growth_data.append({'Metric': metric, 'Growth %': growth_rate})
                    if growth_data:
                        growth_df = pd.DataFrame(growth_data)
                        st.bar_chart(growth_df.set_index('Metric'))
        
        with tab3:
            col1, col2 = st.columns(2)
            with col1:
                st.write("### Liquidity Analysis")
                liquidity_metrics = ['Cash & Equivalents', 'Current Liabilities']
                liquidity_data = df[df['Metric'].isin(liquidity_metrics)]
                if not liquidity_data.empty:
                    st.bar_chart(liquidity_data.set_index('Metric')[years].T)
            
            with col2:
                st.write("### Debt Structure")
                debt_metrics = ['Current Liabilities', 'Long-term Debt']
                debt_data = df[df['Metric'].isin(debt_metrics)]
                if not debt_data.empty:
                    debt_data_pivot = debt_data.set_index('Metric')[years].T
                    debt_data_pivot['Long-term Debt'] = debt_data_pivot.get('Long-term Debt', 0)
                    debt_data_pivot['Current Liabilities'] = debt_data_pivot.get('Current Liabilities', 0)
                    st.bar_chart(debt_data_pivot)
        
        with tab4:
            if hasattr(self, 'financial_ratios') and self.financial_ratios is not None:
                col1, col2 = st.columns(2)
                with col1:
                    st.write("### Liquidity Ratios")
                    ratio_cols = ['Current Ratio', 'Quick Ratio']
                    available_ratios = [col for col in ratio_cols if col in self.financial_ratios.columns]
                    if available_ratios:
                        st.bar_chart(self.financial_ratios[available_ratios])
                
                with col2:
                    st.write("### Profitability Margins")
                    margin_cols = ['Gross Margin %', 'Net Margin %']
                    available_margins = [col for col in margin_cols if col in self.financial_ratios.columns]
                    if available_margins:
                        st.line_chart(self.financial_ratios[available_margins])
                
                col3, col4 = st.columns(2)
                with col3:
                    st.write("### Return Ratios")
                    return_cols = ['ROA %', 'ROE %']
                    available_returns = [col for col in return_cols if col in self.financial_ratios.columns]
                    if available_returns:
                        st.bar_chart(self.financial_ratios[available_returns])
                
                with col4:
                    st.write("### Leverage Ratios")
                    leverage_cols = ['Debt to Asset %', 'Debt to Equity %']
                    available_leverage = [col for col in leverage_cols if col in self.financial_ratios.columns]
                    if available_leverage:
                        st.bar_chart(self.financial_ratios[available_leverage])
    
    def generate_report(self):
        if self.df_analysis is None:
            st.warning("❌ Data belum diproses.")
            return
        
        st.subheader(f"📈 Financial Analysis Report - {self.company_name}")
        years = [col for col in self.df_analysis.columns if str(col).isdigit()]
        
        # Key Metrics Summary
        st.write("### 📊 Key Financial Metrics")
        display_metrics = ['Total Assets', 'Revenue', 'Gross Profit', 'Net Profit', 
                         'Cash & Equivalents', 'Current Liabilities', 'Total Equity']
        
        summary_data = []
        for metric in display_metrics:
            metric_row = self.df_analysis[self.df_analysis['Metric'] == metric]
            if not metric_row.empty:
                row_data = {'Metric': metric}
                for year in years:
                    value = metric_row[year].iloc[0]
                    row_data[year] = f"{value:,.0f}" if pd.notna(value) else "N/A"
                
                if len(years) >= 2:
                    old_val = metric_row[years[0]].iloc[0]
                    new_val = metric_row[years[-1]].iloc[0]
                    if pd.notna(old_val) and pd.notna(new_val) and old_val != 0:
                        change = ((new_val - old_val) / old_val) * 100
                        row_data['Change %'] = f"{change:+.1f}%"
                    else:
                        row_data['Change %'] = "N/A"
                
                summary_data.append(row_data)
        
        if summary_data:
            st.dataframe(pd.DataFrame(summary_data))
        
        # Financial Ratios
        if hasattr(self, 'financial_ratios') and self.financial_ratios is not None:
            st.write("### 📈 Financial Ratios")
            st.dataframe(self.financial_ratios.style.format("{:.2f}"))
        
        # Risk Analysis
        st.write("### ⚠️ Risk Indicators")
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Debt Analysis**")
            debt_data = self.df_analysis[self.df_analysis['Metric'].str.contains('Debt|Liabilities', case=False, na=False)]
            if not debt_data.empty:
                latest_year = years[-1]
                total_debt = debt_data[latest_year].sum()
                st.metric(f"Total Debt ({latest_year})", f"{total_debt:,.0f}")
        
        with col2:
            st.write("**Liquidity Analysis**")
            cash_row = self.df_analysis[self.df_analysis['Metric'] == 'Cash & Equivalents']
            current_liab_row = self.df_analysis[self.df_analysis['Metric'] == 'Current Liabilities']
            
            if not cash_row.empty and not current_liab_row.empty:
                latest_year = years[-1]
                cash = cash_row[latest_year].iloc[0] if not cash_row.empty else 0
                liab = current_liab_row[latest_year].iloc[0] if not current_liab_row.empty else 0
                if liab != 0:
                    coverage = cash / liab
                    status = "✅ Good" if coverage > 1 else "🔄 Adequate" if coverage > 0.5 else "❌ Risk"
                    st.metric(f"Cash Coverage ({latest_year})", f"{coverage:.2f}x", status)
        
        # Growth Analysis
        st.write("### 📈 Growth Analysis")
        if len(years) >= 2:
            growth_metrics = ['Revenue', 'Net Profit', 'Total Assets', 'Total Equity']
            growth_cols = st.columns(len(growth_metrics))
            
            for i, metric in enumerate(growth_metrics):
                metric_row = self.df_analysis[self.df_analysis['Metric'] == metric]
                if not metric_row.empty:
                    old_val = metric_row[years[0]].iloc[0]
                    new_val = metric_row[years[-1]].iloc[0]
                    if pd.notna(old_val) and pd.notna(new_val) and old_val != 0:
                        growth = ((new_val - old_val) / old_val) * 100
                        trend = "📈" if growth > 0 else "📉"
                        with growth_cols[i]:
                            st.metric(f"{metric} Growth", f"{growth:+.1f}%", trend)
        
        # Investment Signals
        st.write("### 🎯 Investment Signals")
        signals = []
        
        # Profitability signal
        net_profit_row = self.df_analysis[self.df_analysis['Metric'] == 'Net Profit']
        if not net_profit_row.empty and len(years) >= 2:
            old_profit = net_profit_row[years[0]].iloc[0]
            new_profit = net_profit_row[years[-1]].iloc[0]
            if pd.notna(old_profit) and pd.notna(new_profit):
                if new_profit > old_profit:
                    signals.append("✅ Profitability improving")
                else:
                    signals.append("⚠️ Profitability declining")
        
        # Liquidity signal
        if hasattr(self, 'financial_ratios') and 'Current Ratio' in self.financial_ratios.columns:
            latest_cr = self.financial_ratios['Current Ratio'].iloc[-1]
            if latest_cr > 1.5:
                signals.append("✅ Strong liquidity position")
            elif latest_cr > 1.0:
                signals.append("🔄 Adequate liquidity")
            else:
                signals.append("❌ Liquidity concerns")
        
        # Growth signal
        revenue_row = self.df_analysis[self.df_analysis['Metric'] == 'Revenue']
        if not revenue_row.empty and len(years) >= 2:
            old_rev = revenue_row[years[0]].iloc[0]
            new_rev = revenue_row[years[-1]].iloc[0]
            if pd.notna(old_rev) and pd.notna(new_rev) and old_rev != 0:
                growth = ((new_rev - old_rev) / old_rev) * 100
                if growth > 10:
                    signals.append("🚀 Strong revenue growth")
                elif growth > 0:
                    signals.append("📈 Positive revenue growth")
                else:
                    signals.append("📉 Revenue declining")
        
        if signals:
            for signal in signals:
                st.info(signal)
        else:
            st.info("No significant investment signals detected")

def analyze_company(file_path, company_name, sheet_name=0, metric_column='Keterangan'):
    analyzer = IDXFinancialAnalyzer(file_path, company_name)
    analyzer.load_data(sheet_name)
    analyzer.preview_data()
    analyzer.standardize_data(metric_column)
    analyzer.calculate_financial_ratios()
    analyzer.create_visualizations()
    analyzer.generate_report()

# Streamlit App
def main():
    st.set_page_config(
        page_title="IDX Financial Analyzer",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.title("📈 IDX Financial Statement Analyzer")
    st.markdown("""
    **Menganalisis laporan keuangan perusahaan IDX**  
    Unggah file Excel laporan keuangan format IDX untuk memulai analisis
    """)
    
    with st.sidebar:
        st.header("Pengaturan Analisis")
        uploaded_file = st.file_uploader("Unggah file Excel", type=["xlsx", "xls"])
        company_name = st.text_input("Nama Perusahaan (opsional)")
        sheet_name = st.text_input("Nama Sheet (default: 0)", value="0")
        metric_column = st.text_input("Kolom Metrik (default: Keterangan)", value="Keterangan")
        analyze_btn = st.button("Mulai Analisis", type="primary")
    
    if uploaded_file and analyze_btn:
        with st.spinner("Memproses data..."):
            # Save uploaded file to temp location
            with open("temp_file.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Run analysis - PERBAIKAN DI SINI
            analyze_company(
                "temp_file.xlsx",
                company_name or uploaded_file.name.split('.')[0],
                int(sheet_name) if sheet_name.isdigit() else sheet_name,
                metric_column
            )
    
    st.sidebar.markdown("---")
    st.sidebar.info("""
    **Panduan Penggunaan:**
    1. Unggah file Excel laporan keuangan format IDX
    2. Sesuaikan pengaturan analisis jika diperlukan
    3. Klik tombol 'Mulai Analisis'
    4. Hasil akan ditampilkan di halaman utama
    """)

if __name__ == "__main__":
    main()
