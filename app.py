import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

class IDXFinancialAnalyzer:
    def __init__(self, file_path, company_name=None):
        """
        Analyzer untuk laporan keuangan dari IDX
        
        Parameters:
        file_path (str): Path ke file Excel (.xls/.xlsx)
        company_name (str): Nama perusahaan (opsional)
        """
        self.file_path = Path(file_path)
        self.company_name = company_name or self.file_path.stem
        self.data = None
        self.processed_data = None
        
    def load_data(self, sheet_name=0):
        """Load data dari file Excel IDX"""
        try:
            # Coba baca dengan berbagai format
            if self.file_path.suffix.lower() == '.xls':
                self.data = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='xlrd')
            else:
                self.data = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='openpyxl')
            
            print(f"✅ Data berhasil dimuat: {self.data.shape[0]} baris, {self.data.shape[1]} kolom")
            return True
            
        except Exception as e:
            print(f"❌ Error loading file: {e}")
            return False
    
    def preview_data(self, rows=10):
        """Preview struktur data"""
        if self.data is None:
            print("❌ Data belum dimuat. Jalankan load_data() terlebih dahulu.")
            return
        
        print("\n📊 PREVIEW DATA:")
        print("="*50)
        print(f"Shape: {self.data.shape}")
        print(f"Columns: {list(self.data.columns)}")
        print("\nFirst few rows:")
        print(self.data.head(rows))
        
        # Deteksi kolom tahun
        year_cols = [col for col in self.data.columns if str(col).isdigit() and len(str(col)) == 4]
        print(f"\n📅 Detected year columns: {year_cols}")
    
    def standardize_data(self, metric_column='Keterangan', auto_detect=True):
        """
        Standardisasi format data IDX
        
        Parameters:
        metric_column (str): Nama kolom yang berisi deskripsi metric
        auto_detect (bool): Otomatis deteksi kolom tahun
        """
        if self.data is None:
            print("❌ Data belum dimuat.")
            return False
        
        df = self.data.copy()
        
        # Auto-detect metric column jika tidak ditemukan
        if metric_column not in df.columns:
            possible_cols = [col for col in df.columns if any(keyword in str(col).lower() 
                           for keyword in ['keterangan', 'description', 'item', 'metric'])]
            if possible_cols:
                metric_column = possible_cols[0]
                print(f"📝 Using metric column: {metric_column}")
            else:
                metric_column = df.columns[0]
                print(f"⚠️ Using first column as metric: {metric_column}")
        
        # Deteksi kolom tahun
        if auto_detect:
            year_cols = [col for col in df.columns if str(col).isdigit() and len(str(col)) == 4]
            year_cols.sort()
        else:
            year_cols = [col for col in df.columns if col != metric_column]
        
        if len(year_cols) < 2:
            print(f"⚠️ Hanya ditemukan {len(year_cols)} kolom tahun. Minimal 2 tahun diperlukan.")
            return False
        
        # Bersihkan dan konversi data numerik
        for col in year_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Filter baris yang memiliki data valid
        df = df.dropna(subset=year_cols, how='all')
        
        # Mapping standard metrics
        metric_mapping = {
            # Assets
            'total aset': 'Total Assets',
            'total aktiva': 'Total Assets', 
            'kas dan setara kas': 'Cash & Equivalents',
            'kas dan bank': 'Cash & Equivalents',
            'piutang usaha': 'Accounts Receivable',
            'persediaan': 'Inventory',
            'aset lancar': 'Current Assets',
            
            # Liabilities  
            'total liabilitas': 'Total Liabilities',
            'total kewajiban': 'Total Liabilities',
            'utang usaha': 'Accounts Payable',
            'liabilitas jangka pendek': 'Current Liabilities',
            'kewajiban lancar': 'Current Liabilities',
            'utang bank jangka panjang': 'Long-term Debt',
            'liabilitas jangka panjang': 'Long-term Debt',
            
            # Equity
            'total ekuitas': 'Total Equity',
            'modal disetor': 'Paid-in Capital',
            
            # Income Statement
            'pendapatan': 'Revenue',
            'penjualan': 'Revenue', 
            'beban pokok penjualan': 'Cost of Goods Sold',
            'laba kotor': 'Gross Profit',
            'laba usaha': 'Operating Profit',
            'laba sebelum pajak': 'Profit Before Tax',
            'laba bersih': 'Net Profit',
            'laba tahun berjalan': 'Net Profit',
            
            # Cash Flow
            'dividen': 'Dividends Paid',
            'dividen dibayar': 'Dividends Paid'
        }
        
        # Standardisasi nama metric
        df['Standard_Metric'] = df[metric_column].str.lower().str.strip()
        for key, value in metric_mapping.items():
            df.loc[df['Standard_Metric'].str.contains(key, na=False), 'Standard_Metric'] = value
        
        # Siapkan data final
        self.processed_data = {
            'metrics': df['Standard_Metric'].tolist(),
            'original_metrics': df[metric_column].tolist(),
            'years': year_cols,
            'data': df[year_cols].values
        }
        
        # Buat DataFrame terstruktur
        structured_data = []
        for i, metric in enumerate(self.processed_data['metrics']):
            row = {'Metric': metric, 'Original': self.processed_data['original_metrics'][i]}
            for j, year in enumerate(year_cols):
                row[str(year)] = self.processed_data['data'][i][j]
            structured_data.append(row)
        
        self.df_analysis = pd.DataFrame(structured_data)
        
        print(f"✅ Data terstandarisasi: {len(self.df_analysis)} metrics, {len(year_cols)} years")
        return True
    
    def calculate_financial_ratios(self):
        """Hitung rasio keuangan utama"""
        if self.df_analysis is None:
            print("❌ Data belum diproses.")
            return None
        
        df = self.df_analysis.copy()
        years = [col for col in df.columns if str(col).isdigit()]
        
        ratios = {}
        
        for year in years:
            year_ratios = {}
            
            # Helper function untuk ambil nilai
            def get_value(metric_name):
                mask = df['Metric'].str.contains(metric_name, case=False, na=False)
                values = df.loc[mask, year]
                return values.iloc[0] if len(values) > 0 and pd.notna(values.iloc[0]) else 0
            
            # Liquidity Ratios
            current_assets = get_value('Current Assets')
            cash = get_value('Cash & Equivalents') 
            current_liab = get_value('Current Liabilities')
            
            if current_liab != 0:
                year_ratios['Current Ratio'] = current_assets / current_liab
                year_ratios['Quick Ratio'] = cash / current_liab
            
            # Profitability Ratios
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
            
            # Leverage Ratios
            total_liab = get_value('Total Liabilities')
            long_term_debt = get_value('Long-term Debt')
            
            if total_assets != 0:
                year_ratios['Debt to Asset %'] = (total_liab / total_assets) * 100
            
            if total_equity != 0:
                year_ratios['Debt to Equity %'] = (total_liab / total_equity) * 100
            
            ratios[year] = year_ratios
        
        self.financial_ratios = pd.DataFrame(ratios).T
        return self.financial_ratios
    
    def create_visualizations(self, save_charts=True):
        """Buat visualisasi analisis keuangan"""
        if self.df_analysis is None:
            print("❌ Data belum diproses.")
            return
        
        # Setup plot style
        plt.style.use('seaborn-v0_8')
        sns.set_palette("husl")
        
        fig = plt.figure(figsize=(20, 15))
        
        # Data preparation
        df = self.df_analysis.copy()
        years = [col for col in df.columns if str(col).isdigit()]
        
        # 1. Key Financial Metrics Overview
        ax1 = plt.subplot(3, 3, 1)
        key_metrics = ['Total Assets', 'Revenue', 'Net Profit']
        key_data = df[df['Metric'].isin(key_metrics)]
        
        if not key_data.empty:
            key_data_pivot = key_data.set_index('Metric')[years]
            key_data_pivot.T.plot(kind='bar', ax=ax1, width=0.8)
            ax1.set_title(f'{self.company_name} - Key Financial Metrics', fontsize=12, fontweight='bold')
            ax1.set_ylabel('Value (in thousands)')
            ax1.tick_params(axis='x', rotation=45)
            ax1.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        
        # 2. Profitability Trend
        ax2 = plt.subplot(3, 3, 2)
        profit_metrics = ['Gross Profit', 'Operating Profit', 'Net Profit']
        profit_data = df[df['Metric'].isin(profit_metrics)]
        
        if not profit_data.empty:
            profit_data_pivot = profit_data.set_index('Metric')[years]
            profit_data_pivot.T.plot(kind='line', marker='o', ax=ax2, linewidth=2)
            ax2.set_title('Profitability Trend', fontsize=12, fontweight='bold')
            ax2.set_ylabel('Profit (in thousands)')
            ax2.grid(True, alpha=0.3)
            ax2.legend()
        
        # 3. Asset Composition
        ax3 = plt.subplot(3, 3, 3)
        latest_year = years[-1]
        asset_metrics = ['Current Assets', 'Total Assets']
        asset_data = df[df['Metric'].isin(asset_metrics)]
        
        if not asset_data.empty and len(asset_data) > 0:
            asset_values = asset_data[latest_year].values
            asset_labels = asset_data['Metric'].values
            colors = plt.cm.Set3(np.linspace(0, 1, len(asset_values)))
            ax3.pie(asset_values, labels=asset_labels, autopct='%1.1f%%', colors=colors)
            ax3.set_title(f'Asset Composition {latest_year}', fontsize=12, fontweight='bold')
        
        # 4. Liquidity Analysis
        ax4 = plt.subplot(3, 3, 4)
        liquidity_metrics = ['Cash & Equivalents', 'Current Liabilities']
        liquidity_data = df[df['Metric'].isin(liquidity_metrics)]
        
        if not liquidity_data.empty:
            liquidity_data_pivot = liquidity_data.set_index('Metric')[years]
            liquidity_data_pivot.T.plot(kind='bar', ax=ax4, width=0.6)
            ax4.set_title('Liquidity Analysis', fontsize=12, fontweight='bold')
            ax4.set_ylabel('Amount (in thousands)')
            ax4.tick_params(axis='x', rotation=45)
            ax4.legend()
        
        # 5. Debt Structure
        ax5 = plt.subplot(3, 3, 5)
        debt_metrics = ['Current Liabilities', 'Long-term Debt']
        debt_data = df[df['Metric'].isin(debt_metrics)]
        
        if not debt_data.empty:
            debt_data_pivot = debt_data.set_index('Metric')[years]
            debt_data_pivot.T.plot(kind='bar', stacked=True, ax=ax5, width=0.6)
            ax5.set_title('Debt Structure', fontsize=12, fontweight='bold')
            ax5.set_ylabel('Debt (in thousands)')
            ax5.tick_params(axis='x', rotation=45)
            ax5.legend()
        
        # 6. Financial Ratios
        if hasattr(self, 'financial_ratios') and self.financial_ratios is not None:
            ax6 = plt.subplot(3, 3, 6)
            ratio_cols = ['Current Ratio', 'Quick Ratio']
            available_ratios = [col for col in ratio_cols if col in self.financial_ratios.columns]
            
            if available_ratios:
                self.financial_ratios[available_ratios].plot(kind='bar', ax=ax6, width=0.6)
                ax6.set_title('Liquidity Ratios', fontsize=12, fontweight='bold')
                ax6.set_ylabel('Ratio')
                ax6.tick_params(axis='x', rotation=45)
                ax6.axhline(y=1, color='red', linestyle='--', alpha=0.7, label='Benchmark')
                ax6.legend()
        
        # 7. Profitability Ratios
        if hasattr(self, 'financial_ratios') and self.financial_ratios is not None:
            ax7 = plt.subplot(3, 3, 7)
            margin_cols = ['Gross Margin %', 'Net Margin %']
            available_margins = [col for col in margin_cols if col in self.financial_ratios.columns]
            
            if available_margins:
                self.financial_ratios[available_margins].plot(kind='line', marker='o', ax=ax7, linewidth=2)
                ax7.set_title('Profitability Margins', fontsize=12, fontweight='bold')
                ax7.set_ylabel('Percentage (%)')
                ax7.grid(True, alpha=0.3)
                ax7.legend()
        
        # 8. Return Ratios
        if hasattr(self, 'financial_ratios') and self.financial_ratios is not None:
            ax8 = plt.subplot(3, 3, 8)
            return_cols = ['ROA %', 'ROE %']
            available_returns = [col for col in return_cols if col in self.financial_ratios.columns]
            
            if available_returns:
                self.financial_ratios[available_returns].plot(kind='bar', ax=ax8, width=0.6)
                ax8.set_title('Return Ratios', fontsize=12, fontweight='bold')
                ax8.set_ylabel('Percentage (%)')
                ax8.tick_params(axis='x', rotation=45)
                ax8.legend()
        
        # 9. Year-over-Year Growth
        ax9 = plt.subplot(3, 3, 9)
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
                colors = ['green' if x > 0 else 'red' for x in growth_df['Growth %']]
                bars = ax9.bar(growth_df['Metric'], growth_df['Growth %'], color=colors, alpha=0.7)
                ax9.set_title(f'YoY Growth ({years[0]}-{years[-1]})', fontsize=12, fontweight='bold')
                ax9.set_ylabel('Growth Rate (%)')
                ax9.tick_params(axis='x', rotation=45)
                ax9.axhline(y=0, color='black', linestyle='-', alpha=0.3)
                
                # Add value labels on bars
                for bar, value in zip(bars, growth_df['Growth %']):
                    height = bar.get_height()
                    ax9.text(bar.get_x() + bar.get_width()/2., height + (0.01 * max(growth_df['Growth %'])),
                            f'{value:.1f}%', ha='center', va='bottom', fontsize=9)
        
        plt.tight_layout(pad=3.0)
        
        if save_charts:
            filename = f'{self.company_name}_financial_analysis.png'
            plt.savefig(filename, dpi=300, bbox_inches='tight')
            print(f"📊 Charts saved as: {filename}")
        
        plt.show()
    
    def generate_report(self):
        """Generate comprehensive financial analysis report"""
        if self.df_analysis is None:
            print("❌ Data belum diproses.")
            return
        
        print("\n" + "="*80)
        print(f"📈 FINANCIAL ANALYSIS REPORT - {self.company_name.upper()}")
        print("="*80)
        
        years = [col for col in self.df_analysis.columns if str(col).isdigit()]
        
        # Key Metrics Summary
        print(f"\n📊 KEY FINANCIAL METRICS:")
        print("-" * 50)
        display_metrics = ['Total Assets', 'Revenue', 'Gross Profit', 'Net Profit', 
                         'Cash & Equivalents', 'Current Liabilities', 'Total Equity']
        
        summary_data = []
        for metric in display_metrics:
            row_data = {'Metric': metric}
            metric_row = self.df_analysis[self.df_analysis['Metric'] == metric]
            
            if not metric_row.empty:
                for year in years:
                    value = metric_row[year].iloc[0]
                    row_data[year] = f"{value:,.0f}" if pd.notna(value) else "N/A"
                
                # Calculate change
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
            summary_df = pd.DataFrame(summary_data)
            print(summary_df.to_string(index=False))
        
        # Financial Ratios
        if hasattr(self, 'financial_ratios') and self.financial_ratios is not None:
            print(f"\n📈 FINANCIAL RATIOS:")
            print("-" * 50)
            print(self.financial_ratios.round(2).to_string())
        
        # Risk Analysis
        print(f"\n⚠️ RISK INDICATORS:")
        print("-" * 50)
        
        # Debt analysis
        debt_data = self.df_analysis[self.df_analysis['Metric'].str.contains('Debt|Liabilities', case=False, na=False)]
        if not debt_data.empty:
            latest_year = years[-1]
            total_debt = debt_data[latest_year].sum()
            print(f"Total Debt ({latest_year}): {total_debt:,.0f}")
        
        # Liquidity analysis
        cash_row = self.df_analysis[self.df_analysis['Metric'] == 'Cash & Equivalents']
        current_liab_row = self.df_analysis[self.df_analysis['Metric'] == 'Current Liabilities']
        
        if not cash_row.empty and not current_liab_row.empty:
            for year in years:
                cash = cash_row[year].iloc[0] if not cash_row.empty else 0
                liab = current_liab_row[year].iloc[0] if not current_liab_row.empty else 0
                if liab != 0:
                    coverage = cash / liab
                    status = "Good" if coverage > 1 else "Watch" if coverage > 0.5 else "Risk"
                    print(f"Cash Coverage {year}: {coverage:.2f}x ({status})")
        
        # Growth Analysis
        print(f"\n📈 GROWTH ANALYSIS:")
        print("-" * 50)
        
        if len(years) >= 2:
            growth_metrics = ['Revenue', 'Net Profit', 'Total Assets', 'Total Equity']
            for metric in growth_metrics:
                metric_row = self.df_analysis[self.df_analysis['Metric'] == metric]
                if not metric_row.empty:
                    old_val = metric_row[years[0]].iloc[0]
                    new_val = metric_row[years[-1]].iloc[0]
                    if pd.notna(old_val) and pd.notna(new_val) and old_val != 0:
                        growth = ((new_val - old_val) / old_val) * 100
                        trend = "📈" if growth > 0 else "📉"
                        print(f"{metric}: {growth:+.1f}% {trend}")
        
        # Investment Signals
        print(f"\n🎯 INVESTMENT SIGNALS:")
        print("-" * 50)
        
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
        
        for signal in signals:
            print(signal)
        
        print("\n" + "="*80)
        print("📋 Analysis completed. Review charts for detailed insights.")
        print("="*80)

    def run_full_analysis(self, file_path, sheet_name=0, metric_column='Keterangan'):
        """Jalankan analisis lengkap dari file Excel IDX"""
        print(f"🚀 Starting analysis for: {self.company_name}")
        print("-" * 50)
        
        # Load data
        if not self.load_data(sheet_name):
            return False
        
        # Preview data
        self.preview_data()
        
        # Standardize data
        if not self.standardize_data(metric_column):
            return False
        
        # Calculate ratios
        self.calculate_financial_ratios()
        
        # Generate visualizations
        self.create_visualizations()
        
        # Generate report
        self.generate_report()
        
        return True

# ===== USAGE EXAMPLE =====

def analyze_company(file_path, company_name=None, sheet_name=0, metric_column='Keterangan'):
    """
    Fungsi helper untuk analisis cepat
    
    Parameters:
    file_path (str): Path ke file Excel
    company_name (str): Nama perusahaan
    sheet_name (int/str): Sheet yang akan dianalisis  
    metric_column (str): Nama kolom metric
    """
    analyzer = IDXFinancialAnalyzer(file_path, company_name)
    return analyzer.run_full_analysis(file_path, sheet_name, metric_column)

# ===== CONTOH PENGGUNAAN =====

"""
# Contoh 1: Analisis langsung
file_path = "ADRO_financial.xlsx"
analyze_company(file_path, "PT Adaro Energy", sheet_name=0)

# Contoh 2: Analisis manual step-by-step
analyzer = IDXFinancialAnalyzer("BBCA_financial.xlsx", "Bank BCA")
analyzer.load_data(sheet_name="Laporan Keuangan")
analyzer.preview_data()
analyzer.standardize_data(metric_column="Keterangan")
analyzer.calculate_financial_ratios()
analyzer.create_visualizations()
analyzer.generate_report()

# Contoh 3: Multiple companies analysis
companies = [
    ("BBRI_financial.xlsx", "Bank BRI"),
    ("TLKM_financial.xlsx", "Telkom Indonesia"), 
    ("UNVR_financial.xlsx", "Unilever Indonesia")
]

for file_path, company_name in companies:
    print(f"\n{'='*60}")
    print(f"Analyzing: {company_name}")
    print(f"{'='*60}")
    analyze_company(file_path, company_name)
"""
