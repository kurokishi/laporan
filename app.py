import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px
from io import BytesIO
import os

# Tambahan untuk PDF
import camelot

st.set_page_config(page_title="IDX Financial Analyzer (Excel + PDF)", layout="wide")
st.title("IDX Financial Analyzer (Excel + PDF)")

# ---------------- Utility Functions ----------------
def clean_number(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float)):
        return x
    s = str(x).strip()
    if s in ["-", "—", ""]:
        return np.nan
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

COMMON_ROW_KEYWORDS = {
    "total_revenue": ["pendapatan", "total pendapatan", "revenue", "penjualan"],
    "net_income": ["laba bersih", "laba tahun berjalan", "net income", "profit"],
    "total_assets": ["total aset", "total assets"],
    "total_equity": ["ekuitas", "total equity"],
    "total_liabilities": ["total liabilitas", "total liabilities", "liabilitas"],
    "current_assets": ["aset lancar", "current assets"],
    "current_liabilities": ["liabilitas lancar", "current liabilities"],
    "cash": ["kas", "kas dan setara kas", "cash and cash equivalents", "cash"]
}

def find_account_column(df, keywords):
    found = {}
    if df.empty:
        return found
    cols = [str(c).lower() for c in df.columns]
    for canonical, kws in keywords.items():
        for kw in kws:
            for i, c in enumerate(cols):
                if kw in c:
                    found[canonical] = df.columns[i]
                    break
            if canonical in found:
                break
    return found

def compute_ratios(income_df, balance_df, cash_df):
    periods = sorted(set(income_df.index) | set(balance_df.index) | set(cash_df.index))
    rows = []
    for p in periods:
        income_map = find_account_column(income_df, {
            "total_revenue": COMMON_ROW_KEYWORDS["total_revenue"],
            "net_income": COMMON_ROW_KEYWORDS["net_income"]
        })
        balance_map = find_account_column(balance_df, {
            "total_assets": COMMON_ROW_KEYWORDS["total_assets"],
            "total_equity": COMMON_ROW_KEYWORDS["total_equity"],
            "total_liabilities": COMMON_ROW_KEYWORDS["total_liabilities"],
            "current_assets": COMMON_ROW_KEYWORDS["current_assets"],
            "current_liabilities": COMMON_ROW_KEYWORDS["current_liabilities"],
            "cash": COMMON_ROW_KEYWORDS["cash"]
        })

        revenue = income_df.at[p, income_map["total_revenue"]] if "total_revenue" in income_map else np.nan
        net_income = income_df.at[p, income_map["net_income"]] if "net_income" in income_map else np.nan
        total_assets = balance_df.at[p, balance_map["total_assets"]] if "total_assets" in balance_map else np.nan
        total_equity = balance_df.at[p, balance_map["total_equity"]] if "total_equity" in balance_map else np.nan
        total_liab = balance_df.at[p, balance_map["total_liabilities"]] if "total_liabilities" in balance_map else np.nan
        current_assets = balance_df.at[p, balance_map["current_assets"]] if "current_assets" in balance_map else np.nan
        current_liab = balance_df.at[p, balance_map["current_liabilities"]] if "current_liabilities" in balance_map else np.nan
        cash = balance_df.at[p, balance_map["cash"]] if "cash" in balance_map else np.nan

        roe = net_income / total_equity if total_equity and not pd.isna(total_equity) else np.nan
        roa = net_income / total_assets if total_assets and not pd.isna(total_assets) else np.nan
        der = total_liab / total_equity if total_equity and not pd.isna(total_equity) else np.nan
        current_ratio = current_assets / current_liab if current_liab and not pd.isna(current_liab) else np.nan
        net_margin = net_income / revenue if revenue and not pd.isna(revenue) else np.nan

        rows.append({
            "period": p,
            "revenue": revenue,
            "net_income": net_income,
            "total_assets": total_assets,
            "total_equity": total_equity,
            "total_liabilities": total_liab,
            "cash": cash,
            "ROE": roe,
            "ROA": roa,
            "DER": der,
            "Current Ratio": current_ratio,
            "Net Margin": net_margin
        })
    return pd.DataFrame(rows).set_index("period")

# ---------------- PDF Parser ----------------
def parse_pdf_tables(file):
    temp_path = "temp.pdf"
    with open(temp_path, "wb") as f:
        f.write(file.read())
    tables = camelot.read_pdf(temp_path, pages="all", flavor="stream")
    os.remove(temp_path)
    dfs = [t.df for t in tables]
    return dfs

# ---------------- Streamlit Flow ----------------
uploaded = st.file_uploader("Upload file laporan keuangan (Excel IDX atau PDF)", type=["xls", "xlsx", "pdf"])
if not uploaded:
    st.stop()

if uploaded.name.lower().endswith(".pdf"):
    st.info("Mendeteksi file PDF — mencoba ekstrak tabel...")
    pdf_tables = parse_pdf_tables(uploaded)
    st.write(f"Berhasil ekstrak {len(pdf_tables)} tabel dari PDF")

    # Untuk demo, kita tampilkan tabel mentah dulu
    for i, df in enumerate(pdf_tables[:3]):
        st.write(f"Tabel {i+1}")
        st.dataframe(df.head())

    st.warning("Tahap selanjutnya: mapping tabel PDF ke Laba Rugi, Neraca, dan Arus Kas sesuai format.")

else:
    st.info("Mendeteksi file Excel — jalankan parser Excel IDX seperti biasa")
    sheets = pd.read_excel(uploaded, sheet_name=None, header=None)
    st.write(f"Terbaca {len(sheets)} sheet:", list(sheets.keys()))
    st.warning("Tahap selanjutnya: proses ke period table seperti versi Excel sebelumnya.")
