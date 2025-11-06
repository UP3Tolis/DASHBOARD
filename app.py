# python -m streamlit run app.py

import streamlit as st
import pandas as pd

# Konfigurasi halaman
st.set_page_config(page_title="Dashboard KPI", layout="wide")
st.title("📊 Dashboard KPI")
st.write("Menampilkan data dari sheet STRG cell A1:F1")

# RAW GitHub URL
github_excel_url = "https://github.com/UP3Tolis/DASHBOARD/raw/refs/heads/main/NKO%20UP3%20TLI.xlsx"

try:
    # Baca langsung sheet STRG
    df = pd.read_excel(github_excel_url, engine="openpyxl", sheet_name="STRG")

    # Ambil range A1:F1 → baris indeks 0, kolom 0–5
    df_range = df.iloc[0:1, 0:6]

    st.subheader("📋 Data dari STRG (A1:F1)")
    st.dataframe(df_range, use_container_width=True)

except Exception as e:
    st.error(f"Gagal membaca file dari GitHub: {e}")