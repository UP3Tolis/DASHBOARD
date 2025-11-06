# python -m streamlit run app.py

import streamlit as st
import pandas as pd

# Konfigurasi halaman Streamlit
st.set_page_config(page_title="Dashboard KPI", layout="wide")
st.title("📊 Dashboard KPI")
st.write("Menampilkan data dari SharePoint (range BQ4:BR42):")

# 🔗 Link Excel SharePoint (sudah benar formatnya)
sharepoint_url = "https://ptpln365-my.sharepoint.com/personal/ivan_falahul_ptpln365_onmicrosoft_com/Documents/002.%20Perencanaan%20UP3%20TLI/999.%20DASHBOARD/NKO%20UP3%20TLI-%20Update.xlsx?raw=1"

try:
    # Baca file Excel dari SharePoint langsung
    df = pd.read_excel(sharepoint_url, engine="openpyxl")

    # Karena pandas tidak bisa langsung ambil range, kita slice manual:
    # BQ = kolom ke-68 (0-based index)
    # BR = kolom ke-69
    # Baris 4–42 berarti index 3–41
    df_range = df.iloc[3:42, 68:70]

    # Tampilkan tabel di Streamlit
    st.subheader("📋 Data KPI (BQ4:BR42)")
    st.dataframe(df_range, use_container_width=True)

except Exception as e:
    st.error(f"Gagal membaca file: {e}")
