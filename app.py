# python -m streamlit run app.py

import streamlit as st
import pandas as pd
import altair as alt

# Fungsi warna
def color_rule(x):
    if x <= 95:
        return "red"
    elif x >= 100:
        return "green"
    else:
        return "orange"

# URL Excel
excel_url = "https://raw.githubusercontent.com/UP3Tolis/DASHBOARD/main/NKO%20UP3%20TLI.xlsx"
sheet_name = "STRG"

# Baca Excel
df = pd.read_excel(excel_url, sheet_name=sheet_name, header=None, engine="openpyxl")

st.set_page_config(page_title="Dashboard KPI", layout="wide")

# Sidebar menu
menu_items = ["UP3", "ULP", "INDICATOR", "---"]
if "selected_menu" not in st.session_state:
    st.session_state.selected_menu = menu_items[0]

st.session_state.selected_menu = st.sidebar.radio(
    label="",
    options=menu_items,
    index=menu_items.index(st.session_state.selected_menu),
    label_visibility="collapsed"
)

# ===== Konten per menu =====
if st.session_state.selected_menu == "UP3":
    # Ambil range AI4 : AV44
    indikator = df.iloc[3:44, 34].reset_index(drop=True)  # kolom AI
    nilai = df.iloc[3:44, 36:48].reset_index(drop=True)  # kolom AK..AV

    bulan = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST",
            "SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"]
    nilai.columns = bulan

    # Gabungkan DataFrame
    df_final = pd.concat([indikator.rename("Indikator"), nilai], axis=1)

    # Tambahkan kategori KPI / PI
    df_final["Kategori"] = ["KPI" if i < 7 else "PI" for i in range(len(df_final))]

    # Pilih bulan di sidebar kanan chart
    left, right = st.columns([2, 3])
    with right:
        selected_month = st.selectbox("", bulan, index=0, label_visibility="collapsed")

        # Tambahkan kolom Nilai bulan terpilih
        df_final["Nilai_Bulan"] = df_final[selected_month]

        # Tambahkan kolom Warna
        df_final["Warna"] = df_final["Nilai_Bulan"].apply(color_rule)

        # Hitung summary
        summary = df_final.groupby(["Kategori", "Warna"]).size().unstack(fill_value=0)

        # Ambil jumlah KPI & PI per warna
        jumlah_kpi_hijau = summary.loc["KPI"].get("green", 0)
        jumlah_kpi_orange = summary.loc["KPI"].get("orange", 0)
        jumlah_kpi_merah  = summary.loc["KPI"].get("red", 0)

        jumlah_pi_hijau   = summary.loc["PI"].get("green", 0)
        jumlah_pi_orange  = summary.loc["PI"].get("orange", 0)
        jumlah_pi_merah   = summary.loc["PI"].get("red", 0)

        # Tampilkan jumlah KPI/PI per warna
        col1, col2, col3 = st.columns(3)
        with col1:
            st.write("🟩 GREEN")
            st.subheader(f"KPI    : {jumlah_kpi_hijau}")
            st.subheader(f"PI     : {jumlah_pi_hijau}")
        with col2:
            st.write("🟧 YELLOW")
            st.subheader(f"KPI : {jumlah_kpi_orange}")
            st.subheader(f"PI  : {jumlah_pi_orange}")
        with col3:
            st.write("🟥 RED")
            st.subheader(f"KPI : {jumlah_kpi_merah}")
            st.subheader(f"PI  : {jumlah_pi_merah}")
            
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.header("")

        with col2:
            st.header("GAP KINERJA")
            # Filter indikator orange & merah
            df_filter = df_final[df_final["Warna"].isin(["orange", "red"])]

            # Pilih bulan terpilih
            df_filter["Nilai_Bulan"] = df_filter[selected_month]

            # Chart Altair
            chart_filter = alt.Chart(df_filter).encode(
                y=alt.Y("Indikator:N", sort=None, title=None, axis=alt.Axis(grid=False)),
                x=alt.X("Nilai_Bulan:Q", title=None, axis=alt.Axis(grid=False, ticks=False, labels=False))
            )

            bars_filter = chart_filter.mark_bar(size=8).encode(
                color=alt.Color(
                    "Warna:N",
                    scale=alt.Scale(
                        domain=["red", "orange"],
                        range=["red", "orange"]
                    ),
                    legend=None
                )
            )

            text_filter = chart_filter.mark_text(
                align="left",
                baseline="middle",
                dx=3
            ).encode(
                text=alt.Text("Nilai_Bulan:Q", format=".2f")
            )

            st.altair_chart(bars_filter + text_filter, use_container_width=False)

    with left:
        # Chart
        tabel = pd.DataFrame({
            "Indikator": df_final["Indikator"],
            "Nilai": df_final[selected_month]
        })
        tabel["warna"] = tabel["Nilai"].apply(color_rule)

        base = alt.Chart(tabel).encode(
            y=alt.Y(
                "Indikator:N",
                sort=None,
                title=None,
                axis=alt.Axis(grid=False, labelOverlap=False, labelLimit=150)
            ),
            x=alt.X(
                "Nilai:Q",
                title=None,
                axis=alt.Axis(grid=False, ticks=False, labels=False)
            )
        )

        bars = base.mark_bar(size=8).encode(
            color=alt.Color(
                "warna:N",
                scale=alt.Scale(
                    domain=["red", "orange", "green"],
                    range=["red", "orange", "green"]
                ),
                legend=None
            )
        )

        text = base.mark_text(
            align="left",
            baseline="middle",
            dx=3
        ).encode(
            text=alt.Text("Nilai:Q", format=".2f")
        )

        chart = (bars + text).properties(height=750, width=600)
        st.altair_chart(chart, use_container_width=False)

elif st.session_state.selected_menu == "ULP":
    st.header("📄 ULP")
    st.write("Ini konten untuk menu ULP")

elif st.session_state.selected_menu == "INDICATOR":
    st.header("📄 INDICATOR")
    st.write("Ini konten untuk menu INDICATOR")

elif st.session_state.selected_menu == "---":
    st.header("📄 ---")
    st.write("---")
