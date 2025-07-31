import streamlit as st
import pandas as pd
from io import BytesIO

# URL del CSV maestro publicado
MASTER_CSV_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "e/2PACX-1vRTq2EZ4kh1-7FD6Q3V0__IJsKzFiqXoBmWxsyeSFFthQcoOiKgnKovFbfhvPqNIA/"
    "pub?output=csv"
)

st.set_page_config(
    page_title="SKU Aggregator con Master",
    page_icon="📦",
    layout="centered",
)

st.title("📦 SKU Aggregator con Master")
st.markdown(
    """
    1. Se carga un **maestro de SKUs** desde Google Sheets.  
    2. Sube tus Excel de **Vitaplena** o **Eggmarket**.  
    3. Se agrupan los SKUs y cantidades, y se vuelcan sobre el maestro.  
    """
)

# 1) Leer y mostrar el maestro
try:
    master = pd.read_csv(MASTER_CSV_URL)
except Exception as e:
    st.error(f"No pude leer el maestro: {e}")
    st.stop()

# Asumimos que el SKU maestro está en la primera columna
master_sku_col = master.columns[0]
master = master[[master_sku_col]].drop_duplicates().copy()
master.columns = ["SKU"]
master["Total"] = 0  # columna a rellenar

st.subheader("📋 Maestro de SKUs")
st.dataframe(master, use_container_width=True)


# 2) Subida de archivos
uploaded = st.file_uploader(
    "Sube tus archivos Excel (xlsx/xls)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded:
    dfs = []
    for file in uploaded:
        df = pd.read_excel(file)
        name = file.name.lower()
        # Detectar columnas según origen
        if "vitaplena" in name:
            sku_col = df.columns[3]
            qty_col = df.columns[5]
        elif "eggmarket" in name:
            sku_col = df.columns[5]
            qty_col = df.columns[6]
        else:
            st.warning(f"No se reconoce {file.name}, uso col 4 y 6 por defecto.")
            sku_col = df.columns[3]
            qty_col = df.columns[5]

        temp = df[[sku_col, qty_col]].copy()
        temp.columns = ["SKU", "Quantity"]

        # Recortar tras ':' si existe
        temp["SKU"] = temp["SKU"].astype(str).apply(
            lambda x: x.split(":", 1)[1] if ":" in x else x
        )
        # Forzar numérico
        temp["Quantity"] = pd.to_numeric(temp["Quantity"], errors="coerce").fillna(0)

        dfs.append(temp)

    # 3) Concatenar y agrupar totales
    all_data = pd.concat(dfs, ignore_index=True)
    summary = (
        all_data
        .groupby("SKU", as_index=False)["Quantity"]
        .sum()
        .rename(columns={"Quantity": "Total"})
    )
    summary["Total"] = summary["Total"].astype(int)

    # 4) Hacer left join sobre el maestro
    result = master[["SKU"]].merge(summary, on="SKU", how="left")
    result["Total"] = result["Total"].fillna(0).astype(int)

    st.subheader("✅ Maestro con Totales Actualizados")
    st.dataframe(result, use_container_width=True)

    # 5) Botón de descarga
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="xlsxwriter") as writer:
        result.to_excel(writer, index=False, sheet_name="Resumen")
    towrite.seek(0)

    st.download_button(
        label="📥 Descargar maestro con totales",
        data=towrite,
        file_name="sku_master_with_totals.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
