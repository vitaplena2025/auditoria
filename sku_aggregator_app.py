import streamlit as st
import pandas as pd
from io import BytesIO

# URL del CSV maestro publicado en Google Sheets
def get_master_url():
    return (
        "https://docs.google.com/spreadsheets/d/"
        "e/2PACX-1vRTq2EZ4kh1-7FD6Q3V0__IJsKzFiqXoBmWxsyeSFFthQcoOiKgnKovFbfhvPqNIA/"
        "pub?output=csv"
    )

st.set_page_config(
    page_title="SKU Aggregator con Master Completo",
    page_icon="📦",
    layout="centered",
)

st.title("📦 SKU Aggregator con Master Completo")
st.markdown(
    """
    1. Se carga un **maestro completo** desde Google Sheets, sin alterar sus columnas.  
    2. Sube tus Excel de **Vitaplena** o **Eggmarket**.  
    3. Se agrupan los SKUs y cantidades y se vuelcan en la columna **Totales**.
    """
)

# 1) Leer y mostrar el maestro completo
try:
    master_df = pd.read_csv(get_master_url())
except Exception as e:
    st.error(f"No pude leer el maestro: {e}")
    st.stop()

st.subheader("📋 Maestro Completo de SKUs")
st.dataframe(master_df, use_container_width=True)

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
            st.warning(f"No se reconoce {file.name}, usando col 4 y 6 por defecto.")
            sku_col = df.columns[3]
            qty_col = df.columns[5]

        temp = df[[sku_col, qty_col]].copy()
        temp.columns = ["SKU", "Quantity"]
        # Recortar parte tras ':' si existe
        temp["SKU"] = temp["SKU"].astype(str).apply(
            lambda x: x.split(":", 1)[1] if ":" in x else x
        )
        # Forzar numérico a Quantity
        temp["Quantity"] = pd.to_numeric(temp["Quantity"], errors="coerce").fillna(0)
        dfs.append(temp)

    # 3) Concatenar y agrupar totales
    all_data = pd.concat(dfs, ignore_index=True)
    summary = (
        all_data
        .groupby("SKU", as_index=False)["Quantity"]
        .sum()
        .rename(columns={"Quantity": "Totales"})
    )
    summary["Totales"] = summary["Totales"].astype(int)

    # 4) Actualizar columna Totales en master_df sin alterar otras columnas
    # Asegurar que exista la columna 'Totales'
    if 'Totales' not in master_df.columns:
        master_df['Totales'] = 0
    # Crear diccionario de mapeo y asignar
    totals_map = dict(zip(summary['SKU'], summary['Totales']))
    master_df['Totales'] = master_df[master_df.columns[0]].map(totals_map).fillna(0).astype(int)

    st.subheader("✅ Maestro con Totales Actualizados")
    st.dataframe(master_df, use_container_width=True)

    # 5) Botón de descarga del maestro modificado
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="xlsxwriter") as writer:
        master_df.to_excel(writer, index=False, sheet_name="MaestroConTotales")
    towrite.seek(0)

    st.download_button(
        label="📥 Descargar maestro con totales",
        data=towrite,
        file_name="sku_master_with_totals.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
