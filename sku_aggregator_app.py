import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="SKU Aggregator",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="auto",
)

st.title("📦 SKU Aggregator")
st.markdown(
    """
    Sube uno o más archivos Excel de **Vitaplena** o **Eggmarket**  
    y obtén un resumen consolidado de todos los SKUs con sus cantidades totales.
    """
)

uploaded = st.file_uploader(
    "Sube tus archivos Excel (xlsx/xls)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded:
    dfs = []
    for file in uploaded:
        # Leer cada Excel
        df = pd.read_excel(file)
        name = file.name.lower()
        # Detectar columnas según origen
        if "vitaplena" in name:
            sku_col = df.columns[3]       # columna 4
            qty_col = df.columns[5]       # columna F (índice 5)
        elif "eggmarket" in name:
            sku_col = df.columns[5]       # columna 6
            qty_col = df.columns[6]       # columna 7
        else:
            st.warning(f"No se reconoce {file.name}, usando col 4 y 6 por defecto.")
            sku_col = df.columns[3]
            qty_col = df.columns[5]
        temp = df[[sku_col, qty_col]].copy()
        temp.columns = ["SKU", "Quantity"]
        dfs.append(temp)

    # Concatenar y agrupar
    all_data = pd.concat(dfs, ignore_index=True)
    summary = (
        all_data
        .groupby("SKU", as_index=False)
        .agg({"Quantity": "sum"})
        .sort_values("Quantity", ascending=False)
    )

    st.success("✅ Resumen generado:")
    st.dataframe(summary, use_container_width=True)

    # Prepara descarga de Excel
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="xlsxwriter") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")
        writer.save()
    towrite.seek(0)

    st.download_button(
        label="📥 Descargar resumen (Excel)",
        data=towrite,
        file_name="sku_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )