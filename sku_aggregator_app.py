import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(
    page_title="SKU Aggregator con Master Formateado",
    page_icon="📦",
    layout="wide",
)

st.title("📦 SKU Aggregator - Maestro Formateado")
st.markdown(
    """
    1. Sube tu **archivo maestro** (Excel) con layout y colores propios.
    2. Sube tus archivos de ventas (Vitaplena/Eggmarket).
    3. La app rellenará la columna **O** ("Totales") a partir de la fila 4,
       usando SKUs en la columna **B**.
    """
)

# Layout de uploaders
col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader(
        "Sube tu archivo maestro (xlsx)",
        type=["xlsx"],
        key="master"
    )
with col2:
    sales_files = st.file_uploader(
        "Sube archivos de ventas (xlsx/xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="sales"
    )

if not master_file:
    st.info("Por favor, sube primero el archivo maestro para conservar su formato.")

if master_file:
    # Cargar workbook para preservar formato
    wb = load_workbook(filename=master_file, data_only=False)
    ws = wb.active

    # Preview maestro: mostrar primeras 10 SKUs desde B4
    skus_sample = [ws.cell(row=r, column=2).value for r in range(4, min(ws.max_row+1, 14))]
    st.subheader("📋 SKUs del Maestro (muestra B4:B13)")
    st.write(skus_sample)

    if sales_files:
        # Procesar archivos de ventas
        dfs = []
        for sf in sales_files:
            df_sales = pd.read_excel(sf)
            name = sf.name.lower()
            if "vitaplena" in name:
                sku_col, qty_col = df_sales.columns[3], df_sales.columns[5]
            elif "eggmarket" in name:
                sku_col, qty_col = df_sales.columns[5], df_sales.columns[6]
            else:
                sku_col, qty_col = df_sales.columns[3], df_sales.columns[5]

            tmp = df_sales[[sku_col, qty_col]].copy()
            tmp.columns = ["SKU", "Quantity"]
            # Extraer parte tras ':' si existe
            tmp["SKU"] = tmp["SKU"].astype(str).apply(lambda x: x.split(':', 1)[1] if ':' in x else x)
            tmp["Quantity"] = pd.to_numeric(tmp["Quantity"], errors="coerce").fillna(0)
            dfs.append(tmp)

        # Unir y agrupar totales
        all_data = pd.concat(dfs, ignore_index=True)
        summary = (
            all_data.groupby("SKU", as_index=False)["Quantity"].sum()
                   .rename(columns={"Quantity": "Total"})
        )
        summary["Total"] = summary["Total"].astype(int)

        # Mapeo de totales
        totals_map = {str(row.SKU): row.Total for row in summary.itertuples()}

        # Actualizar celdas en columna O (15) empezando en fila 4
        for r in range(4, ws.max_row + 1):
            sku_val = ws.cell(row=r, column=2).value
            if sku_val is None:
                continue
            sku_key = str(sku_val).split(':',1)[-1] if ':' in str(sku_val) else str(sku_val)
            ws.cell(row=r, column=15, value=totals_map.get(sku_key, 0))

        # Preview resultados actualizados filas 4-13
        updated = [(ws.cell(r,2).value, ws.cell(r,15).value) for r in range(4, min(ws.max_row+1,14))]
        df_preview = pd.DataFrame(updated, columns=["SKU", "Totales"] )
        st.subheader("✅ Totales Actualizados (B4:O13)")
        st.table(df_preview)

        # Botón de descarga del maestro con formato intacto
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        st.download_button(
            "📥 Descargar Maestro con Totales",
            data=output,
            file_name="maestro_con_totales.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Sube al menos un archivo de ventas para procesar los totales.")
