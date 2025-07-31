import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(
    page_title="SKU Aggregator con Master Formateado",
    page_icon="📦",
    layout="centered",
)

st.title("📦 SKU Aggregator con Master Formateado")
st.markdown(
    """
    1. Sube tu **archivo maestro (Excel)** con formato personalizado.  
    2. Sube tus Excel de **Vitaplena** o **Eggmarket**.  
    3. La app actualizará la columna **Totales** a partir de la celda O4, 
       coincidiendo con los SKUs listados en la columna B desde la fila 4.
    """
)

# 1) Subida del maestro Excel
master_file = st.file_uploader(
    "Sube tu archivo maestro con formato (xlsx)",
    type=["xlsx"],
    key="master"
)

if master_file:
    # Cargar libro con openpyxl para preservar formato
    wb = load_workbook(filename=master_file, data_only=False)
    ws = wb.active  # usa la primera hoja

    # Mostrar SKUs del maestro desde B4
    skus_preview = []
    for row in ws.iter_rows(min_row=4, max_row=min(ws.max_row, 10), min_col=2, max_col=2, values_only=True):
        skus_preview.append(row[0])
    st.subheader("📋 SKUs del Maestro (desde B4)")
    st.table({"SKU": skus_preview})

    # 2) Subida de archivos de ventas
    uploaded = st.file_uploader(
        "Sube tus archivos Excel (Vitaplena/Eggmarket)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="sales"
    )

    if uploaded:
        # Procesar archivos de ventas
        dfs = []
        for file in uploaded:
            df = pd.read_excel(file)
            name = file.name.lower()
            if "vitaplena" in name:
                sku_col, qty_col = df.columns[3], df.columns[5]
            elif "eggmarket" in name:
                sku_col, qty_col = df.columns[5], df.columns[6]
            else:
                sku_col, qty_col = df.columns[3], df.columns[5]

            temp = df[[sku_col, qty_col]].copy()
            temp.columns = ["SKU", "Quantity"]
            # Recortar parte tras ':' si existe
            temp["SKU"] = temp["SKU"].astype(str).apply(
                lambda x: x.split(':',1)[1] if ':' in x else x
            )
            # Forzar numérico a Quantity
            temp["Quantity"] = pd.to_numeric(temp["Quantity"], errors="coerce").fillna(0)
            dfs.append(temp)

        # Agrupar totales
        all_data = pd.concat(dfs, ignore_index=True)
        summary = (
            all_data
            .groupby("SKU", as_index=False)["Quantity"]
            .sum()
            .rename(columns={"Quantity": "Total"})
        )
        summary["Total"] = summary["Total"].astype(int)

        # 3) Actualizar Totales en el workbook
        # SKUs en col B (2), Totales en col O (15), datos desde fila 4
        totals_map = {row.SKU: row.Total for row in summary.itertuples()}
        for row in range(4, ws.max_row + 1):
            sku_cell = ws.cell(row=row, column=2).value
            if sku_cell is None:
                continue
            sku_str = str(sku_cell).split(':',1)[-1] if ':' in str(sku_cell) else str(sku_cell)
            total_val = totals_map.get(sku_str, 0)
            ws.cell(row=row, column=15, value=total_val)

        # Mostrar resultados actualizados
        results_preview = []
        for row in range(4, min(ws.max_row, 10) + 1):
            sku = ws.cell(row=row, column=2).value
            total = ws.cell(row=row, column=15).value
            results_preview.append((sku, total))
        st.subheader("✅ Totales Actualizados (fila 4+)")
        st.table({"SKU": [r[0] for r in results_preview], "Totales": [r[1] for r in results_preview]})

        # 4) Descargar workbook conservando formato
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        st.download_button(
            "📥 Descargar Maestro con Totales", 
            data=output, 
            file_name="maestro_con_totales.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
