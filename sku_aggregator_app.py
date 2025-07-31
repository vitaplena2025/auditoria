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
    3. La app actualizará la **columna O** ("Totales") de tu maestro, conservando el layout y colores.
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

    # Leer master a DataFrame para mostrar (opcional)
    master_df = pd.DataFrame(ws.values)
    st.subheader("📋 Vista previa del Maestro (solo datos)")
    st.dataframe(master_df.head(10), use_container_width=True)

    # 2) Subida de archivos de ventas
    uploaded = st.file_uploader(
        "Sube tus archivos Excel (Vitaplena/Eggmarket)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="sales"
    )

    if uploaded:
        # Procesar ventas
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
            temp["SKU"] = temp["SKU"].astype(str).apply(
                lambda x: x.split(':',1)[1] if ':' in x else x
            )
            temp["Quantity"] = pd.to_numeric(temp["Quantity"], errors="coerce").fillna(0)
            dfs.append(temp)

        all_data = pd.concat(dfs, ignore_index=True)
        summary = (
            all_data.groupby("SKU", as_index=False)["Quantity"].sum()
                  .rename(columns={"Quantity": "Total"})
        )
        summary["Total"] = summary["Total"].astype(int)

        # 3) Actualizar columna O en el workbook
        # Asumimos encabezados en fila 1, SKUs en columna A, Totales en columna O
        sku_map = {str(r['SKU']): r['Total'] for _, r in summary.iterrows()}
        for row in range(2, ws.max_row + 1):
            sku_cell = ws.cell(row=row, column=1).value
            if sku_cell is None:
                continue
            sku_str = str(sku_cell).split(':',1)[-1] if ':' in str(sku_cell) else str(sku_cell)
            total_val = sku_map.get(sku_str, 0)
            ws.cell(row=row, column=15, value=total_val)  # col 15 = O

        # Mostrar resultado
        st.subheader("✅ Totales actualizados en Maestro (se conserva formato)")
        # Mostrar primeras 10 filas con SKU y Totales
        updated = []
        for row in range(1, min(ws.max_row, 11) + 1):
            sku = ws.cell(row=row, column=1).value
            total = ws.cell(row=row, column=15).value
            updated.append((sku, total))
        st.table(pd.DataFrame(updated, columns=[master_df.iloc[0,0], 'Totales']))

        # 4) Descargar workbook actualizado
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        st.download_button(
            "📥 Descargar Maestro con Totales", 
            data=output, 
            file_name="maestro_con_totales.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
