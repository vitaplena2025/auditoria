import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(
    page_title="SKU Aggregator - Maestro con Formato",
    page_icon="📦",
    layout="wide",
)

st.title("📦 SKU Aggregator - Maestro con Formato")
st.markdown(
    """
    1. Sube tu **archivo maestro** (Excel) con layout y colores propios.  
    2. Sube tus archivos de ventas (Vitaplena/Eggmarket).  
    3. La app rellenará la columna **master** (encabezado en O3), 
       buscando SKUs en columna A desde la fila 4.
    """
)

# Subida del maestro
master_file = st.file_uploader(
    "1️⃣ Sube tu archivo maestro (xlsx)",
    type=["xlsx"], key="master"
)
# Subida de archivos de ventas
sales_files = st.file_uploader(
    "2️⃣ Sube archivos de ventas (xlsx/xls)",
    type=["xlsx", "xls"], accept_multiple_files=True, key="sales"
)

if not master_file:
    st.info("Por favor, sube primero el archivo maestro para conservar su formato.")

if master_file:
    # Cargar workbook para preservar formato
    wb = load_workbook(filename=master_file, data_only=False)
    ws = wb.active

    # Identificar índice de columna 'master' en la fila 3
    master_col = None
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=3, column=col).value
        if isinstance(header, str) and header.strip().lower() == 'master':
            master_col = col
            break
    if master_col is None:
        st.error("No se encontró la columna 'master' en la fila 3.")
        st.stop()

    # Mostrar encabezado de master
    st.write(f"Columna 'master' encontrada en la posición: {master_col} (columna {chr(64+master_col)})")

    if not sales_files:
        st.info("Ahora sube al menos un archivo de ventas para procesar los totales.")

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
            tmp["SKU"] = tmp["SKU"].astype(str).apply(
                lambda x: x.split(':', 1)[1] if ':' in x else x
            )
            tmp["Quantity"] = pd.to_numeric(tmp["Quantity"], errors="coerce").fillna(0).astype(int)
            dfs.append(tmp)

        # Unir y agrupar totales de ventas
        all_data = pd.concat(dfs, ignore_index=True)
        summary = (
            all_data.groupby("SKU", as_index=False)["Quantity"].sum()
                   .rename(columns={"Quantity": "Total"})
        )
        summary["Total"] = summary["Total"].astype(int)
        totals_map = {str(r.SKU): r.Total for r in summary.itertuples()}

        # Actualizar columna 'master' en master_file
        for r in range(4, ws.max_row + 1):
            sku_cell = ws.cell(row=r, column=1).value
            if sku_cell is None:
                continue
            sku_key = str(sku_cell).split(':',1)[-1] if ':' in str(sku_cell) else str(sku_cell)
            total_val = totals_map.get(sku_key, 0)
            ws.cell(row=r, column=master_col, value=total_val)

        # Vista previa de filas 4-13: SKU y Totales en master_col
        preview = []
        for r in range(4, min(ws.max_row, 13) + 1):
            sku = ws.cell(row=r, column=1).value
            total = ws.cell(row=r, column=master_col).value
            preview.append((sku, total))
        st.subheader("✅ Totales Actualizados en Maestro (fila 4-13)")
        st.table(pd.DataFrame(preview, columns=["SKU", "Totales"]))

        # Botón de descarga maestro formateado con totales
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        st.download_button(
            "📥 Descargar Maestro con Totales",
            data=output,
            file_name="maestro_con_totales.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
