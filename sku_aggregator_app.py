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
    2. Sube el archivo **Vitaplena.xlsx** (SKUs en columna D desde fila 2).  
    3. Sube el archivo **Eggmarket.xlsx** (SKUs en columna F desde fila 7).  
    4. La app contará todas las apariciones de cada SKU en ambos archivos.  
    5. Actualiza la columna **master** (encabezado en O3) en el maestro, sin alterar formato.
    """
)

col1, col2, col3 = st.columns([1,1,1])
with col1:
    master_file = st.file_uploader(
        "1️⃣ Sube tu archivo maestro (xlsx)",
        type=["xlsx"], key="master"
    )
with col2:
    vitaplena_file = st.file_uploader(
        "2️⃣ Sube Vitaplena.xlsx", type=["xlsx","xls"], key="vita"
    )
with col3:
    egg_file = st.file_uploader(
        "3️⃣ Sube Eggmarket.xlsx", type=["xlsx","xls"], key="egg"
    )

if not master_file:
    st.info("Por favor, sube primero el archivo maestro para conservar su formato.")

if master_file and vitaplena_file and egg_file:
    # Cargar workbook con formato
    wb = load_workbook(filename=master_file, data_only=False)
    ws = wb.active  # primera hoja

    # Leer SKUs de Vitaplena (columna D desde fila 2)
    df_vita = pd.read_excel(vitaplena_file, usecols=[3], skiprows=1, names=["SKU"])
    # Leer SKUs de Eggmarket (columna F desde fila 7)
    df_egg = pd.read_excel(egg_file, usecols=[5], skiprows=6, names=["SKU"])

    # Limpiar y contar SKUs (mantener parte después de ':' si existe)
    df_vita["SKU"] = df_vita["SKU"].astype(str).apply(lambda x: x.split(':',1)[1] if ':' in x else x)
    df_egg["SKU"] = df_egg["SKU"].astype(str).apply(lambda x: x.split(':',1)[1] if ':' in x else x)

    counts = pd.concat([df_vita, df_egg], ignore_index=True)
    summary = counts["SKU"].value_counts().rename_axis("SKU").reset_index(name="Total")

    # Identificar columna 'master' en fila 3
    master_col = None
    for col in range(1, ws.max_column+1):
        if str(ws.cell(row=3, column=col).value).strip().lower() == 'master':
            master_col = col
            break
    if master_col is None:
        st.error("No se encontró la columna 'master' en la fila 3 del maestro.")
        st.stop()

    # Actualizar totales en maestro: SKUs en columna A desde fila 4
    total_map = dict(zip(summary['SKU'], summary['Total']))
    for r in range(4, ws.max_row+1):
        sku_cell = ws.cell(row=r, column=1).value
        if sku_cell is None:
            continue
        sku_key = str(sku_cell)
        sku_key = sku_key.split(':',1)[-1] if ':' in sku_key else sku_key
        ws.cell(row=r, column=master_col, value=total_map.get(sku_key, 0))

    # Vista previa de los primeros 10 SKUs con Totales
    preview = []
    for r in range(4, min(ws.max_row, 13)+1):
        preview.append((
            ws.cell(row=r, column=1).value,
            ws.cell(row=r, column=master_col).value
        ))
    st.subheader("✅ Totales actualizados en Maestro")
    st.table(pd.DataFrame(preview, columns=["SKU","Totales"]))

    # Descargar maestro con totales, conservando formato
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    st.download_button(
        "📥 Descargar Maestro con Totales",
        data=output,
        file_name="maestro_con_totales.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
elif master_file:
    st.info("Sube Vitaplena.xlsx y Eggmarket.xlsx para procesar SKUs.")

else:
    if not master_file:
        pass
    else:
        st.info("Esperando los archivos de ventas (Vitaplena y Eggmarket).")
