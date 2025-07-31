import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(
    page_title="SKU Aggregator - Maestro y Distribución de Paletas",
    page_icon="📦",
    layout="wide",
)

st.title("📦 SKU Aggregator con Distribución de Paletas")
st.markdown(
    """
    1. Sube tu **archivo maestro** (Excel) con layout y colores propios.  
    2. Sube **Vitaplena.xlsx** (SKUs en columna D desde fila 2).  
    3. Sube **Eggmarket.xlsx** (SKUs en columna F desde fila 7).  
    4. La app sumará todas las apariciones de cada SKU.  
    5. Actualiza la columna **master** (encabezado en O3) y distribuye los totales
       en paletas (columnas B a M) con máximo 30 bultos por paleta.
    """
)

# Uploaders
col1, col2, col3 = st.columns([1,1,1])
with col1:
    master_file = st.file_uploader("1️⃣ Sube tu archivo maestro (xlsx)", type=["xlsx"], key="master")
with col2:
    vitaplena_file = st.file_uploader("2️⃣ Sube Vitaplena.xlsx", type=["xlsx","xls"], key="vita")
with col3:
    egg_file = st.file_uploader("3️⃣ Sube Eggmarket.xlsx", type=["xlsx","xls"], key="egg")

if not master_file:
    st.info("Por favor, sube primero el archivo maestro para conservar su formato.")
elif not (vitaplena_file and egg_file):
    st.info("Ahora sube Vitaplena.xlsx y Eggmarket.xlsx para procesar SKUs.")
else:
    # Cargar workbook para preservar formato
    wb = load_workbook(filename=master_file, data_only=False)
    ws = wb.active  # Primera hoja

    # 1) Leer y limpiar SKUs de Vitaplena (columna D desde fila 2)
    df_vita = pd.read_excel(vitaplena_file, usecols=[3], skiprows=1, names=["SKU"])
    # 2) Leer y limpiar SKUs de Eggmarket (columna F desde fila 7)
    df_egg = pd.read_excel(egg_file, usecols=[5], skiprows=6, names=["SKU"])
    # Extraer parte tras ':' si existe
    for df in (df_vita, df_egg):
        df["SKU"] = df["SKU"].astype(str).apply(lambda x: x.split(':',1)[-1] if ':' in x else x)

    # 3) Concatenar y contar
    counts = pd.concat([df_vita, df_egg], ignore_index=True)
    summary = counts["SKU"].value_counts().rename_axis("SKU").reset_index(name="Total")

    # 4) Identificar columna 'master' en fila 3
    master_col = None
    for col in range(1, ws.max_column+1):
        if str(ws.cell(row=3, column=col).value).strip().lower() == 'master':
            master_col = col
            break
    if master_col is None:
        st.error("No se encontró la columna 'master' en la fila 3 del maestro.")
        st.stop()

    # 5) Actualizar totales y distribuir en paletas B:M
    # Map de totales
    total_map = dict(zip(summary['SKU'], summary['Total']))
    # Columnas de paletas: B=2 hasta M=13
    pallet_cols = list(range(2, 14))
    MAX_PER_PALLET = 30

    for r in range(4, ws.max_row + 1):
        sku_cell = ws.cell(row=r, column=1).value
        if sku_cell is None:
            continue
        sku_key = str(sku_cell).split(':',1)[-1] if ':' in str(sku_cell) else str(sku_cell)
        total_val = int(total_map.get(sku_key, 0))
        # Escribir total en columna master
        ws.cell(row=r, column=master_col, value=total_val)
        # Distribuir en paletas
        remaining = total_val
        for col in pallet_cols:
            if remaining <= 0:
                ws.cell(row=r, column=col, value=None)
            else:
                assign = min(MAX_PER_PALLET, remaining)
                ws.cell(row=r, column=col, value=assign)
                remaining -= assign

    # 6) Vista previa primeras 10 filas
    preview = []
    headers = ['SKU', 'Total'] + [f'Pal{i}' for i in range(1, len(pallet_cols)+1)]
    for r in range(4, min(ws.max_row, 13) + 1):
        row_vals = [
            ws.cell(row=r, column=1).value,
            ws.cell(row=r, column=master_col).value
        ] + [ws.cell(row=r, column=col).value for col in pallet_cols]
        preview.append(row_vals)
    st.subheader("✅ Preview Distribución de Paletas (filas 4-13)")
    st.dataframe(pd.DataFrame(preview, columns=headers), use_container_width=True)

    # 7) Botón de descarga maestro con formato intacto
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    st.download_button(
        "📥 Descargar Maestro con Totales y Paletas", 
        data=output,
        file_name="maestro_con_totales_paletas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
