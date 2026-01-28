import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# Configuración de librerías para diseño de Excel
try:
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_STYLING = True
except ImportError:
    EXCEL_STYLING = False

st.set_page_config(page_title="Gestión de Repuestos Nacional", layout="wide")

# --- BANNER SUPERIOR ---
hoy = datetime.now()
st.markdown(f"""
    <style>
    .main-banner {{
        background: linear-gradient(90deg, #1F4E78 0%, #2E75B6 100%);
        padding: 20px;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin-bottom: 20px;
    }}
    </style>
    <div class="main-banner">
        <h1>🛠️ GESTIÓN DE REPUESTOS: PRIORIDAD GO</h1>
        <p>Reporte Consolidado al <b>{hoy.strftime("%d/%m/%Y")}</b></p>
    </div>
    """, unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube el archivo de órdenes (Excel o CSV)", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 1. LECTURA FLEXIBLE
        if uploaded_file.name.endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file, sep=',', engine='python', encoding='utf-8')
            except:
                df = pd.read_csv(uploaded_file, sep=';', engine='python', encoding='latin-1')
        else:
            df = pd.read_excel(uploaded_file)

        df.columns = df.columns.str.strip()
        all_cols = df.columns.tolist()

        # 2. MAPEO DE COLUMNAS
        st.sidebar.header("⚙️ Configuración")
        def detectar(targets):
            for t in targets:
                for col in all_cols:
                    if t.lower() in col.lower(): return all_cols.index(col)
            return 0

        c_fecha = st.sidebar.selectbox("Fecha", all_cols, index=detectar(['Fecha', 'Fec']))
        c_tech = st.sidebar.selectbox("Técnico", all_cols, index=detectar(['Técnico', 'Tech', 'Responsable']))
        c_estado = st.sidebar.selectbox("Estado", all_cols, index=detectar(['Estado', 'Status']))
        c_rep = st.sidebar.selectbox("Repuestos", all_cols, index=detectar(['Repuestos', 'Piezas']))
        c_orden = st.sidebar.selectbox("Orden #", all_cols, index=detectar(['Orden', 'ID', '#']))

        # 3. PROCESAMIENTO
        df['Fecha_DT'] = pd.to_datetime(df[c_fecha], dayfirst=True, errors='coerce')
        df['Dias_Antiguedad'] = (hoy - df['Fecha_DT']).dt.days
        df['Alerta'] = df['Dias_Antiguedad'].apply(lambda x: "🚩 CRÍTICO (+15d)" if x > 15 else "OK")

        # Filtros de Estado
        mask_solicita = df[c_estado].str.contains('Solicita', case=False, na=False)
        mask_proceso = df[c_estado].str.contains('Proceso/Repuestos', case=False, na=False)
        mask_tiene_rep = df[c_rep].astype(str).str.strip().apply(lambda x: len(x) > 2 and x.lower() != 'nan')
        
        df_filtrado = df[(mask_solicita | (mask_proceso & mask_tiene_rep))].copy()

        # --- LÓGICA DE PRIORIDAD GO ---
        # 0 para técnicos que empiezan con GO (prioridad alta), 1 para los demás.
        df_filtrado['Prioridad'] = df_filtrado[c_tech].str.upper().str.startswith('GO').map({True: 0, False: 1})
        
        # Ordenamos: 1ero por Prioridad, 2do por Nombre de Técnico, 3ero por Antigüedad
        df_filtrado = df_filtrado.sort_values(by=['Prioridad', c_tech, 'Dias_Antiguedad'], ascending=[True, True, False])

        if not df_filtrado.empty:
            # MÉTRICAS
            m1, m2, m3 = st.columns(3)
            m1.metric("📦 Total Órdenes", len(df_filtrado))
            m2.metric("🚩 Críticas (>15d)", len(df_filtrado[df_filtrado['Dias_Antiguedad'] > 15]))
            m3.metric("🧑‍🔧 Técnicos", df_filtrado[c_tech].nunique())

            # PREPARACIÓN EXCEL
            df_filtrado['Enviado'] = "[  ]"
            cols_excel = ['Enviado', 'Alerta', c_orden, c_fecha, c_tech, c_estado, c_rep, 'Dias_Antiguedad']
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtrado[cols_excel].to_excel(writer, index=False, sheet_name='Reporte')
                ws = writer.sheets['Reporte']
                
                if EXCEL_STYLING:
                    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
                    go_fill = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid') # Azul claro para GO
                    sep_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
                    
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = Font(color='FFFFFF', bold=True)

                    # Estilo y separadores
                    row = 2
                    while row <= ws.max_row:
                        idx_t = cols_excel.index(c_tech) + 1
                        curr_val = str(ws.cell(row=row, column=idx_t).value)
                        prev_val = str(ws.cell(row=row-1, column=idx_t).value) if row > 2 else None
                        
                        # Resaltar filas de técnicos GO en el Excel
                        if curr_val.upper().startswith('GO'):
                            for col in range(1, len(cols_excel) + 1):
                                ws.cell(row=row, column=col).fill = go_fill

                        if prev_val and curr_val != prev_val and prev_val != c_tech:
                            ws.insert_rows(row)
                            ws.cell(row=row, column=1).value = f"📍 SERVICIO: {curr_val}"
                            ws.cell(row=row, column=1).fill = sep_fill
                            ws.cell(row=row, column=1).font = Font(bold=True)
                            row += 1
                        row += 1

            st.download_button(
                label="📥 Descargar Reporte (GO Primero)",
                data=output.getvalue(),
                file_name=f"Reporte_Prioridad_GO_{hoy.strftime('%d-%m')}.xlsx",
                use_container_width=True
            )

            # VISTA WEB AGRUPADA
            # Obtenemos la lista de técnicos respetando el nuevo orden de prioridad
            orden_tecnicos = df_filtrado[c_tech].unique()
            
            for taller in orden_tecnicos:
                sub = df_filtrado[df_filtrado[c_tech] == taller]
                es_go = taller.upper().startswith('GO')
                icono = "🏢" if es_go else "🔧"
                
                with st.expander(f"{icono} {taller} ({len(sub)} órdenes)"):
                    st.dataframe(sub[cols_excel], hide_index=True, use_container_width=True)
        else:
            st.warning("No hay órdenes pendientes para mostrar.")

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
