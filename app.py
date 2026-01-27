import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# Configuración de librerías para diseño de Excel
try:
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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
        <h1>🛠️ GESTIÓN DE REPUESTOS NACIONAL (INCLUYE GO)</h1>
        <p>Reporte Consolidado al <b>{hoy.strftime("%d/%m/%Y")}</b></p>
    </div>
    """, unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube cualquier archivo Excel o CSV", type=["csv", "xlsx", "xls"])

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

        # 2. SELECCIÓN INTELIGENTE DE COLUMNAS
        st.sidebar.header("⚙️ Mapeo de Columnas")
        
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
        # Convertir fecha y calcular antigüedad
        df['Fecha_DT'] = pd.to_datetime(df[c_fecha], dayfirst=True, errors='coerce')
        df['Dias_Antiguedad'] = (hoy - df['Fecha_DT']).dt.days
        df['Alerta'] = df['Dias_Antiguedad'].apply(lambda x: "🚩 CRÍTICO (+15d)" if x > 15 else "OK")

        # Filtros: Solicita o Proceso/Repuestos (con contenido)
        mask_solicita = df[c_estado].str.contains('Solicita', case=False, na=False)
        mask_proceso = df[c_estado].str.contains('Proceso/Repuestos', case=False, na=False)
        mask_tiene_rep = df[c_rep].astype(str).str.strip().apply(lambda x: len(x) > 2 and x.lower() != 'nan')
        
        # ELIMINAMOS LA EXCLUSIÓN DE "GO" para que aparezcan en la lista
        df_filtrado = df[(mask_solicita | (mask_proceso & mask_tiene_rep))].copy()
        df_filtrado = df_filtrado.sort_values(by=[c_tech, 'Dias_Antiguedad'], ascending=[True, False])

        if not df_filtrado.empty:
            # MÉTRICAS
            criticas = len(df_filtrado[df_filtrado['Dias_Antiguedad'] > 15])
            m1, m2, m3 = st.columns(3)
            m1.metric("📦 Órdenes Totales", len(df_filtrado))
            m2.metric("🚩 > 15 Días", criticas, delta=f"{criticas} urgentes", delta_color="inverse")
            m3.metric("🧑‍🔧 Servicios Activos", df_filtrado[c_tech].nunique())

            # PREPARACIÓN EXCEL
            df_filtrado['Enviado'] = "[  ]"
            cols_finales = ['Enviado', 'Alerta', c_orden, c_fecha, c_tech, c_estado, c_rep, 'Dias_Antiguedad']
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtrado[cols_finales].to_excel(writer, index=False, sheet_name='Reporte')
                ws = writer.sheets['Reporte']
                
                if EXCEL_STYLING:
                    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
                    sep_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
                    
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = Font(color='FFFFFF', bold=True)
                        cell.alignment = Alignment(horizontal='center')

                    # Separadores por técnico
                    row = 2
                    while row <= ws.max_row:
                        idx_t = cols_finales.index(c_tech) + 1
                        curr = ws.cell(row=row, column=idx_t).value
                        prev = ws.cell(row=row-1, column=idx_t).value if row > 2 else None
                        
                        if prev and curr != prev and prev != c_tech:
                            ws.insert_rows(row)
                            ws.cell(row=row, column=1).value = f"📍 SERVICIO: {curr}"
                            ws.cell(row=row, column=1).fill = sep_fill
                            ws.cell(row=row, column=1).font = Font(bold=True)
                            row += 1
                        row += 1

            st.download_button(
                label="📥 Descargar Reporte (Incluye GO y Técnicos Nacionales)",
                data=output.getvalue(),
                file_name=f"Repuestos_General_{hoy.strftime('%d-%m')}.xlsx",
                use_container_width=True
            )

            # VISTA DETALLADA
            for taller in sorted(df_filtrado[c_tech].unique()):
                sub = df_filtrado[df_filtrado[c_tech] == taller]
                retrasos = len(sub[sub['Dias_Antiguedad'] > 15])
                with st.expander(f"📍 {taller} ({len(sub)} órdenes) {'⚠️ '+str(retrasos)+' CRÍTICAS' if retrasos > 0 else ''}"):
                    st.dataframe(sub[cols_finales], hide_index=True, use_container_width=True)
        else:
            st.warning("No se encontraron órdenes pendientes con los filtros actuales.")

    except Exception as e:
        st.error(f"Error: {e}")
