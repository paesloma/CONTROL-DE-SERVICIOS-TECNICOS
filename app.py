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

st.set_page_config(page_title="Gestión de Repuestos Universal", layout="wide")

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
        <h1>🛠️ GESTIÓN DE REPUESTOS UNIVERSAL</h1>
        <p>Corte de Control: <b>{hoy.strftime("%d/%m/%Y")}</b></p>
    </div>
    """, unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube tu archivo (Excel o CSV)", type=["csv", "xlsx", "xls"])

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

        # 2. SELECCIÓN DE COLUMNAS (Barra lateral)
        st.sidebar.header("⚙️ Configuración de Columnas")
        st.sidebar.info("Si el sistema no detecta las columnas, selecciónalas aquí:")
        
        def find_col(options, targets):
            for t in targets:
                for o in options:
                    if t.lower() in o.lower(): return o
            return options[0]

        col_fecha = st.sidebar.selectbox("Columna de Fecha", all_cols, index=all_cols.index(find_col(all_cols, ['Fecha', 'Fec'])))
        col_tech = st.sidebar.selectbox("Columna de Técnico", all_cols, index=all_cols.index(find_col(all_cols, ['Técnico', 'Tech', 'Responsable'])))
        col_estado = st.sidebar.selectbox("Columna de Estado", all_cols, index=all_cols.index(find_col(all_cols, ['Estado', 'Status'])))
        col_rep = st.sidebar.selectbox("Columna de Repuestos", all_cols, index=all_cols.index(find_col(all_cols, ['Repuestos', 'Piezas', 'Parte'])))
        col_orden = st.sidebar.selectbox("Columna de Orden", all_cols, index=all_cols.index(find_col(all_cols, ['Orden', 'ID', 'Número'])))

        # 3. PROCESAMIENTO
        df['Fecha_DT'] = pd.to_datetime(df[col_fecha], dayfirst=True, errors='coerce')
        df['Dias_Antiguedad'] = (hoy - df['Fecha_DT']).dt.days
        df['Aviso'] = df['Dias_Antiguedad'].apply(lambda x: "🚩 +15 DÍAS" if x > 15 else "OK")

        # Filtros (Lógica solicitada)
        cond_solicita = df[col_estado].str.contains('Solicita', case=False, na=False)
        es_proceso = df[col_estado].str.contains('Proceso/Repuestos', case=False, na=False)
        tiene_repuestos = df[col_rep].astype(str).str.strip().apply(lambda x: len(x) > 2 and x.lower() != 'nan')
        
        # Exclusión de técnicos internos
        patron_excluir = r'^GO|STDIGICENT|STBMDIGI|TCLCUE|TCLCUENC'
        mask_tecnico = ~df[col_tech].str.upper().str.contains(patron_excluir, na=False, regex=True)
        
        df_filtrado = df[(cond_solicita | (es_proceso & tiene_repuestos)) & mask_tecnico].copy()
        df_filtrado = df_filtrado.sort_values(by=[col_tech, 'Dias_Antiguedad'], ascending=[True, False])

        if not df_filtrado.empty:
            # MÉTRICAS
            m1, m2, m3 = st.columns(3)
            m1.metric("📦 Órdenes", len(df_filtrado))
            m2.metric("🚩 Críticas (>15d)", len(df_filtrado[df_filtrado['Dias_Antiguedad'] > 15]))
            m3.metric("🧑‍🔧 Técnicos", df_filtrado[col_tech].nunique())

            # PREPARACIÓN EXCEL
            df_filtrado['Check'] = "[  ]"
            columnas_finales = ['Check', 'Aviso', col_orden, col_fecha, col_tech, col_estado, col_rep, 'Dias_Antiguedad']
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtrado[columnas_finales].to_excel(writer, index=False, sheet_name='Reporte')
                ws = writer.sheets['Reporte']
                
                if EXCEL_STYLING:
                    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
                    sep_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = Font(color='FFFFFF', bold=True)

                    # Separadores por técnico
                    row = 2
                    while row <= ws.max_row:
                        idx_t = columnas_finales.index(col_tech) + 1
                        curr = ws.cell(row=row, column=idx_t).value
                        prev = ws.cell(row=row-1, column=idx_t).value if row > 2 else None
                        if prev and curr != prev and prev != col_tech:
                            ws.insert_rows(row)
                            ws.cell(row=row, column=1).value = f"📍 TALLER: {curr}"
                            ws.cell(row=row, column=1).fill = sep_fill
                            ws.cell(row=row, column=1).font = Font(bold=True)
                            row += 1
                        row += 1

            st.download_button(
                label="📥 Descargar Reporte Personalizado",
                data=output.getvalue(),
                file_name=f"Reporte_Repuestos_{hoy.strftime('%d-%m')}.xlsx",
                use_container_width=True
            )

            # VISTA WEB
            for taller in sorted(df_filtrado[col_tech].unique()):
                sub = df_filtrado[df_filtrado[col_tech] == taller]
                retrasos = len(sub[sub['Dias_Antiguedad'] > 15])
                with st.expander(f"📍 {taller} ({len(sub)} órdenes) {'⚠️ '+str(retrasos)+' retrasadas' if retrasos > 0 else ''}"):
                    st.dataframe(sub[columnas_finales], hide_index=True)
        else:
            st.warning("No hay órdenes que coincidan con los filtros de Repuestos.")

    except Exception as e:
        st.error(f"Error procesando el archivo: {e}")
