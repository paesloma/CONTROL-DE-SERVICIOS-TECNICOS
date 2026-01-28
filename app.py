import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# Configuración de diseño para Excel
try:
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_STYLING = True
except ImportError:
    EXCEL_STYLING = False

st.set_page_config(page_title="Gestión de Repuestos - Descarga por Secciones", layout="wide")

# Función para generar el Excel (Reutilizable para global e individual)
def generar_excel(dataframe, nombre_seccion, col_tech):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Reporte')
        ws = writer.sheets['Reporte']
        if EXCEL_STYLING:
            header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
            go_fill = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid')
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = Font(color='FFFFFF', bold=True)
            # Resaltar si es GO
            for r in range(2, ws.max_row + 1):
                tech_val = str(ws.cell(row=r, column=dataframe.columns.get_loc(col_tech)+1).value)
                if tech_val.upper().startswith('GO'):
                    for c in range(1, len(dataframe.columns)+1):
                        ws.cell(row=r, column=c).fill = go_fill
    return output.getvalue()

# --- BANNER ---
hoy = datetime.now()
st.markdown(f"""
    <div style="background: linear-gradient(90deg, #1F4E78 0%, #2E75B6 100%); padding: 20px; border-radius: 15px; color: white; text-align: center; margin-bottom: 20px;">
        <h1>🛠️ GESTIÓN DE REPUESTOS: REPORTES SEGMENTADOS</h1>
        <p>Descarga el reporte general o <b>individual por taller</b> | {hoy.strftime("%d/%m/%Y")}</p>
    </div>
    """, unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube el archivo de órdenes", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file, sep=',', engine='python', encoding='utf-8')
            except:
                df = pd.read_csv(uploaded_file, sep=';', engine='python', encoding='latin-1')
        else:
            df = pd.read_excel(uploaded_file)

        df.columns = df.columns.str.strip()
        c_fecha, c_tech, c_estado, c_rep, c_orden, c_prod, c_serie = 'Fecha', 'Técnico', 'Estado', 'Repuestos', '#Orden', 'Producto', 'Serie/Artículo'

        # Procesamiento
        df['Fecha_DT'] = pd.to_datetime(df[c_fecha], dayfirst=True, errors='coerce')
        df['Dias_Antiguedad'] = (hoy - df['Fecha_DT']).dt.days
        df['Es_Critico'] = df['Dias_Antiguedad'] > 15
        df['Alerta'] = df['Dias_Antiguedad'].apply(lambda x: "🚩 CRÍTICO (+15d)" if x > 15 else "OK")

        # Filtros de Exclusión (Anulados, Reclamos, etc.)
        df['Estado_UPPER'] = df[c_estado].astype(str).str.upper().str.strip()
        palabras_prohibidas = ['ANULADA', 'ANULADO', 'FACTURADO', 'TERMINADO', 'CERRADA', 'ENTREGADO', 'REPARADO', 'RECLAMO PROVEEDOR']
        df_limpio = df[~df['Estado_UPPER'].str.contains('|'.join(palabras_prohibidas), na=False)].copy()

        # Filtro de Inclusión
        df_limpio['es_go'] = df_limpio[c_tech].str.upper().str.startswith('GO', na=False)
        df_filtrado = df_limpio[ (df_limpio['es_go'] == True) | (df_limpio['Estado_UPPER'].str.contains('SOLICITA|PROCESO/REPUESTOS', na=False)) ].copy()

        # Ordenamiento
        df_filtrado['Prioridad'] = df_filtrado['es_go'].map({True: 0, False: 1})
        df_filtrado = df_filtrado.sort_values(by=['Prioridad', c_tech, 'Dias_Antiguedad'], ascending=[True, True, False])

        if not df_filtrado.empty:
            # Botón General
            df_filtrado['Enviado'] = "[  ]"
            cols_finales = ['Enviado', 'Alerta', c_orden, c_fecha, c_tech, c_estado, c_prod, c_serie, c_rep, 'Dias_Antiguedad']
            
            data_global = generar_excel(df_filtrado[cols_finales], "General", c_tech)
            st.download_button("📥 DESCARGAR REPORTE GENERAL (TODOS)", data_global, 
                               file_name=f"Reporte_General_{hoy.strftime('%d-%m')}.xlsx", use_container_width=True, type="primary")

            # SECCIONES POR TALLER
            st.markdown("### 🧑‍🔧 Detalle por Taller / Centro GO")
            for taller in df_filtrado[c_tech].unique():
                sub = df_filtrado[df_filtrado[c_tech] == taller]
                n_criticos = sub['Es_Critico'].sum()
                es_go = taller.upper().startswith('GO')
                
                label = f"{'🏢' if es_go else '🔧'} {taller} ({len(sub)} órdenes)"
                if n_criticos > 0: label += f" | 🚩 {n_criticos} CRÍTICOS"
                
                with st.expander(label):
                    # Botón de descarga específico para este taller
                    data_taller = generar_excel(sub[cols_finales], taller, c_tech)
                    st.download_button(
                        label=f"📥 Descargar Reporte de {taller}",
                        data=data_taller,
                        file_name=f"Reporte_{taller.replace(' ', '_')}_{hoy.strftime('%d-%m')}.xlsx",
                        key=f"btn_{taller}"
                    )
                    st.dataframe(sub[cols_finales], hide_index=True, use_container_width=True)
        else:
            st.warning("No hay órdenes activas para mostrar.")

    except Exception as e:
        st.error(f"Error: {e}")
