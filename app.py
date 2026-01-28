import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# Configuración de diseño para Excel
try:
    from openpyxl.styles import Font, PatternFill
    EXCEL_STYLING = True
except ImportError:
    EXCEL_STYLING = False

st.set_page_config(page_title="Gestión Postventa - Gerardo Ortiz", layout="wide")

# --- FUNCIONES DE APOYO ---
def generar_excel(dataframe, col_tech):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Reporte')
        ws = writer.sheets['Reporte']
        if EXCEL_STYLING:
            header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = Font(color='FFFFFF', bold=True)
    return output.getvalue()

def generar_txt_mensaje_exacto(taller, total_pedidos, n_criticos):
    # Formato exacto según la imagen proporcionada por el usuario
    mensaje = (
        f"Estimados {taller}:\n\n"
        f"Reciban un cordial saludo, el presente mensaje es para consultarles el estado de las siguientes ordenes de servicio:\n\n"
        f"pendientes ({total_pedidos})\n\n"
        f"Criticas ( {n_criticos})\n\n"
        f"A la favorable atencion de la presente y comentarios sobre las ordenes agradezco su atencion.\n\n"
        f"Atentamente,\n"
        f"Departamento Postventa Gerardo Ortiz"
    )
    return mensaje

# --- BANNER ---
hoy = datetime.now()
st.markdown(f"""
    <div style="background: linear-gradient(90deg, #1F4E78 0%, #2E75B6 100%); padding: 20px; border-radius: 15px; color: white; text-align: center; margin-bottom: 20px;">
        <h1>🛠️ PANEL DE COMUNICACIÓN POSTVENTA</h1>
        <p>Generación de Mensajes TXT con Formato Oficial | {hoy.strftime("%d/%m/%Y")}</p>
    </div>
    """, unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube el archivo de órdenes", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    try:
        # LECTURA DE ARCHIVO
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file, sep=',', engine='python', encoding='utf-8')
            except: df = pd.read_csv(uploaded_file, sep=';', engine='python', encoding='latin-1')
        else:
            df = pd.read_excel(uploaded_file)

        df.columns = df.columns.str.strip()
        c_fecha, c_tech, c_estado, c_rep, c_orden, c_prod, c_serie = 'Fecha', 'Técnico', 'Estado', 'Repuestos', '#Orden', 'Producto', 'Serie/Artículo'

        # PROCESAMIENTO DE DATOS
        df['Fecha_DT'] = pd.to_datetime(df[c_fecha], dayfirst=True, errors='coerce')
        df['Dias_Antiguedad'] = (hoy - df['Fecha_DT']).dt.days
        df['Es_Critico'] = df['Dias_Antiguedad'] > 15
        df['Alerta'] = df['Dias_Antiguedad'].apply(lambda x: "🚩 CRÍTICO (+15d)" if x > 15 else "OK")

        # FILTROS DE EXCLUSIÓN (Anti-Facturados y Anulados)
        df['Estado_UPPER'] = df[c_estado].astype(str).str.upper().str.strip()
        palabras_prohibidas = ['ANULADA', 'ANULADO', 'FACTURADO', 'TERMINADO', 'CERRADA', 'ENTREGADO', 'REPARADO', 'RECLAMO PROVEEDOR']
        df_limpio = df[~df['Estado_UPPER'].str.contains('|'.join(palabras_prohibidas), na=False)].copy()

        # FILTRO DE TÉCNICOS (GO y Nacionales con estados específicos)
        df_limpio['es_go'] = df_limpio[c_tech].str.upper().str.startswith('GO', na=False)
        df_filtrado = df_limpio[ (df_limpio['es_go'] == True) | (df_limpio['Estado_UPPER'].str.contains('SOLICITA|PROCESO/REPUESTOS', na=False)) ].copy()

        # ORDENAMIENTO
        df_filtrado['Prioridad'] = df_filtrado['es_go'].map({True: 0, False: 1})
        df_filtrado = df_filtrado.sort_values(by=['Prioridad', c_tech, 'Dias_Antiguedad'], ascending=[True, True, False])

        if not df_filtrado.empty:
            df_filtrado['Enviado'] = "[  ]"
            cols_finales = ['Enviado', 'Alerta', c_orden, c_fecha, c_tech, c_estado, c_prod, c_serie, c_rep, 'Dias_Antiguedad']

            # SECCIONES POR TALLER
            for taller in df_filtrado[c_tech].unique():
                sub = df_filtrado[df_filtrado[c_tech] == taller]
                n_criticos = sub['Es_Critico'].sum()
                
                label = f"{'🏢' if taller.upper().startswith('GO') else '🔧'} {taller} ({len(sub)} órdenes)"
                if n_criticos > 0: label += f" | 🚩 {n_criticos} CRÍTICOS"
                
                with st.expander(label):
                    c1, c2 = st.columns(2)
                    
                    with c1:
                        # Botón Excel
                        data_excel = generar_excel(sub[cols_finales], c_tech)
                        st.download_button(f"📥 Descargar Excel", data_excel, f"Reporte_{taller}.xlsx", key=f"ex_{taller}", use_container_width=True)
                    
                    with c2:
                        # Botón TXT con el formato de la imagen
                        mensaje_txt = generar_txt_mensaje_exacto(taller, len(sub), n_criticos)
                        st.download_button(f"📄 Generar Mensaje (TXT)", mensaje_txt, f"Mensaje_{taller}.txt", key=f"txt_{taller}", use_container_width=True)
                    
                    st.dataframe(sub[cols_finales], hide_index=True, use_container_width=True)
        else:
            st.warning("No hay órdenes activas con los filtros aplicados.")
    except Exception as e:
        st.error(f"Error: {e}")
