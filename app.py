import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

try:
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_STYLING = True
except ImportError:
    EXCEL_STYLING = False

st.set_page_config(page_title="Gestión de Repuestos - Filtro Avanzado", layout="wide")

# --- BANNER ---
hoy = datetime.now()
st.markdown(f"""
    <div style="background: linear-gradient(90deg, #1F4E78 0%, #2E75B6 100%); padding: 20px; border-radius: 15px; color: white; text-align: center; margin-bottom: 20px;">
        <h1>🛠️ CONTROL DE REPUESTOS: FILTRO DE EXCLUSIÓN TOTAL</h1>
        <p><b>Ocultando:</b> Anulados, Facturados, Terminados y Reclamos Proveedor</p>
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
        
        # Columnas fijas
        c_fecha, c_tech, c_estado, c_rep, c_orden, c_prod, c_serie = 'Fecha', 'Técnico', 'Estado', 'Repuestos', '#Orden', 'Producto', 'Serie/Artículo'

        # Procesamiento de fechas
        df['Fecha_DT'] = pd.to_datetime(df[c_fecha], dayfirst=True, errors='coerce')
        df['Dias_Antiguedad'] = (hoy - df['Fecha_DT']).dt.days
        df['Alerta'] = df['Dias_Antiguedad'].apply(lambda x: "🚩 CRÍTICO (+15d)" if x > 15 else "OK")

        # --- LÓGICA DE FILTRADO REFORZADA ---
        df['Estado_UPPER'] = df[c_estado].astype(str).str.upper().str.strip()

        # LISTA NEGRA (Exclusión absoluta para todos, incluidos GO)
        # Se añade 'RECLAMO PROVEEDOR'
        palabras_prohibidas = [
            'ANULADA', 'ANULADO', 'FACTURADO', 'TERMINADO', 
            'CERRADA', 'ENTREGADO', 'REPARADO', 'RECLAMO PROVEEDOR'
        ]
        regex_prohibido = '|'.join(palabras_prohibidas)
        
        # Primero borramos todo lo que esté en la lista negra
        df_base_limpia = df[~df['Estado_UPPER'].str.contains(regex_prohibido, na=False)].copy()

        # LISTA BLANCA (Para Nacionales)
        palabras_permitidas = ['SOLICITA', 'PROCESO/REPUESTOS']
        regex_permitido = '|'.join(palabras_permitidas)

        df_base_limpia['es_go'] = df_base_limpia[c_tech].str.upper().str.startswith('GO', na=False)
        
        # FILTRO FINAL:
        # GO pasan todos (si no están en lista negra)
        # Nacionales deben estar en lista blanca
        df_filtrado = df_base_limpia[ 
            (df_base_limpia['es_go'] == True) | 
            (df_base_limpia['Estado_UPPER'].str.contains(regex_permitido, na=False)) 
        ].copy()

        # Ordenar: GO Primero
        df_filtrado['Prioridad'] = df_filtrado['es_go'].map({True: 0, False: 1})
        df_filtrado = df_filtrado.sort_values(by=['Prioridad', c_tech, 'Dias_Antiguedad'], ascending=[True, True, False])

        if not df_filtrado.empty:
            st.success(f"✅ Filtro aplicado: {len(df_filtrado)} órdenes activas. Se han ocultado reclamos y anulados.")
            
            # EXCEL
            df_filtrado['Enviado'] = "[  ]"
            cols_finales = ['Enviado', 'Alerta', c_orden, c_fecha, c_tech, c_estado, c_prod, c_serie, c_rep, 'Dias_Antiguedad']
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtrado[cols_finales].to_excel(writer, index=False, sheet_name='Reporte')
                ws = writer.sheets['Reporte']
                if EXCEL_STYLING:
                    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
                    go_fill = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid')
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = Font(color='FFFFFF', bold=True)
                    for r in range(2, ws.max_row + 1):
                        tech_val = str(ws.cell(row=r, column=cols_finales.index(c_tech)+1).value)
                        if tech_val.upper().startswith('GO'):
                            for c in range(1, len(cols_finales)+1):
                                ws.cell(row=r, column=c).fill = go_fill

            st.download_button("📥 Descargar Reporte Depurado", output.getvalue(), 
                               file_name=f"Reporte_Limpio_{hoy.strftime('%d-%m')}.xlsx", use_container_width=True)

            # VISTA WEB
            for taller in df_filtrado[c_tech].unique():
                sub = df_filtrado[df_filtrado[c_tech] == taller]
                es_go = taller.upper().startswith('GO')
                with st.expander(f"{'🏢' if es_go else '🔧'} {taller} - {len(sub)} órdenes"):
                    st.dataframe(sub[cols_finales], hide_index=True, use_container_width=True)
        else:
            st.warning("⚠️ No hay órdenes activas con los filtros actuales.")

    except Exception as e:
        st.error(f"Error: {e}")
