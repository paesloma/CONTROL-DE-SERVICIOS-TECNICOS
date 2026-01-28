import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

try:
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_STYLING = True
except ImportError:
    EXCEL_STYLING = False

st.set_page_config(page_title="Gestión de Repuestos - Filtro Estricto", layout="wide")

# --- BANNER ---
hoy = datetime.now()
st.markdown(f"""
    <div style="background: linear-gradient(90deg, #1F4E78 0%, #2E75B6 100%); padding: 20px; border-radius: 15px; color: white; text-align: center; margin-bottom: 20px;">
        <h1>🛠️ CONTROL DE REPUESTOS: FILTRO DE EXCLUSIÓN</h1>
        <p><b>Excluyendo:</b> Facturados y Terminados | <b>Prioridad:</b> Centros GO | {hoy.strftime("%d/%m/%Y")}</p>
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
        all_cols = df.columns.tolist()

        # Mapeo (Sidebar)
        st.sidebar.header("⚙️ Configuración")
        def detectar(targets):
            for t in targets:
                for col in all_cols:
                    if t.lower() in col.lower(): return all_cols.index(col)
            return 0

        c_fecha = st.sidebar.selectbox("Fecha", all_cols, index=detectar(['Fecha']))
        c_tech = st.sidebar.selectbox("Técnico", all_cols, index=detectar(['Técnico']))
        c_estado = st.sidebar.selectbox("Estado", all_cols, index=detectar(['Estado']))
        c_rep = st.sidebar.selectbox("Repuestos", all_cols, index=detectar(['Repuestos']))
        c_orden = st.sidebar.selectbox("Orden #", all_cols, index=detectar(['Orden', '#']))
        c_prod = st.sidebar.selectbox("Producto", all_cols, index=detectar(['Producto']))
        c_serie = st.sidebar.selectbox("Serie/Artículo", all_cols, index=detectar(['Serie']))

        # Procesamiento
        df['Fecha_DT'] = pd.to_datetime(df[c_fecha], dayfirst=True, errors='coerce')
        df['Dias_Antiguedad'] = (hoy - df['Fecha_DT']).dt.days
        df['Alerta'] = df['Dias_Antiguedad'].apply(lambda x: "🚩 CRÍTICO (+15d)" if x > 15 else "OK")

        # --- LÓGICA DE FILTRADO ANTI-FACTURADOS ---
        df['es_go'] = df[c_tech].str.upper().str.startswith('GO', na=False)
        
        # 1. Normalizar estados (quitar espacios y poner en minúsculas para comparar)
        df['Estado_Clean'] = df[c_estado].astype(str).str.strip()

        # 2. Definir lo que SÍ queremos para Nacionales
        mask_solicita = df['Estado_Clean'].str.contains('Solicita', case=False, na=False)
        mask_proceso = df['Estado_Clean'].str.contains('Proceso/Repuestos', case=False, na=False)
        
        # 3. Definir lo que NO queremos (Exclusión explícita)
        # Si contiene "Facturado" o "Terminado", lo marcamos como Falso para nacionales
        mask_excluir = df['Estado_Clean'].str.contains('Facturado|Terminado|Cerrada|Entregado', case=False, na=False)

        # 4. Aplicar Filtro
        # Para Nacionales: (Solicita O Proceso) Y NO (Excluidos)
        # Para GO: Pasan todos (como solicitaste anteriormente)
        df_filtrado = df[ (df['es_go'] == True) | ((mask_solicita | mask_proceso) & ~mask_excluir) ].copy()

        # Ordenar: GO primero
        df_filtrado['Prioridad'] = df_filtrado['es_go'].map({True: 0, False: 1})
        df_filtrado = df_filtrado.sort_values(by=['Prioridad', c_tech, 'Dias_Antiguedad'], ascending=[True, True, False])

        if not df_filtrado.empty:
            # MÉTRICAS
            m1, m2 = st.columns(2)
            m1.metric("📦 Órdenes en Gestión", len(df_filtrado))
            m2.metric("🚩 Retrasadas (>15d)", len(df_filtrado[df_filtrado['Dias_Antiguedad'] > 15]))

            # PREPARACIÓN EXCEL
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

            st.download_button("📥 Descargar Reporte Limpio", output.getvalue(), 
                               file_name=f"Reporte_Repuestos_Final_{hoy.strftime('%d-%m')}.xlsx", use_container_width=True)

            # VISTA WEB
            for taller in df_filtrado[c_tech].unique():
                sub = df_filtrado[df_filtrado[c_tech] == taller]
                es_go = taller.upper().startswith('GO')
                with st.expander(f"{'🏢' if es_go else '🔧'} {taller} - {len(sub)} órdenes"):
                    st.dataframe(sub[cols_finales], hide_index=True, use_container_width=True)
        else:
            st.warning("No se encontraron órdenes que cumplan los filtros de gestión activa.")

    except Exception as e:
        st.error(f"Error: {e}")
