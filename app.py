import streamlit as st
import pandas as pd
import datetime
import sqlite3
import plotly.express as px
import os
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import streamlit.components.v1 as components

# Configuración de página
st.set_page_config(layout="wide", page_title="Control de Obra Eléctrica Avanzado", page_icon="⚡")

# --- INICIALIZAR RUTA/ESTADO DE PROYECTO ---
if "proyecto_activo" not in st.session_state:
    st.session_state.proyecto_activo = None

# --- FUNCIÓN DE ENVÍO DE CORREO ELECTRÓNICO ---
def enviar_reporte_correo(destinatarios, asunto, cuerpo, archivo_bytes, nombre_archivo):
    try:
        EMAIL_EMISOR = st.secrets["SMTP_EMAIL"]
        PASSWORD_EMISOR = st.secrets["SMTP_PASSWORD"]
    except Exception:
        st.error("❌ No se encontraron las credenciales 'SMTP_EMAIL' y 'SMTP_PASSWORD' en Secrets de Streamlit.")
        return False

    msg = MIMEMultipart()
    msg['From'] = EMAIL_EMISOR
    msg['To'] = ", ".join(destinatarios)
    msg['Subject'] = asunto

    msg.attach(MIMEText(cuerpo, 'html'))

    adjunto = MIMEApplication(archivo_bytes, _subtype="xlsx")
    adjunto.add_header('Content-Disposition', 'attachment', filename=nombre_archivo)
    msg.attach(adjunto)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_EMISOR, PASSWORD_EMISOR)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"❌ Error al enviar el correo: {e}")
        return False

# --- MENÚ LATERAL DIRECTO ---
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3227/3227840.png", width=70)
st.sidebar.markdown("👤 **Modo:** Acceso Total Abierto")

st.sidebar.title("Navegación del Sistema")
opcion = st.sidebar.radio("Ir a la pestaña:", [
    "📈 Analítica Avanzada y KPIs", 
    "📂 Visión por Proyecto y Detalle",
    "📦 Inventario y Conteo de Columnas", 
    "📝 Carga y Gestión de Campo", 
    "📥 Migración Inicial (Excel)"
])

# Estilos personalizados CSS para modernizar el Dashboard y optimizar espacio vertical
st.markdown("""
    <style>
    /* Reducir el margen superior del contenedor principal */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 1rem !important;
    }
    .stApp {
        background-color: #0e1117;
    }
    .kpi-card {
        background: linear-gradient(135deg, #1f293d 0%, #111827 100%);
        padding: 22px;
        border-radius: 12px;
        border-left: 5px solid #3b82f6;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        margin-bottom: 15px;
    }
    .kpi-title {
        color: #9ca3af;
        font-size: 14px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        color: #ffffff;
        font-size: 28px;
        font-weight: 700;
        margin-top: 5px;
    }
    .kpi-delta {
        font-size: 13px;
        margin-top: 5px;
        font-weight: 500;
    }
    </style>
""", unsafe_allow_html=True)

st.title("⚡ Panel de Control de Obra Eléctrica Avanzado")
st.markdown("---")

DB_NAME = "obra_trazabilidad.db"

def conectar_db():
    return sqlite3.connect(DB_NAME)

def inicializar_db():
    conn = conectar_db()
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(piquetes)")
    columnas = [col[1] for col in cursor.fetchall()]
    
    if len(columnas) > 0 and "cantidad_aisladores" not in columnas:
        conn.close()
        try: os.remove(DB_NAME)
        except: pass
        conn = conectar_db()
        cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS piquetes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tramo TEXT,
            piquete TEXT UNIQUE,
            tipo_estructura TEXT,
            longitud_poste TEXT,
            cantidad_aisladores INTEGER DEFAULT 3,
            excavacion TEXT,
            verticalizado TEXT,
            tendido TEXT,
            flechado TEXT,
            engrampado TEXT,
            montaje_aislador TEXT,
            montaje_riendas TEXT,
            fecha_montaje TEXT,
            tipo_de_equipo TEXT,
            anexo_montaje TEXT,
            red_line TEXT,
            observacion_ofm TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cronogramas (
            tramo TEXT PRIMARY KEY,
            inicio TEXT,
            entrega TEXT
        )
    """)
    conn.commit()
    conn.close()

inicializar_db()

HITOS_OBRA = [
    "excavacion", 
    "verticalizado", 
    "montaje_riendas", 
    "montaje_aislador", 
    "tendido", 
    "flechado", 
    "engrampado"
]

CARPA_ARCHIVOS = "archivos_obra"
if not os.path.exists(CARPA_ARCHIVOS):
    os.makedirs(CARPA_ARCHIVOS)

# -------------------------------------------------------------------------
# MÓDULO 5: MIGRACIÓN INICIAL DESDE EXCEL
# -------------------------------------------------------------------------
if opcion == "📥 Migración Inicial (Excel)":
    st.subheader("📥 Inicialización y Carga de Planilla Maestra Excel")
    st.markdown("Cargue el archivo Excel inicial para estructurar los piquetes y frentes de trabajo de forma limpia.")
    
    archivo_excel = st.file_uploader("Suba la planilla de Trazabilidad (.xlsx)", type=["xlsx"], key="uploader_excel_maestro")
    nombre_proyecto_manual = st.text_input("Ingrese el Nombre del Proyecto / Frente (Ej: WM, LTM):", value="WM", key="input_nombre_proyecto")
    
    boton_procesar = st.button("🔄 Procesar y Migrar Datos a la Base", type="primary")

    if boton_procesar:
        if archivo_excel is None:
            st.error("❌ Por favor, primero seleccione y suba un archivo Excel.")
        else:
            try:
                archivo_excel.seek(0)
                df_test = pd.read_excel(archivo_excel, nrows=15)
                skip_rows = 0
                
                for i, row in df_test.iterrows():
                    valores_fila = [str(val).strip().upper() for val in row.values if pd.notna(val)]
                    if "PIQUETE" in valores_fila or "TIPO ESTRUCTURA" in valores_fila:
                        skip_rows = i
                        break
                
                archivo_excel.seek(0)
                df = pd.read_excel(archivo_excel, skiprows=skip_rows)
                
                df.columns = df.columns.astype(str).str.strip().str.upper()
                df = df.loc[:, ~df.columns.duplicated()]
                
                if "PIQUETE" not in df.columns:
                    st.error(f"❌ No se encontró la columna 'PIQUETE'. Columnas detectadas: {list(df.columns)}")
                else:
                    df = df.dropna(subset=["PIQUETE"])
                    
                    conn = conectar_db()
                    conn.execute("DELETE FROM piquetes WHERE tramo = ?", (nombre_proyecto_manual,))
                    
                    def get_val(row, col_name):
                        val = row.get(col_name)
                        if isinstance(val, pd.Series):
                            val = val.iloc[0]
                        if pd.isna(val) or str(val).strip().lower() in ["nan", "none", "", "nat"]:
                            return None
                        return str(val).strip()

                    registros_cargados = 0
                    for _, row in df.iterrows():
                        piquete_val = get_val(row, "PIQUETE")
                        if piquete_val:
                            cant_aisl = get_val(row, "CANTIDAD AISLADORES") or get_val(row, "CANT_AISLADORES") or get_val(row, "AISLADORES")
                            try:
                                cant_aisl = int(float(cant_aisl)) if cant_aisl else 3
                            except:
                                cant_aisl = 3

                            conn.execute("""
                                INSERT OR REPLACE INTO piquetes (
                                    tramo, piquete, tipo_estructura, longitud_poste, cantidad_aisladores, excavacion, verticalizado, 
                                    tendido, flechado, engrampado, montaje_aislador, montaje_riendas, fecha_montaje, observacion_ofm
                                )
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                nombre_proyecto_manual,
                                piquete_val,
                                get_val(row, "TIPO ESTRUCTURA") or "S/D",
                                get_val(row, "LONGITUD POSTE"),
                                cant_aisl,
                                get_val(row, "EXCAV PIQUETES"),
                                get_val(row, "VERTICALIZADO"),
                                get_val(row, "TENDIDO DE CONDUCTOR"),
                                get_val(row, "FLECHADO"),
                                get_val(row, "ENGRAMPADO"),
                                get_val(row, "MONTAJE DE AISLADOR"),
                                get_val(row, "MONTAJE RIENDAS"),
                                get_val(row, "FECHA DE LIBERACION"),
                                get_val(row, "OBSERVACION")
                            ))
                            registros_cargados += 1
                                
                    conn.commit()
                    conn.close()
                    
                    if registros_cargados > 0:
                        st.session_state.proyecto_activo = nombre_proyecto_manual
                        st.success(f"✔️ ¡Migración exitosa! Se procesaron y guardaron {registros_cargados} piquetes.")
                        st.rerun()
            except Exception as e:
                st.error(f"❌ Error al procesar el Excel: {e}")

# -------------------------------------------------------------------------
# MÓDULO 2: VISIÓN POR PROYECTO Y DETALLE
# -------------------------------------------------------------------------
elif opcion == "📂 Visión por Proyecto y Detalle":
    st.subheader("📂 Visión Detallada por Proyecto / Frente")
    
    conn = conectar_db()
    try:
        proyectos_df = pd.read_sql_query("SELECT DISTINCT tramo FROM piquetes", conn)
        lista_proyectos = [t for t in proyectos_df['tramo'].dropna().tolist() if str(t).strip().lower() != 'nan' and str(t).strip() != '']
    except Exception as e:
        st.error(f"Error al conectar con la base de datos: {e}")
        conn.close()
        lista_proyectos = []

    if not lista_proyectos:
        st.warning("No hay datos cargados en la base de datos aún. Realiza una migración desde la pestaña de Excel.")
        conn.close()
    else:
        idx_defecto = 0
        if st.session_state.proyecto_activo in lista_proyectos:
            idx_defecto = lista_proyectos.index(st.session_state.proyecto_activo)

        proyecto_sel = st.selectbox("Seleccione el Proyecto / Frente a consultar:", lista_proyectos, index=idx_defecto)

        if proyecto_sel:
            df_proyecto = pd.read_sql_query("SELECT * FROM piquetes WHERE tramo = ?", conn, params=(proyecto_sel,))
            
            peso_por_hito = 100 / len(HITOS_OBRA)
            df_proyecto["Avance_%"] = 0
            for hito in HITOS_OBRA:
                df_proyecto["Avance_%"] += df_proyecto[hito].notna().astype(int) * peso_por_hito
            df_proyecto["Avance_%"] = df_proyecto["Avance_%"].round().astype(int)

            total_piquetes = len(df_proyecto)
            avance_promedio = int(df_proyecto["Avance_%"].mean()) if total_piquetes > 0 else 0
            completados = len(df_proyecto[df_proyecto["Avance_%"] == 100])

            total_aisladores = df_proyecto["cantidad_aisladores"].fillna(3).sum()
            aisladores_instalados = df_proyecto[df_proyecto["montaje_aislador"].notna()]["cantidad_aisladores"].fillna(3).sum()

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f"""<div class='kpi-card' style='border-left-color: #3b82f6;'><div class='kpi-title'>📍 Total Piquetes</div><div class='kpi-value'>{total_piquetes}</div></div>""", unsafe_allow_html=True)
            with col2:
                st.markdown(f"""<div class='kpi-card' style='border-left-color: #10b981;'><div class='kpi-title'>✅ Piquetes 100% Finalizados</div><div class='kpi-value'>{completados}</div></div>""", unsafe_allow_html=True)
            with col3:
                st.markdown(f"""<div class='kpi-card' style='border-left-color: #f59e0b;'><div class='kpi-title'>📈 % Avance Físico Promedio</div><div class='kpi-value'>{avance_promedio}%</div></div>""", unsafe_allow_html=True)
            with col4:
                st.markdown(f"""<div class='kpi-card' style='border-left-color: #8b5cf6;'><div class='kpi-title'>🔌 Aisladores Instalados</div><div class='kpi-value'>{int(aisladores_instalados)} / {int(total_aisladores)}</div></div>""", unsafe_allow_html=True)

            st.markdown("---")
            st.subheader("🔍 Filtrar y Explorar Piquetes Cargados")
            
            buscar_piquete = st.text_input("Buscar por Número / Código de Piquete:")
            if buscar_piquete:
                df_proyecto = df_proyecto[df_proyecto['piquete'].astype(str).str.contains(buscar_piquete, case=False, na=False)]

            st.dataframe(df_proyecto, use_container_width=True)

            csv = df_proyecto.to_csv(index=False).encode('utf-8')
            st.download_button(
                label=f"📥 Descargar datos del proyecto {proyecto_sel} (CSV)",
                data=csv,
                file_name=f"reporte_{proyecto_sel}.csv",
                mime="text/csv",
            )

    conn.close()

# -------------------------------------------------------------------------
# MÓDULO 3: INVENTARIO Y CONTEO POR COLUMNAS
# -------------------------------------------------------------------------
elif opcion == "📦 Inventario y Conteo de Columnas":
    st.subheader("📦 Métricas de Inventario y Control de Columnas")
    
    conn = conectar_db()
    df_obra = pd.read_sql_query("SELECT * FROM piquetes", conn)
    conn.close()
    
    if df_obra.empty:
        st.info("No hay datos cargados en el sistema de control.")
    else:
        peso_por_hito = 100 / len(HITOS_OBRA)
        df_obra["Avance_%"] = 0
        for hito in HITOS_OBRA:
            df_obra["Avance_%"] += df_obra[hito].notna().astype(int) * peso_por_hito
        df_obra["Avance_%"] = df_obra["Avance_%"].round().astype(int)

        tramos_validos = [t for t in df_obra["tramo"].dropna().unique() if str(t).strip().lower() != "nan" and str(t).strip() != ""]
        
        idx_defecto = 0
        if st.session_state.proyecto_activo in tramos_validos:
            idx_defecto = tramos_validos.index(st.session_state.proyecto_activo)
            
        tramo_sel = st.selectbox("Filtrar Análisis por Frente/Tramo:", tramos_validos, index=idx_defecto)
        df_inv = df_obra[df_obra["tramo"] == tramo_sel].copy()
        
        total_piquetes_frente = len(df_inv)
        total_aisladores_frente = df_inv["cantidad_aisladores"].fillna(3).sum()
        aisladores_colocados_frente = df_inv[df_inv["montaje_aislador"].notna()]["cantidad_aisladores"].fillna(3).sum()
        
        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f"""<div class='kpi-card' style='border-left-color: #3b82f6;'><div class='kpi-title'>📍 Total Piquetes</div><div class='kpi-value'>{total_piquetes_frente}</div></div>""", unsafe_allow_html=True)
        with m2:
            st.markdown(f"""<div class='kpi-card' style='border-left-color: #8b5cf6;'><div class='kpi-title'>🔌 Aisladores Montados</div><div class='kpi-value'>{int(aisladores_colocados_frente)} / {int(total_aisladores_frente)}</div></div>""", unsafe_allow_html=True)
        with m3:
            st.markdown(f"""<div class='kpi-card' style='border-left-color: #10b981;'><div class='kpi-title'>📈 % Avance Físico Promedio</div><div class='kpi-value'>{int(df_inv["Avance_%"].mean() if total_piquetes_frente > 0 else 0)}%</div></div>""", unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### 📊 Auditoría y Estado de Parámetros por Columna")
        
        columnas_analizar = {
            "piquete": "📍 Nombre / Código de Piquete",
            "tipo_estructura": "🏗️ Tipos de Estructura",
            "longitud_poste": "📏 Longitud de Poste",
            "cantidad_aisladores": "🔌 Cantidad de Aisladores por Poste",
            "tipo_de_equipo": "⚙️ Equipos de Maniobra",
            "Avance_%": "📈 Porcentaje de Avance Individual"
        }
        
        col_sel = st.selectbox("Seleccione Parámetro a Graficar:", list(columnas_analizar.values()))
        col_real = [k for k, v in columnas_analizar.items() if v == col_sel][0]
        
        df_filtrado_grafico = df_inv[[col_real]].copy()
        df_filtrado_grafico[col_real] = df_filtrado_grafico[col_real].replace(["None", "nan", "-", ""], None)
        df_filtrado_grafico = df_filtrado_grafico.dropna()
        
        if df_filtrado_grafico.empty:
            st.warning("No se detectaron registros válidos cargados para este parámetro específico.")
        else:
            if col_real not in ["Avance_%", "cantidad_aisladores"]:
                df_filtrado_grafico[col_real] = df_filtrado_grafico[col_real].astype(str)
                
            df_frecuencia = df_filtrado_grafico[col_real].value_counts().reset_index()
            df_frecuencia.columns = [col_sel, "Cantidad de Piquetes"]
            
            if col_real in ["Avance_%", "cantidad_aisladores"]:
                df_frecuencia = df_frecuencia.sort_values(by=col_sel)
            
            g_col1, g_col2 = st.columns([3, 2])
            with g_col1:
                fig_inv = px.bar(df_frecuencia, x=col_sel, y="Cantidad de Piquetes", text="Cantidad de Piquetes", 
                                 color="Cantidad de Piquetes", color_continuous_scale="Viridis")
                fig_inv.update_layout(xaxis_type='category' if col_real not in ["Avance_%", "cantidad_aisladores"] else 'linear')
                st.plotly_chart(fig_inv, use_container_width=True)
            with g_col2:
                fig_donut = px.pie(df_frecuencia, names=col_sel, values="Cantidad de Piquetes", hole=0.4, 
                                   color_discrete_sequence=px.colors.qualitative.Safe)
                st.plotly_chart(fig_donut, use_container_width=True)
                
        with st.expander("🔍 Ver Listado Detallado de este Frente"):
            st.dataframe(df_inv[["piquete", "tipo_estructura", "longitud_poste", "cantidad_aisladores", "tipo_de_equipo", "Avance_%"]], use_container_width=True)

# -------------------------------------------------------------------------
# MÓDULO 4: CARGA Y GESTIÓN DE CAMPO
# -------------------------------------------------------------------------
elif opcion == "📝 Carga y Gestión de Campo":
    st.subheader("📝 Gestión Operativa y Certificación de Avances")
    
    conn = conectar_db()
    df_combos = pd.read_sql_query("SELECT tramo, piquete FROM piquetes", conn)
    conn.close()
    
    if df_combos.empty:
        st.info("Sin registros operativos. Por favor, cargue la planilla inicial en la pestaña de Migración.")
    else:
        tramos_fijos = [t for t in df_combos["tramo"].dropna().unique() if str(t).strip().lower() != "nan" and str(t).strip() != ""]
        
        idx_defecto = 0
        if st.session_state.proyecto_activo in tramos_fijos:
            idx_defecto = tramos_fijos.index(st.session_state.proyecto_activo)
            
        tramo_sel = st.selectbox("Seleccione Frente de Trabajo:", tramos_fijos, index=idx_defecto)
        piquetes_filtrados = df_combos[df_combos["tramo"] == tramo_sel]["piquete"].unique()
        piquete_sel = st.selectbox("Estructura / Piquete Específico:", piquetes_filtrados)
        
        conn = conectar_db()
        p_info = pd.read_sql_query("SELECT * FROM piquetes WHERE piquete = ?", conn, params=[piquete_sel]).iloc[0]
        conn.close()
        
        l_poste = p_info["longitud_poste"] if p_info["longitud_poste"] and p_info["longitud_poste"] != "None" else "N/A"
        cant_aisl_actual = int(p_info["cantidad_aisladores"]) if pd.notna(p_info["cantidad_aisladores"]) else 3
        
        st.info(f"📏 **LONGITUD POSTE:** {l_poste} | 🏗️ **ESTRUCTURA:** {p_info['tipo_estructura']} | 🔌 **AISLADORES REQUERIDOS:** {cant_aisl_actual}")
        
        st.markdown("### 📂 Documentación Actualizada del Piquete")
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            doc_anexo_actual = p_info["anexo_montaje"] if p_info["anexo_montaje"] and p_info["anexo_montaje"] != "None" else None
            if doc_anexo_actual and os.path.exists(os.path.join(CARPA_ARCHIVOS, doc_anexo_actual)):
                st.write(f"📄 **Anexo Técnico Activo:** `{doc_anexo_actual}`")
                with open(os.path.join(CARPA_ARCHIVOS, doc_anexo_actual), "rb") as file:
                    st.download_button(label="📥 Descargar Anexo Montaje", data=file, file_name=doc_anexo_actual, mime="application/octet-stream", key="dl_anexo")
            else:
                st.warning("⚠️ No hay ningún Anexo Técnico cargado.")
                
        with col_dl2:
            doc_red_actual = p_info["red_line"] if p_info["red_line"] and p_info["red_line"] != "None" else None
            if doc_red_actual and os.path.exists(os.path.join(CARPA_ARCHIVOS, doc_red_actual)):
                st.write(f"🗺️ **Plano Red Line Activo:** `{doc_red_actual}`")
                with open(os.path.join(CARPA_ARCHIVOS, doc_red_actual), "rb") as file:
                    st.download_button(label="📥 Descargar Red Line", data=file, file_name=doc_red_actual, mime="application/octet-stream", key="dl_redline")
            else:
                st.warning("⚠️ No hay ningún plano Red Line cargado.")
                
        st.markdown("---")
        
        with st.form("form_trazabilidad_avanzado"):
            def convertir_a_fecha(val):
                if val and val != "None" and val != "-" and val != "":
                    try: return pd.to_datetime(val).date()
                    except: return None
                return None

            st.markdown("##### ⚙️ Configuración y Cronograma de Hitos")
            
            c_aisl1, c_aisl2 = st.columns(2)
            with c_aisl1:
                cant_aisladores_input = st.number_input("🔌 Cantidad de Aisladores en esta Estructura:", min_value=1, max_value=20, value=cant_aisl_actual)
            
            st.markdown("---")
            col1, col2 = st.columns(2)
            with col1:
                f_excav = st.date_input("1. EXCAV PIQUETES", value=convertir_a_fecha(p_info["excavacion"]))
                f_vert = st.date_input("2. VERTICALIZADO", value=convertir_a_fecha(p_info["verticalizado"]))
                f_riendas = st.date_input("3. MONTAJE RIENDAS", value=convertir_a_fecha(p_info["montaje_riendas"]))
                f_aislador = st.date_input("4. MONTAJE DE AISLADOR (FECHA)", value=convertir_a_fecha(p_info["montaje_aislador"]))
            with col2:
                f_tendido = st.date_input("5. TENDIDO DE CONDUCTOR", value=convertir_a_fecha(p_info["tendido"]))
                f_flechado = st.date_input("6. FLECHADO", value=convertir_a_fecha(p_info["flechado"]))
                f_engramp = st.date_input("7. ENGRAMPADO", value=convertir_a_fecha(p_info["engrampado"]))

            st.markdown("---")
            f_montaje = st.date_input("Fecha Montaje / Liberación Final", value=convertir_a_fecha(p_info["fecha_montaje"]))

            st.markdown("##### 📥 Carga / Actualización de Documentos Técnicos")
            col_arch1, col_arch2 = st.columns(2)
            with col_arch1:
                archivo_anexo = st.file_uploader("Subir Nuevo ANEXO MONTAJE", type=["docx", "xlsx", "pdf", "xls"])
            with col_arch2:
                archivo_redline = st.file_uploader("Subir Nuevo RED LINE", type=["docx", "xlsx", "pdf", "xls"])

            if st.form_submit_button("💾 Actualizar Historial de Trazabilidad y Archivos"):
                nombre_anexo = p_info["anexo_montaje"]
                nombre_redline = p_info["red_line"]
                
                if archivo_anexo is not None:
                    nombre_anexo = f"Anexo_{piquete_sel}_{archivo_anexo.name}"
                    with open(os.path.join(CARPA_ARCHIVOS, nombre_anexo), "wb") as f:
                        f.write(archivo_anexo.getbuffer())
                        
                if archivo_redline is not None:
                    nombre_redline = f"RedLine_{piquete_sel}_{archivo_redline.name}"
                    with open(os.path.join(CARPA_ARCHIVOS, nombre_redline), "wb") as f:
                        f.write(archivo_redline.getbuffer())

                conn = conectar_db()
                conn.execute("""
                    UPDATE piquetes SET cantidad_aisladores=?, excavacion=?, verticalizado=?, montaje_riendas=?, montaje_aislador=?, tendido=?, flechado=?, engrampado=?, fecha_montaje=?, anexo_montaje=?, red_line=?
                    WHERE piquete=?
                """, (int(cant_aisladores_input), str(f_excav) if f_excav else None, str(f_vert) if f_vert else None, str(f_riendas) if f_riendas else None,
                      str(f_aislador) if f_aislador else None, str(f_tendido) if f_tendido else None, str(f_flechado) if f_flechado else None, 
                      str(f_engramp) if f_engramp else None, str(f_montaje) if f_montaje else None, 
                      str(nombre_anexo) if nombre_anexo else None, str(nombre_redline) if nombre_redline else None, piquete_sel))
                conn.commit()
                conn.close()
                
                st.session_state.proyecto_activo = tramo_sel
                st.success(f"✔️ Historial, aisladores ({cant_aisladores_input} ud) y documentación de {piquete_sel} actualizados correctamente.")
                st.rerun()

# -------------------------------------------------------------------------
# MÓDULO 1: ANALÍTICA AVANZADA Y KPIS
# -------------------------------------------------------------------------
else:
    conn = conectar_db()
    df_obra = pd.read_sql_query("SELECT * FROM piquetes", conn)
    df_cronogramas = pd.read_sql_query("SELECT * FROM cronogramas", conn)
    conn.close()
    
    if df_obra.empty:
        st.info("No existen registros de obra para procesar analíticas. Vaya al módulo de Migración.")
    else:
        for hito in HITOS_OBRA:
            df_obra[hito] = pd.to_datetime(df_obra[hito], errors='coerce')
            
        peso_por_hito = 100 / len(HITOS_OBRA)
        df_obra["Avance_%"] = 0
        for hito in HITOS_OBRA:
            df_obra["Avance_%"] += df_obra[hito].notna().astype(int) * peso_por_hito
        df_obra["Avance_%"] = df_obra["Avance_%"].round().astype(int)

        tramos_validos = [t for t in df_obra["tramo"].dropna().unique() if str(t).strip().lower() != "nan" and str(t).strip() != ""]
        
        idx_defecto = 0
        if st.session_state.proyecto_activo in tramos_validos:
            idx_defecto = tramos_validos.index(st.session_state.proyecto_activo)
            
        tramo_sel = st.selectbox("Frente Operativo / Proyecto Seleccionado:", tramos_validos, index=idx_defecto)
        df_tramo = df_obra[df_obra["tramo"] == tramo_sel]

        c_actual = df_cronogramas[df_cronogramas["tramo"] == tramo_sel]
        val_ini = pd.to_datetime(c_actual["inicio"].iloc[0]).date() if not c_actual.empty else datetime.date.today()
        val_ent = pd.to_datetime(c_actual["entrega"].iloc[0]).date() if not c_actual.empty else (datetime.date.today() + datetime.timedelta(days=60))

        with st.expander("⚙️ Ajustes Contractuales Avanzados de Plazos"):
            col_f1, col_f2, col_f3 = st.columns([2, 2, 1])
            with col_f1: f_inicio = st.date_input("Fecha Inicio Contractual", val_ini, key=f"i_{tramo_sel}")
            with col_f2: f_entrega = st.date_input("Fecha Fin de Obra", val_ent, key=f"e_{tramo_sel}")
            with col_f3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Fijar Parámetros"):
                    conn = conectar_db()
                    conn.execute("INSERT OR REPLACE INTO cronogramas (tramo, inicio, entrega) VALUES (?, ?, ?)", (tramo_sel, str(f_inicio), str(f_entrega)))
                    conn.commit()
                    conn.close()
                    st.session_state.proyecto_activo = tramo_sel
                    st.rerun()

        inicio_base = pd.to_datetime(f_inicio)
        entrega_base = pd.to_datetime(f_entrega)
        hoy = pd.to_datetime(datetime.date.today())
        
        dias_transcurridos = max(1, (hoy - inicio_base).days)
        avance_promedio = df_tramo["Avance_%"].mean() if not df_tramo.empty else 0

        if avance_promedio <= 0.5:
            ritmo_diario = 0.0
            fin_proyectado = entrega_base
            desviacion_dias = 0
        else:
            ritmo_diario = avance_promedio / dias_transcurridos
            dias_proyectados_totales = int(100 / ritmo_diario)
            dias_proyectados_totales = min(dias_proyectados_totales, 1825)
            fin_proyectado = inicio_base + pd.Timedelta(days=dias_proyectados_totales)
            desviacion_dias = (fin_proyectado - entrega_base).days
        
        hitos_completados = sum(df_tramo[hito].notna().sum() for hito in HITOS_OBRA) if not df_tramo.empty else 0
        productividad_media = round(hitos_completados / dias_transcurridos, 2)

        total_excavados = df_tramo["excavacion"].notna().sum() if not df_tramo.empty else 0
        total_verticalizados = df_tramo["verticalizado"].notna().sum() if not df_tramo.empty else 0
        total_tendidos = df_tramo["tendido"].notna().sum() if not df_tramo.empty else 0

        # --- TARJETAS KPIS SUPERIORES ---
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        
        with kpi1:
            st.markdown(f"""
                <div class='kpi-card' style='border-left-color: #3b82f6;'>
                    <div class='kpi-title'>Avance Físico Consolidado</div>
                    <div class='kpi-value'>{int(avance_promedio)}%</div>
                    <div class='kpi-delta' style='color: #a7f3d0;'>⚡ Eficiencia: {round(ritmo_diario, 2)}% / día</div>
                </div>
            """, unsafe_allow_html=True)

        with kpi2:
            st.markdown(f"""
                <div class='kpi-card' style='border-left-color: #10b981;'>
                    <div class='kpi-title'>Productividad de Obra</div>
                    <div class='kpi-value'>{productividad_media} <span style='font-size:14px;color:#9ca3af;'>hitos/día</span></div>
                    <div class='kpi-delta' style='color: #9ca3af;'>Total hitos logrados: {hitos_completados}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with kpi3:
            color_desv = "#f87171" if desviacion_dias > 0 else "#34d399"
            txt_desv = f"+ {desviacion_dias} días de retraso" if desviacion_dias > 0 else f"{abs(desviacion_dias)} días adelantado"
            st.markdown(f"""
                <div class='kpi-card' style='border-left-color: {color_desv};'>
                    <div class='kpi-title'>Desviación Contractual</div>
                    <div class='kpi-value'>{abs(desviacion_dias)} <span style='font-size:14px;color:#9ca3af;'>días</span></div>
                    <div class='kpi-delta' style='color: {color_desv};'>{txt_desv}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with kpi4:
            st.markdown(f"""
                <div class='kpi-card' style='border-left-color: #f59e0b;'>
                    <div class='kpi-title'>Proyección de Cierre Real</div>
                    <div class='kpi-value' style='font-size:22px; margin-top:12px;'>{fin_proyectado.strftime('%d/%m/%Y')}</div>
                    <div class='kpi-delta' style='color: #9ca3af;'>Plazo Contrato: {entrega_base.strftime('%d/%m/%Y')}</div>
                </div>
            """, unsafe_allow_html=True)

        # --- DIAGNÓSTICO OPERATIVO DE METAS ---
        st.markdown("<h3 style='color:#ffffff; font-size:18px; margin-top:20px;'>⚠️ Diagnóstico Operativo de Flujo</h3>", unsafe_allow_html=True)
        
        dias_restantes = max(0, (entrega_base - hoy).days)
        total_piquetes_tramo = len(df_tramo)
        piquetes_completados_100 = len(df_tramo[df_tramo["Avance_%"] == 100])
        piquetes_pendientes = max(0, total_piquetes_tramo - piquetes_completados_100)
        
        if piquetes_pendientes == 0:
            alerta_color = "#10b981"
            txt_meta = "0 piquetes"
            txt_sub = "¡Frente 100% finalizado!"
            ritmo_requerido_piquetes = 0.0
        elif dias_restantes <= 0:
            alerta_color = "#ef4444"
            txt_meta = f"{piquetes_pendientes} piquetes"
            txt_sub = "⚠️ Plazo contractual vencido."
            ritmo_requerido_piquetes = 999.0
        else:
            ritmo_requerido_piquetes = round(piquetes_pendientes / dias_restantes, 2)
            alerta_color = "#ef4444" if ritmo_requerido_piquetes > 2.0 else "#f59e0b"
            txt_meta = f"{ritmo_requerido_piquetes} piquetes/día"
            txt_sub = f"Faltan {piquetes_pendientes} piquetes en {dias_restantes} días rest."

        col_bot1, col_bot2 = st.columns([1, 2])
        with col_bot1:
            st.markdown(f"""
                <div class='kpi-card' style='border-left-color: {alerta_color}; background: #1e2230;'>
                    <div class='kpi-title' style='color: {alerta_color};'>🎯 Meta para Cumplir Plazo</div>
                    <div class='kpi-value' style='color: {alerta_color}; font-size:24px;'>{txt_meta}</div>
                    <div class='kpi-delta' style='color: #9ca3af;'>{txt_sub}</div>
                </div>
            """, unsafe_allow_html=True)
        
        with col_bot2:
            df_frentes = pd.DataFrame({
                "Frente Operativo": ["1. Excavación", "2. Verticalizado", "5. Tendido"],
                "Piquetes Completados": [total_excavados, total_verticalizados, total_tendidos]
            })
            fig_frentes = px.bar(df_frentes, x="Piquetes Completados", y="Frente Operativo", orientation='h',
                                 text="Piquetes Completados", color="Piquetes Completados", color_continuous_scale="Darkmint")
            fig_frentes.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=140, showlegend=False, coloraxis_showscale=False)
            st.plotly_chart(fig_frentes, use_container_width=True)

        # --- PROYECCIÓN BASADA EN RITMO REAL DE CARGA DIARIA ---
        st.markdown("<h3 style='color:#ffffff; font-size:18px; margin-top:25px;'>📈 Proyección de Cumplimiento por Ritmo Real de Campo</h3>", unsafe_allow_html=True)

        df_historico = df_tramo[df_tramo["Avance_%"] == 100].copy()
        df_historico["fecha_montaje"] = pd.to_datetime(df_historico["fecha_montaje"], errors='coerce')
        df_historico = df_historico.dropna(subset=["fecha_montaje"]).sort_values("fecha_montaje")

        col_sl1, col_sl2 = st.columns([2, 1])
        with col_sl1:
            dias_evaluacion = st.slider("Evaluación de ritmo reciente (días):", min_value=7, max_value=30, value=14, step=7)

        fecha_limite_eval = hoy - pd.Timedelta(days=dias_evaluacion)
        completados_recientes = len(df_historico[df_historico["fecha_montaje"] >= fecha_limite_eval])
        ritmo_real_diario = round(completados_recientes / dias_evaluacion, 2)

        fechas_futuras = [hoy + pd.Timedelta(days=i) for i in range(dias_restantes + 1)]

        incremento_meta = (total_piquetes_tramo - piquetes_completados_100) / dias_restantes if dias_restantes > 0 else 0
        curva_meta = [min(total_piquetes_tramo, piquetes_completados_100 + (incremento_meta * i)) for i in range(dias_restantes + 1)]
        curva_real = [min(total_piquetes_tramo, piquetes_completados_100 + (ritmo_real_diario * i)) for i in range(dias_restantes + 1)]

        df_proyeccion = pd.DataFrame({
            "Fecha": fechas_futuras + fechas_futuras,
            "Piquetes Finalizados": curva_meta + curva_real,
            "Trayectoria": ["Meta Necesaria (Contrato)"] * len(fechas_futuras) + [f"Tendencia Real ({ritmo_real_diario} piq/día)"] * len(fechas_futuras)
        })

        fig_proy = px.line(
            df_proyeccion, 
            x="Fecha", 
            y="Piquetes Finalizados", 
            color="Trayectoria",
            color_discrete_map={
                "Meta Necesaria (Contrato)": "#10b981", 
                f"Tendencia Real ({ritmo_real_diario} piq/día)": "#ef4444" if ritmo_real_diario < ritmo_requerido_piquetes else "#3b82f6"
            }
        )
        
        fig_proy.add_hline(y=total_piquetes_tramo, line_dash="dash", line_color="#9ca3af", annotation_text=f"Total Objetivo: {total_piquetes_tramo} piq.")
        fig_proy.update_layout(margin=dict(l=10, r=10, t=20, b=10), height=320)
        st.plotly_chart(fig_proy, use_container_width=True)

        col_summary1, col_summary2 = st.columns(2)
        with col_summary1:
            st.info(f"📊 **Ritmo Promedio Reciente ({dias_evaluacion} días):** `{ritmo_real_diario} piquetes/día` ({completados_recientes} piquetes listos)")
        with col_summary2:
            if ritmo_real_diario >= ritmo_requerido_piquetes:
                st.success(f"✅ **Proyección Cierre:** Manteniendo este ritmo, se alcanza la meta antes del `{entrega_base.strftime('%d/%m/%Y')}`.")
            else:
                dias_necesarios_extra = int((total_piquetes_tramo - piquetes_completados_100) / ritmo_real_diario) if ritmo_real_diario > 0 else 999
                fecha_est_cierre = hoy + pd.Timedelta(days=dias_necesarios_extra)
                st.error(f"🚨 **Proyección Cierre:** Al ritmo actual se finalizará el `{fecha_est_cierre.strftime('%d/%m/%Y')}`. Se requiere aumentar **+{round(ritmo_requerido_piquetes - ritmo_real_diario, 2)} piq/día**.")

        # --- SIMULACIÓN DE PLAZO CONTRACTUAL (GANTT) ---
        st.markdown("<h3 style='color:#ffffff; font-size:18px; margin-top:20px;'>📊 Simulación de Plazo Contractual vs Proyección de Ritmo Actual</h3>", unsafe_allow_html=True)
        df_gantt = pd.DataFrame([
            {"Línea de Tiempo": "Plazo Comprometido por Contrato", "Inicio": inicio_base, "Fin": entrega_base, "Condición": "Contrato Base"},
            {"Línea de Tiempo": "Proyección por Avance de Campo", "Inicio": inicio_base, "Fin": fin_proyectado, "Condición": "Proyección Real de Obra"}
        ])
        fig = px.timeline(df_gantt, x_start="Inicio", x_end="Fin", y="Línea de Tiempo", color="Condición", 
                          color_discrete_map={"Contrato Base": "#1e3a8a", "Proyección Real de Obra": "#b45309"})
        fig.update_yaxes(autorange="reversed", title="")
        fig.update_layout(margin=dict(l=20, r=20, t=10, b=20), height=180)
        st.plotly_chart(fig, use_container_width=True)

        # --- MATRIZ COMPLETA Y REPORTE ---
        st.markdown("<h3 style='color:#ffffff; font-size:18px; margin-top:20px;'>📋 Matriz Completa de Trazabilidad</h3>", unsafe_allow_html=True)
        df_mostrar = df_tramo.copy()
        for hito in HITOS_OBRA:
            df_mostrar[hito] = df_mostrar[hito].dt.strftime('%d/%m/%Y').fillna("-")
        
        columnas_visibles = ["tramo", "piquete", "tipo_estructura", "longitud_poste", "cantidad_aisladores", "Avance_%", "anexo_montaje", "red_line"] + HITOS_OBRA
        st.markdown("---")
        st.markdown("### 📥 Exportar y Notificar Reportes de Trazabilidad")
        
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
            df_mostrar[columnas_visibles].to_excel(writer, sheet_name=f"Progreso_{tramo_sel}", index=False)
        
        col_exp1, col_exp2 = st.columns(2)
        
        with col_exp1:
            st.download_button(
                label="📊 Descargar Base Actualizada (Excel)",
                data=buffer_excel.getvalue(),
                file_name=f"Trazabilidad_{tramo_sel}_{datetime.date.today().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
                use_container_width=True
            )
            
        with col_exp2:
            st.markdown(
                """
                <style>
                .btn-print {
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    background-color: #1f2937;
                    color: #ffffff;
                    padding: 0.5rem 1rem;
                    border-radius: 0.5rem;
                    border: 1px solid #4b5563;
                    cursor: pointer;
                    width: 100%;
                    height: 38px;
                    font-size: 14px;
                    font-family: inherit;
                    transition: background-color 0.2s;
                }
                .btn-print:hover {
                    background-color: #374151;
                    border-color: #6b7280;
                }
                </style>
                """, unsafe_allow_html=True
            )
            
            components.html(
                """
                <button class="btn-print" onclick="window.parent.print()">📄 Guardar Reporte / KPIs (PDF)</button>
                """,
                height=45
            )

        # --- ENVÍO DE CORREO ---
        st.markdown("---")
        with st.expander("✉️ Enviar Reporte de Avance por Correo Electrónico"):
            emails_input = st.text_input("Correos destinatarios (separados por coma):", placeholder="ejemplo@empresa.com, director@empresa.com")
            
            if st.button("🚀 Enviar Reporte Ahora", type="primary"):
                lista_correos = [e.strip() for e in emails_input.split(",") if e.strip()]
                if not lista_correos:
                    st.warning("Por favor ingresa al menos un correo válido.")
                else:
                    cuerpo_html = f"""
                    <h2>⚡ Reporte Automático de Control de Obra</h2>
                    <p>Se adjunta la planilla de trazabilidad actualizada para el frente <b>{tramo_sel}</b>.</p>
                    <ul>
                        <li><b>Fecha de Emisión:</b> {datetime.date.today().strftime('%d/%m/%Y')}</li>
                        <li><b>Avance Físico Consolidado:</b> {int(avance_promedio)}%</li>
                        <li><b>Meta Requerida:</b> {txt_meta} ({txt_sub})</li>
                        <li><b>Ritmo Real Actual:</b> {ritmo_real_diario} piquetes/día</li>
                        <li><b>Proyección Fin de Obra:</b> {fin_proyectado.strftime('%d/%m/%Y')}</li>
                    </ul>
                    <p><i>Reporte generado por el Panel de Control de Obra.</i></p>
                    """
                    
                    exito = enviar_reporte_correo(
                        destinatarios=lista_correos,
                        asunto=f"⚡ Reporte de Obra - Frente {tramo_sel} ({datetime.date.today().strftime('%d_%m_%Y')})",
                        cuerpo=cuerpo_html,
                        archivo_bytes=buffer_excel.getvalue(),
                        nombre_archivo=f"Trazabilidad_{tramo_sel}_{datetime.date.today().strftime('%d_%m_%Y')}.xlsx"
                    )
                    if exito:
                        st.success(f"✔️ ¡Reporte enviado con éxito a: {', '.join(lista_correos)}!")