import streamlit as st
import pandas as pd
import datetime
import sqlite3
import plotly.express as px
import os

# Configuración de página con tema premium expandido
st.set_page_config(layout="wide", page_title="Control de Obra Eléctrica Avanzado", page_icon="⚡")

# --- INICIALIZAR RUTA/ESTADO DE PROYECTO ---
if "proyecto_activo" not in st.session_state:
    st.session_state.proyecto_activo = None

# --- MENÚ LATERAL DIRECTO ---
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3227/3227840.png", width=70)
st.sidebar.markdown("👤 **Modo:** Acceso Total Abierto")

st.sidebar.title("Navegación del Sistema")
opcion = st.sidebar.radio("Ir a la pestaña:", [
    "📈 Analítica Avanzada y KPIs", 
    "📦 Inventario y Conteo de Columnas", 
    "📝 Carga y Gestión de Campo", 
    "📥 Migración Inicial (Excel)"
])

# Estilos personalizados CSS para modernizar el Dashboard
st.markdown("""
    <style>
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
    
    if len(columnas) > 0 and "tipo_de_equipo" not in columnas:
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
            excavacion TEXT,
            verticalizado TEXT,
            montaje_riendas TEXT,
            tendido TEXT,
            flechado TEXT,
            engrampado TEXT,
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

# CARPETA PARA DOCUMENTOS ADJUNTOS
CARPA_ARCHIVOS = "archivos_obra"
if not os.path.exists(CARPA_ARCHIVOS):
    os.makedirs(CARPA_ARCHIVOS)

# -------------------------------------------------------------------------
# MÓDULO 4: MIGRACIÓN INICIAL DESDE EXCEL
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
                    if "TRAMO - TAG" in valores_fila or "PIQUETE" in valores_fila or "ESTRUCTURA" in valores_fila:
                        skip_rows = i + 1
                        break
                
                archivo_excel.seek(0)
                df = pd.read_excel(archivo_excel, skiprows=skip_rows)
                df.columns = df.columns.str.strip().str.upper()
                
                col_piquete_encontrada = None
                posibles_nombres = ["PIQUETE", "PIQUETES", "NRO PIQUETE", "ESTRUCTURA", "TAG", "COD_PIQUETE"]
                for col in df.columns:
                    if col in posibles_nombres:
                        col_piquete_encontrada = col
                        break
                
                if not col_piquete_encontrada:
                    st.error(f"❌ No se pudo identificar la columna de Piquetes. Columnas detectadas: {list(df.columns)}")
                else:
                    df = df.dropna(subset=[col_piquete_encontrada])
                    
                    conn = conectar_db()
                    conn.execute("DELETE FROM piquetes WHERE tramo = ?", (nombre_proyecto_manual,))
                    
                    registros_cargados = 0
                    for _, row in df.iterrows():
                        piquete_val = str(row.get(col_piquete_encontrada, "")).strip()
                        if piquete_val and piquete_val.lower() != "nan" and piquete_val != "":
                            conn.execute("""
                                INSERT OR REPLACE INTO piquetes (tramo, piquete, tipo_estructura, excavacion, verticalizado, 
                                                    montaje_riendas, tendido, flechado, engrampado, fecha_montaje,
                                                    tipo_de_equipo, anexo_montaje, red_line, observacion_ofm)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                nombre_proyecto_manual, piquete_val, str(row.get("TIPO ESTRUCTURA", "S/D")),
                                str(row.get("EXCAV PIQUETES", "")) if pd.notna(row.get("EXCAV PIQUETES")) else None,
                                str(row.get("VERTICALIZADO", "")) if pd.notna(row.get("VERTICALIZADO")) else None,
                                str(row.get("MONTAJE RIENDAS", "")) if pd.notna(row.get("MONTAJE RIENDAS")) else None,
                                str(row.get("TENDIDO", "")) if pd.notna(row.get("TENDIDO")) else None,
                                str(row.get("FLECHADO", "")) if pd.notna(row.get("FLECHADO")) else None,
                                str(row.get("ENGRAMPADO", "")) if pd.notna(row.get("ENGRAMPADO")) else None,
                                str(row.get("FECHA DE MONTAJE", "")) if pd.notna(row.get("FECHA DE MONTAJE")) else None,
                                str(row.get("TIPO DE EQUIPO", "")) if pd.notna(row.get("TIPO DE EQUIPO")) else None,
                                str(row.get("ANEXO MONTAJE", "")) if pd.notna(row.get("ANEXO MONTAJE")) else None,
                                str(row.get("RED LINE", "")) if pd.notna(row.get("RED LINE")) else None,
                                str(row.get("OBSERVACION - OFM", "")) if pd.notna(row.get("OBSERVACION - OFM")) else None
                            ))
                            registros_cargados += 1
                                
                    conn.commit()
                    conn.close()
                    
                    if registros_cargados > 0:
                        st.session_state.proyecto_activo = nombre_proyecto_manual
                        st.success(f"✔️ ¡Migración exitosa! Se guardaron {registros_cargados} piquetes.")
                        st.rerun()
            except Exception as e:
                st.error(f"❌ Error al procesar el Excel: {e}")

# -------------------------------------------------------------------------
# MÓDULO 2: INVENTARIO Y CONTEO POR COLUMNAS (CORREGIDO Y DINÁMICO)
# -------------------------------------------------------------------------
elif opcion == "📦 Inventario y Conteo de Columnas":
    st.subheader("📦 Métricas de Inventario y Control de Columnas")
    
    conn = conectar_db()
    df_obra = pd.read_sql_query("SELECT * FROM piquetes", conn)
    conn.close()
    
    if df_obra.empty:
        st.info("No hay datos cargados en el sistema de control.")
    else:
        # Calcular el porcentaje de avance por piquete para graficarlo aquí también
        hitos = ["excavacion", "verticalizado", "montaje_riendas", "tendido", "flechado", "engrampado"]
        peso_por_hito = 100 / len(hitos)
        df_obra["Avance_%"] = 0
        for hito in hitos:
            df_obra["Avance_%"] += df_obra[hito].notna().astype(int) * peso_por_hito
        df_obra["Avance_%"] = df_obra["Avance_%"].round().astype(int)

        tramos_validos = [t for t in df_obra["tramo"].dropna().unique() if str(t).strip().lower() != "nan" and str(t).strip() != ""]
        
        idx_defecto = 0
        if st.session_state.proyecto_activo in tramos_validos:
            idx_defecto = tramos_validos.index(st.session_state.proyecto_activo)
            
        tramo_sel = st.selectbox("Filtrar Análisis por Frente/Tramo:", tramos_validos, index=idx_defecto)
        df_inv = df_obra[df_obra["tramo"] == tramo_sel].copy()
        
        # --- NUEVA SECCIÓN DE MÉTRICAS GLOBALES DEL FRENTE ---
        total_piquetes_frente = len(df_inv)
        avance_medio_frente = df_inv["Avance_%"].mean() if total_piquetes_frente > 0 else 0
        
        m1, m2 = st.columns(2)
        with m1:
            st.markdown(f"""<div class='kpi-card' style='border-left-color: #3b82f6;'><div class='kpi-title'>📍 Total Piquetes en Frente {tramo_sel}</div><div class='kpi-value'>{total_piquetes_frente}</div></div>""", unsafe_allow_html=True)
        with m2:
            st.markdown(f"""<div class='kpi-card' style='border-left-color: #10b981;'><div class='kpi-title'>📈 Avance Físico Promedio del Frente</div><div class='kpi-value'>{int(avance_medio_frente)}%</div></div>""", unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### 📊 Auditoría y Estado de Parámetros por Columna")
        
        columnas_analizar = {
            "piquete": "📍 Nombre / Código de Piquete",
            "tipo_estructura": "🏗️ Tipos de Estructura",
            "tipo_de_equipo": "⚙️ Equipos de Maniobra",
            "Avance_%": "📈 Porcentaje de Avance Individual"
        }
        
        col_sel = st.selectbox("Seleccione Parámetro a Graficar:", list(columnas_analizar.values()))
        col_real = [k for k, v in columnas_analizar.items() if v == col_sel][0]
        
        # Limpieza flexible de nulos para gráficos de distribución
        df_filtrado_grafico = df_inv[[col_real]].copy()
        df_filtrado_grafico[col_real] = df_filtrado_grafico[col_real].replace(["None", "nan", "-", ""], None)
        df_filtrado_grafico = df_filtrado_grafico.dropna()
        
        if df_filtrado_grafico.empty:
            st.warning("No se detectaron registros válidos cargados para este parámetro específico.")
        else:
            # Forzar tipo string para el conteo categórico (excepto si es el número de avance)
            if col_real != "Avance_%":
                df_filtrado_grafico[col_real] = df_filtrado_grafico[col_real].astype(str)
                
            df_frecuencia = df_filtrado_grafico[col_real].value_counts().reset_index()
            df_frecuencia.columns = [col_sel, "Cantidad de Piquetes"]
            
            # Ordenar por porcentaje si es la columna de avance
            if col_real == "Avance_%":
                df_frecuencia = df_frecuencia.sort_values(by=col_sel)
            
            g_col1, g_col2 = st.columns([3, 2])
            with g_col1:
                fig_inv = px.bar(df_frecuencia, x=col_sel, y="Cantidad de Piquetes", text="Cantidad de Piquetes", 
                                 color="Cantidad de Piquetes", color_continuous_scale="Viridis")
                fig_inv.update_layout(xaxis_type='category' if col_real != "Avance_%" else 'linear')
                st.plotly_chart(fig_inv, width='stretch')
            with g_col2:
                fig_donut = px.pie(df_frecuencia, names=col_sel, values="Cantidad de Piquetes", hole=0.4, 
                                   color_discrete_sequence=px.colors.qualitative.Safe)
                st.plotly_chart(fig_donut, width='stretch')
                
        # Mostrar la lista plana de control abajo para verificación visual rápida
        with st.expander("🔍 Ver Listado Detallado de este Frente"):
            st.dataframe(df_inv[["piquete", "tipo_estructura", "tipo_de_equipo", "Avance_%"]], width='stretch')

# -------------------------------------------------------------------------
# MÓDULO 3: CARGA Y GESTIÓN DE CAMPO
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
        
        # Volvemos a leer de la base de datos para obtener el estado más fresco antes del formulario
        conn = conectar_db()
        p_info = pd.read_sql_query("SELECT * FROM piquetes WHERE piquete = ?", conn, params=[piquete_sel]).iloc[0]
        conn.close()
        
        eq_excel = p_info["tipo_de_equipo"] if p_info["tipo_de_equipo"] and p_info["tipo_de_equipo"] != "None" else "No especificado en Excel"
        st.info(f"⚙️ **TIPO DE EQUIPO (Importado del Excel):** {eq_excel}")
        
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

            st.markdown("##### 📅 Cronograma de Hitos del Piquete")
            col1, col2 = st.columns(2)
            with col1:
                f_excav = st.date_input("1. EXCAV PIQUETES", value=convertir_a_fecha(p_info["excavacion"]))
                f_vert = st.date_input("2. VERTICALIZADO", value=convertir_a_fecha(p_info["verticalizado"]))
                f_riendas = st.date_input("3. MONTAJE RIENDAS", value=convertir_a_fecha(p_info["montaje_riendas"]))
            with col2:
                f_tendido = st.date_input("4. TENDIDO", value=convertir_a_fecha(p_info["tendido"]))
                f_flechado = st.date_input("5. FLECHADO", value=convertir_a_fecha(p_info["flechado"]))
                f_engramp = st.date_input("6. ENGRAMPADO", value=convertir_a_fecha(p_info["engrampado"]))

            st.markdown("---")
            f_montaje = st.date_input("Fecha Montaje Mecánico Final", value=convertir_a_fecha(p_info["fecha_montaje"]))

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
                    UPDATE piquetes SET excavacion=?, verticalizado=?, montaje_riendas=?, tendido=?, flechado=?, engrampado=?, fecha_montaje=?, anexo_montaje=?, red_line=?
                    WHERE piquete=?
                """, (str(f_excav) if f_excav else None, str(f_vert) if f_vert else None, str(f_riendas) if f_riendas else None,
                      str(f_tendido) if f_tendido else None, str(f_flechado) if f_flechado else None, str(f_engramp) if f_engramp else None,
                      str(f_montaje) if f_montaje else None, str(nombre_anexo) if nombre_anexo else None, str(nombre_redline) if nombre_redline else None, piquete_sel))
                conn.commit()
                conn.close()
                
                st.session_state.proyecto_activo = tramo_sel
                st.success(f"✔️ Historial y documentación de {piquete_sel} actualizados correctamente.")
                # st.rerun() fuerza a Streamlit a redibujar la pantalla inmediatamente leyendo los datos nuevos
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
        hitos = ["excavacion", "verticalizado", "montaje_riendas", "tendido", "flechado", "engrampado"]
        for hito in hitos:
            df_obra[hito] = pd.to_datetime(df_obra[hito], errors='coerce')
            
        peso_por_hito = 100 / len(hitos)
        df_obra["Avance_%"] = 0
        for hito in hitos:
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
        
        hitos_completados = sum(df_tramo[hito].notna().sum() for hito in hitos) if not df_tramo.empty else 0
        productividad_media = round(hitos_completados / dias_transcurridos, 2)

        total_excavados = df_tramo["excavacion"].notna().sum() if not df_tramo.empty else 0
        total_verticalizados = df_tramo["verticalizado"].notna().sum() if not df_tramo.empty else 0
        total_tendidos = df_tramo["tendido"].notna().sum() if not df_tramo.empty else 0
        
        descalce_civil_montaje = max(0, total_excavados - total_verticalizados)

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

        st.markdown("<h3 style='color:#ffffff; font-size:18px; margin-top:20px;'>⚠️ Diagnóstico Operativo de Flujo</h3>", unsafe_allow_html=True)
        col_bot1, col_bot2 = st.columns([1, 2])
        with col_bot1:
            color_alerta_bot = "#ef4444" if descalce_civil_montaje > 5 else "#f59e0b" if descalce_civil_montaje > 0 else "#10b981"
            st.markdown(f"""
                <div class='kpi-card' style='border-left-color: {color_alerta_bot}; background: #1e2230;'>
                    <div class='kpi-title' style='color: #ef4444;'>🚨 Cuello de Botella: Civil vs Izado</div>
                    <div class='kpi-value' style='color: {color_alerta_bot};'>{descalce_civil_montaje} <span style='font-size:14px;color:#9ca3af;'>piquetes</span></div>
                    <div class='kpi-delta' style='color: #9ca3af;'>Pozos excavados esperando estructura.</div>
                </div>
            """, unsafe_allow_html=True)
        
        with col_bot2:
            df_frentes = pd.DataFrame({
                "Frente Operativo": ["1. Excavación", "2. Verticalizado", "4. Tendido"],
                "Piquetes Completados": [total_excavados, total_verticalizados, total_tendidos]
            })
            fig_frentes = px.bar(df_frentes, x="Piquetes Completados", y="Frente Operativo", orientation='h',
                                 text="Piquetes Completados", color="Piquetes Completados", color_continuous_scale="Darkmint")
            fig_frentes.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=140, showlegend=False, coloraxis_showscale=False)
            st.plotly_chart(fig_frentes, width='stretch')

        st.markdown("<h3 style='color:#ffffff; font-size:18px; margin-top:20px;'>📊 Simulación de Plazo Contractual vs Proyección de Ritmo Actual</h3>", unsafe_allow_html=True)
        df_gantt = pd.DataFrame([
            {"Línea de Tiempo": "Plazo Comprometido por Contrato", "Inicio": inicio_base, "Fin": entrega_base, "Condición": "Contrato Base"},
            {"Línea de Tiempo": "Proyección por Avance de Campo", "Inicio": inicio_base, "Fin": fin_proyectado, "Condición": "Proyección Real de Obra"}
        ])
        fig = px.timeline(df_gantt, x_start="Inicio", x_end="Fin", y="Línea de Tiempo", color="Condición", 
                          color_discrete_map={"Contrato Base": "#1e3a8a", "Proyección Real de Obra": "#b45309"})
        fig.update_yaxes(autorange="reversed", title="")
        fig.update_layout(margin=dict(l=20, r=20, t=10, b=20), height=180)
        st.plotly_chart(fig, width='stretch')

        st.markdown("<h3 style='color:#ffffff; font-size:18px; margin-top:20px;'>📋 Matriz Completa de Trazabilidad</h3>", unsafe_allow_html=True)
        df_mostrar = df_tramo.copy()
        for hito in hitos:
            df_mostrar[hito] = df_mostrar[hito].dt.strftime('%d/%m/%Y').fillna("-")
        
        columnas_visibles = ["tramo", "piquete", "tipo_estructura", "Avance_%", "anexo_montaje", "red_line"] + hitos
        # --- SECCIÓN DE EXPORTACIÓN DE REPORTES CORREGIDA ---
        st.markdown("---")
        st.markdown("### 📥 Exportar Reportes de Trazabilidad")
        
        # 1. Preparación del búfer en memoria para el archivo Excel
        import io
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
            # Alternativa nativa y segura: Código HTML/JS embebido en un componente ligero
            import streamlit.components.v1 as components
            
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
            
            # Renderiza un botón real que llama a la ventana de impresión del sistema de forma segura
            components.html(
                """
                <button class="btn-print" onclick="window.parent.print()">📄 Guardar Reporte / KPIs (PDF)</button>
                """,
                height=45
            )