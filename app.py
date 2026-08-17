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
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

DRIVE_FOLDER_ID = "1Gv2m6_uBDjQkM4QluwO_Zmhjdw4LM85X"
DB_FILE_NAME = "obra_trazabilidad.db"

# --- FUNCIONES GOOGLE DRIVE CON MANEJO MEJORADO DE ERRORES ---
def get_drive_service():
    if "gcp_service_account" not in st.secrets:
        return None
    creds_dict = dict(st.secrets["gcp_service_account"])
    
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

    creds = service_account.Credentials.from_service_account_info(
        creds_dict, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

def download_db_from_drive():
    try:
        service = get_drive_service()
        if not service:
            return
        query = f"name = '{DB_FILE_NAME}' and trashed = false"
        if DRIVE_FOLDER_ID:
            query += f" and '{DRIVE_FOLDER_ID}' in parents"
            
        results = service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])
        
        if files:
            file_id = files[0]['id']
            request = service.files().get_media(fileId=file_id)
            with open(DB_FILE_NAME, 'wb') as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
            print("Base de datos descargada de Drive correctamente.")
    except Exception as e:
        print(f"Error al descargar de Drive: {e}")

def upload_db_to_drive():
    try:
        service = get_drive_service()
        if not service:
            return False
        query = f"name = '{DB_FILE_NAME}' and trashed = false"
        if DRIVE_FOLDER_ID:
            query += f" and '{DRIVE_FOLDER_ID}' in parents"
            
        results = service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])
        
        media = MediaFileUpload(DB_FILE_NAME, mimetype='application/x-sqlite3', resumable=True)
        
        if files:
            file_id = files[0]['id']
            archivo = service.files().update(fileId=file_id, media_body=media).execute()
        else:
            file_metadata = {'name': DB_FILE_NAME}
            if DRIVE_FOLDER_ID:
                file_metadata['parents'] = [DRIVE_FOLDER_ID]
            archivo = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        
        st.success(f"✔️ Backup sincronizado en Google Drive con éxito (ID: {archivo.get('id')}).")
        return True
    except Exception as e:
        st.error(f"❌ Error al sincronizar con Google Drive: {e}")
        return False

# Sincronizar descarga inicial al arrancar la app
if "db_descargada" not in st.session_state:
    download_db_from_drive()
    st.session_state.db_descargada = True

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(layout="wide", page_title="Control de Obra Eléctrica Avanzado", page_icon="⚡")

# --- INICIALIZAR RUTA/ESTADO DE PROYECTO ---
if "proyecto_activo" not in st.session_state:
    st.session_state.proyecto_activo = None

# --- LISTA OFICIAL DE LAS 9 ETAPAS OPERATIVAS (HITOS) ---
HITOS_OBRA = [
    "excavacion", 
    "verticalizado", 
    "desfile_de_poste",
    "montaje_riendas", 
    "armado_de_crucetas",
    "montaje_aislador", 
    "tendido", 
    "flechado", 
    "engrampado"
]

NOMBRES_HITOS = {
    "excavacion": "1. Excavación de Piquetes",
    "verticalizado": "2. Verticalizado",
    "desfile_de_poste": "3. Desfile de Poste",
    "montaje_riendas": "4. Montaje de Riendas",
    "armado_de_crucetas": "5. Armado de Crucetas",
    "montaje_aislador": "6. Montaje de Aislador",
    "tendido": "7. Tendido de Conductor",
    "flechado": "8. Flechado",
    "engrampado": "9. Engrampado"
}

def calcular_avance_piquete(row_piquete):
    hitos_validos = 0
    hitos_completados = 0
    
    for hito in HITOS_OBRA:
        val = str(row_piquete.get(hito, "")).strip().upper()
        if val in ["N/A", "NO APLICA", "NA", "N/D"]:
            continue
        hitos_validos += 1
        if val and val not in ["NONE", "NAN", "", "NAT", "-"]:
            hitos_completados += 1
            
    if hitos_validos == 0:
        return 100.0 if hitos_completados == 0 else 0.0
        
    return round((hitos_completados / hitos_validos) * 100, 2)

def normalizar_fecha(val):
    if pd.isna(val) or val is None:
        return None
    val_str = str(val).strip()
    if val_str.upper() in ["N/A", "NO APLICA", "NA", "N/D"]:
        return "N/A"
    if val_str.lower() in ["nan", "none", "", "nat", "-"]:
        return None
    try:
        dt = pd.to_datetime(val_str, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.strftime('%Y-%m-%d')
    except Exception:
        pass
    return val_str

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

# Estilos personalizados CSS
st.markdown("""
    <style>
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 1rem !important;
    }
    .stApp {
        background-color: #0b0f17;
    }
    .kpi-card {
        background: linear-gradient(145deg, #131b2a 0%, #0d131f 100%);
        padding: 18px 20px;
        border-radius: 12px;
        border-left: 5px solid #3b82f6;
        border-top: 1px solid rgba(255, 255, 255, 0.05);
        border-right: 1px solid rgba(255, 255, 255, 0.05);
        border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.4);
        margin-bottom: 12px;
    }
    .kpi-title {
        color: #94a3b8;
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
    }
    .kpi-value {
        color: #f8fafc;
        font-size: 26px;
        font-weight: 800;
        margin-top: 4px;
    }
    .kpi-delta {
        font-size: 12px;
        margin-top: 4px;
        font-weight: 600;
    }
    .status-card {
        background: #111827;
        padding: 16px;
        border-radius: 10px;
        border: 1px solid #1f2937;
        height: 100%;
    }
    .hito-card {
        background: #1e293b;
        padding: 12px 15px;
        border-radius: 8px;
        border-left: 4px solid #38bdf8;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("⚡ Panel de Control de Obra Eléctrica Avanzado")
st.markdown("---")

def conectar_db():
    return sqlite3.connect(DB_FILE_NAME)

def inicializar_db():
    conn = conectar_db()
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(piquetes)")
    columnas = [col[1] for col in cursor.fetchall()]
    
    if len(columnas) > 0 and ("cabezal" not in columnas or "armado_de_crucetas" not in columnas):
        conn.close()
        try: 
            os.remove(DB_FILE_NAME)
        except: 
            pass
        conn = conectar_db()
        cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS piquetes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tramo TEXT,
            piquete TEXT UNIQUE,
            tipo_estructura TEXT,
            cabezal TEXT,
            longitud_poste TEXT,
            cantidad_aisladores INTEGER DEFAULT 0,
            metros_tendido REAL DEFAULT 0,
            m3_excavacion REAL DEFAULT 0,
            excavacion TEXT,
            verticalizado TEXT,
            desfile_de_poste TEXT,
            montaje_riendas TEXT,
            armado_de_crucetas TEXT,
            montaje_aislador TEXT,
            tendido TEXT,
            flechado TEXT,
            engrampado TEXT,
            fecha_montaje TEXT,
            tipo_de_equipo TEXT,
            anexo_montaje TEXT,
            idi TEXT,
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
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS metas_ritmo (
            tramo TEXT PRIMARY KEY,
            ritmo_objetivo REAL,
            fecha_meta TEXT
        )
    """)
    conn.commit()
    conn.close()

inicializar_db()

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
    nombre_proyecto_manual = st.text_input("Ingrese el Nombre del Proyecto / Frente (Ej: WM, LTM):", value="WM-LMT2026", key="input_nombre_proyecto")
    
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
                    valores_fila = [str(val).strip().upper().replace("_", " ") for val in row.values if pd.notna(val)]
                    if "PIQUETE" in valores_fila or "TIPO ESTRUCTURA" in valores_fila:
                        skip_rows = i
                        break
                
                archivo_excel.seek(0)
                df = pd.read_excel(archivo_excel, skiprows=skip_rows)
                
                df.columns = df.columns.astype(str).str.strip().str.upper().str.replace("_", " ")
                df = df.loc[:, ~df.columns.duplicated()]
                
                if "PIQUETE" not in df.columns:
                    st.error(f"❌ No se encontró la columna 'PIQUETE'. Columnas detectadas: {list(df.columns)}")
                else:
                    df = df.dropna(subset=["PIQUETE"])
                    
                    conn = conectar_db()
                    conn.execute("DELETE FROM piquetes WHERE tramo = ?", (nombre_proyecto_manual,))
                    
                    def get_val(row, *col_names):
                        for c in col_names:
                            if c in row.index:
                                val = row.get(c)
                                if isinstance(val, pd.Series):
                                    val = val.iloc[0]
                                if pd.notna(val):
                                    val_str = str(val).strip()
                                    if val_str.lower() not in ["nan", "none", "", "nat"]:
                                        return val_str
                        return None

                    registros_cargados = 0
                    for _, row in df.iterrows():
                        piquete_val = get_val(row, "PIQUETE")
                        if piquete_val:
                            cant_aisl_raw = get_val(row, "CANTIDAD AISLADORES", "CANT AISLADORES", "CANT. AISLADORES", "AISLADORES", "CANT AISLADOR", "CANTIDAD DE AISLADORES", "AISLADOR")
                            if cant_aisl_raw is not None and str(cant_aisl_raw).strip() != "":
                                try:
                                    val_clean = str(cant_aisl_raw).replace(',', '.').strip()
                                    cant_aisl = int(float(val_clean))
                                except (ValueError, TypeError):
                                    cant_aisl = 0
                            else:
                                cant_aisl = 0

                            m_tend = get_val(row, "METROS TENDIDO", "METROS", "VANO", "DISTANCIA")
                            try:
                                m_tend = float(str(m_tend).replace(',', '.').strip()) if m_tend is not None else 0.0
                            except (ValueError, TypeError):
                                m_tend = 0.0

                            m3_exc = get_val(row, "M3 EXCAVACION", "VOLUMEN EXCAVACION", "EXCAVACION M3")
                            try:
                                m3_exc = float(str(m3_exc).replace(',', '.').strip()) if m3_exc is not None else 0.0
                            except (ValueError, TypeError):
                                m3_exc = 0.0

                            f_excav = normalizar_fecha(get_val(row, "1. EXCAV PIQUETES", "EXCAV PIQUETES", "EXCAVACION"))
                            f_vert = normalizar_fecha(get_val(row, "2. VERTICALIZADO", "VERTICALIZADO"))
                            f_desf = normalizar_fecha(get_val(row, "3. DESFILE DE POSTE", "DESFILE DE POSTE", "DESFILE"))
                            f_riend = normalizar_fecha(get_val(row, "4. MONTAJE RIENDAS", "MONTAJE RIENDAS", "RIENDAS"))
                            f_cruc = normalizar_fecha(get_val(row, "5. ARMADO DE CRUCETAS", "ARMADO DE CRUCETAS", "ARMADO DE CRUCETA", "CRUCETAS"))
                            f_aisl = normalizar_fecha(get_val(row, "6. MONTAJE DE AISLADOR (FECHA)", "MONTAJE DE AISLADOR", "MONTAJE AISLADOR"))
                            f_tend = normalizar_fecha(get_val(row, "7. TENDIDO DE CONDUCTOR", "TENDIDO DE CONDUCTOR", "TENDIDO"))
                            f_flec = normalizar_fecha(get_val(row, "8. FLECHADO", "FLECHADO"))
                            f_engr = normalizar_fecha(get_val(row, "9. ENGRAMPADO", "ENGRAMPADO"))
                            f_liber = normalizar_fecha(get_val(row, "10. FECHA MONTAJE / LIBERACIÓN FINAL", "FECHA DE LIBERACION", "FECHA MONTAJE"))

                            conn.execute("""
                                INSERT OR REPLACE INTO piquetes (
                                    tramo, piquete, tipo_estructura, cabezal, longitud_poste, cantidad_aisladores, metros_tendido, m3_excavacion, excavacion, verticalizado, 
                                    desfile_de_poste, montaje_riendas, armado_de_crucetas, montaje_aislador, tendido, flechado, engrampado, fecha_montaje, observacion_ofm
                                )
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                nombre_proyecto_manual, piquete_val, get_val(row, "TIPO ESTRUCTURA") or "S/D",
                                get_val(row, "CABEZAL") or "S/D", get_val(row, "LONGITUD POSTE"), cant_aisl, m_tend, m3_exc,
                                f_excav, f_vert, f_desf, f_riend, f_cruc, f_aisl, f_tend, f_flec, f_engr, f_liber,
                                get_val(row, "OBSERVACION", "OBSERVACION OFM")
                            ))
                            registros_cargados += 1
                                
                    conn.commit()
                    conn.close()
                    
                    if registros_cargados > 0:
                        upload_db_to_drive()
                        st.session_state.proyecto_activo = nombre_proyecto_manual
                        st.success(f"✔️ ¡Migración exitosa! Se procesaron {registros_cargados} piquetes correctamente.")
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
        st.session_state.proyecto_activo = proyecto_sel

        if proyecto_sel:
            df_proyecto = pd.read_sql_query("SELECT * FROM piquetes WHERE tramo = ?", conn, params=(proyecto_sel,))
            df_proyecto["Avance_%"] = df_proyecto.apply(calcular_avance_piquete, axis=1)

            total_piquetes = len(df_proyecto)
            avance_promedio = int(df_proyecto["Avance_%"].mean()) if total_piquetes > 0 else 0
            completados = len(df_proyecto[df_proyecto["Avance_%"] >= 99.9])

            total_aisladores = pd.to_numeric(df_proyecto["cantidad_aisladores"], errors='coerce').fillna(0).astype(int).sum()
            aisladores_instalados = pd.to_numeric(df_proyecto[df_proyecto["montaje_aislador"].notna() & (~df_proyecto["montaje_aislador"].astype(str).str.upper().isin(["N/A", "NO APLICA"]))]["cantidad_aisladores"], errors='coerce').fillna(0).astype(int).sum()

            total_metros = df_proyecto["metros_tendido"].fillna(0).sum()
            metros_tendidos = df_proyecto[df_proyecto["tendido"].notna() & (~df_proyecto["tendido"].astype(str).str.upper().isin(["N/A", "NO APLICA"]))]["metros_tendido"].fillna(0).sum()

            total_m3 = df_proyecto["m3_excavacion"].fillna(0).sum()
            m3_excavados = df_proyecto[df_proyecto["excavacion"].notna() & (~df_proyecto["excavacion"].astype(str).str.upper().isin(["N/A", "NO APLICA"]))]["m3_excavacion"].fillna(0).sum()

            col1, col2, col3, col4, col5, col6 = st.columns(6)
            with col1:
                st.markdown(f"""<div class='kpi-card' style='border-left-color: #3b82f6;'><div class='kpi-title'>📍 Total Piquetes</div><div class='kpi-value'>{total_piquetes}</div></div>""", unsafe_allow_html=True)
            with col2:
                st.markdown(f"""<div class='kpi-card' style='border-left-color: #10b981;'><div class='kpi-title'>✅ Piquetes 100%</div><div class='kpi-value'>{completados}</div></div>""", unsafe_allow_html=True)
            with col3:
                st.markdown(f"""<div class='kpi-card' style='border-left-color: #f59e0b;'><div class='kpi-title'>📈 % Avance Físico</div><div class='kpi-value'>{avance_promedio}%</div></div>""", unsafe_allow_html=True)
            with col4:
                st.markdown(f"""<div class='kpi-card' style='border-left-color: #8b5cf6;'><div class='kpi-title'>🔌 Aisladores Montados</div><div class='kpi-value'>{int(aisladores_instalados)} / {int(total_aisladores)}</div></div>""", unsafe_allow_html=True)
            with col5:
                st.markdown(f"""<div class='kpi-card' style='border-left-color: #06b6d4;'><div class='kpi-title'>📏 Metros Tendidos</div><div class='kpi-value'>{int(metros_tendidos)} / {int(total_metros)} m</div></div>""", unsafe_allow_html=True)
            with col6:
                st.markdown(f"""<div class='kpi-card' style='border-left-color: #eab308;'><div class='kpi-title'>⛏️ Excavación (m³)</div><div class='kpi-value'>{m3_excavados:.1f} / {total_m3:.1f} m³</div></div>""", unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("### 🛠️ Estado de los 9 Hitos Operativos")
            
            cols_hitos = st.columns(3)
            for idx, hito in enumerate(HITOS_OBRA):
                piquetes_aplicables = df_proyecto[~df_proyecto[hito].astype(str).str.upper().isin(["N/A", "NO APLICA", "NA", "N/D"])]
                total_aplicable = len(piquetes_aplicables)
                cant_completada = piquetes_aplicables[hito].notna().sum()
                pct_hito = (cant_completada / total_aplicable * 100) if total_aplicable > 0 else 0
                
                with cols_hitos[idx % 3]:
                    st.markdown(f"""
                        <div class='hito-card'>
                            <div style='font-size: 13px; font-weight: 700; color: #38bdf8;'>{NOMBRES_HITOS[hito]}</div>
                            <div style='font-size: 20px; font-weight: 800; color: #ffffff; margin-top:2px;'>{cant_completada} / {total_aplicable} <span style='font-size:13px; color:#94a3b8;'>({int(pct_hito)}%)</span></div>
                        </div>
                    """, unsafe_allow_html=True)
                    st.progress(int(pct_hito) / 100)

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
        df_obra["Avance_%"] = df_obra.apply(calcular_avance_piquete, axis=1)

        tramos_validos = [t for t in df_obra["tramo"].dropna().unique() if str(t).strip().lower() != "nan" and str(t).strip() != ""]
        
        idx_defecto = 0
        if st.session_state.proyecto_activo in tramos_validos:
            idx_defecto = tramos_validos.index(st.session_state.proyecto_activo)
            
        tramo_sel = st.selectbox("Filtrar Análisis por Frente/Tramo:", tramos_validos, index=idx_defecto)
        st.session_state.proyecto_activo = tramo_sel
        df_inv = df_obra[df_obra["tramo"] == tramo_sel].copy()
        
        total_piquetes_frente = len(df_inv)
        total_aisladores_frente = pd.to_numeric(df_inv["cantidad_aisladores"], errors='coerce').fillna(0).astype(int).sum()
        aisladores_colocados_frente = pd.to_numeric(df_inv[df_inv["montaje_aislador"].notna() & (~df_inv["montaje_aislador"].astype(str).str.upper().isin(["N/A", "NO APLICA"]))]["cantidad_aisladores"], errors='coerce').fillna(0).astype(int).sum()
        
        total_m_frente = df_inv["metros_tendido"].fillna(0).sum()
        m_colocados_frente = df_inv[df_inv["tendido"].notna() & (~df_inv["tendido"].astype(str).str.upper().isin(["N/A", "NO APLICA"]))]["metros_tendido"].fillna(0).sum()

        total_m3_frente = df_inv["m3_excavacion"].fillna(0).sum()
        m3_ejecutados_frente = df_inv[df_inv["excavacion"].notna() & (~df_inv["excavacion"].astype(str).str.upper().isin(["N/A", "NO APLICA"]))]["m3_excavacion"].fillna(0).sum()

        m1, m2, m3, m4, m5 = st.columns(5)
        with m1:
            st.markdown(f"""<div class='kpi-card' style='border-left-color: #3b82f6;'><div class='kpi-title'>📍 Total Piquetes</div><div class='kpi-value'>{total_piquetes_frente}</div></div>""", unsafe_allow_html=True)
        with m2:
            st.markdown(f"""<div class='kpi-card' style='border-left-color: #8b5cf6;'><div class='kpi-title'>🔌 Aisladores Montados</div><div class='kpi-value'>{int(aisladores_colocados_frente)} / {int(total_aisladores_frente)}</div></div>""", unsafe_allow_html=True)
        with m3:
            st.markdown(f"""<div class='kpi-card' style='border-left-color: #06b6d4;'><div class='kpi-title'>📏 Metros Tendidos</div><div class='kpi-value'>{int(m_colocados_frente)} / {int(total_m_frente)} m</div></div>""", unsafe_allow_html=True)
        with m4:
            st.markdown(f"""<div class='kpi-card' style='border-left-color: #eab308;'><div class='kpi-title'>⛏️ Excavación (m³)</div><div class='kpi-value'>{m3_ejecutados_frente:.1f} / {total_m3_frente:.1f} m³</div></div>""", unsafe_allow_html=True)
        with m5:
            st.markdown(f"""<div class='kpi-card' style='border-left-color: #10b981;'><div class='kpi-title'>📈 % Avance Físico</div><div class='kpi-value'>{int(df_inv["Avance_%"].mean() if total_piquetes_frente > 0 else 0)}%</div></div>""", unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### 📊 Auditoría y Estado de Parámetros por Columna")
        
        columnas_analizar = {
            "cabezal": "🧩 Cabezal de Estructura",
            "longitud_poste": "📏 Longitud de Poste",
            "cantidad_aisladores": "🔌 Cantidad de Aisladores por Poste",
            "metros_tendido": "📏 Metros de Vano / Tendido",
            "m3_excavacion": "⛏️ Volúmenes de Excavación (m³)",
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
            if col_real not in ["Avance_%", "cantidad_aisladores", "metros_tendido", "m3_excavacion"]:
                df_filtrado_grafico[col_real] = df_filtrado_grafico[col_real].astype(str)
                
            df_frecuencia = df_filtrado_grafico[col_real].value_counts().reset_index()
            df_frecuencia.columns = [col_sel, "Cantidad de Piquetes"]
            
            if col_real in ["Avance_%", "cantidad_aisladores", "metros_tendido", "m3_excavacion"]:
                df_frecuencia = df_frecuencia.sort_values(by=col_sel)
            
            g_col1, g_col2 = st.columns([3, 2])
            with g_col1:
                fig_inv = px.bar(df_frecuencia, x=col_sel, y="Cantidad de Piquetes", text="Cantidad de Piquetes", 
                                 color="Cantidad de Piquetes", color_continuous_scale="Viridis")
                fig_inv.update_layout(xaxis_type='category' if col_real not in ["Avance_%", "cantidad_aisladores", "metros_tendido", "m3_excavacion"] else 'linear')
                st.plotly_chart(fig_inv, use_container_width=True)
            with g_col2:
                fig_donut = px.pie(df_frecuencia, names=col_sel, values="Cantidad de Piquetes", hole=0.4, 
                                   color_discrete_sequence=px.colors.qualitative.Safe)
                st.plotly_chart(fig_donut, use_container_width=True)
                
        with st.expander("🔍 Ver Listado Detallado de este Frente"):
            st.dataframe(df_inv[["piquete", "tipo_estructura", "cabezal", "longitud_poste", "cantidad_aisladores", "metros_tendido", "m3_excavacion", "tipo_de_equipo", "Avance_%"]], use_container_width=True)

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
        st.session_state.proyecto_activo = tramo_sel
        piquetes_filtrados = df_combos[df_combos["tramo"] == tramo_sel]["piquete"].unique()
        piquete_sel = st.selectbox("Estructura / Piquete Específico:", piquetes_filtrados)
        
        conn = conectar_db()
        p_info = pd.read_sql_query("SELECT * FROM piquetes WHERE piquete = ?", conn, params=[piquete_sel]).iloc[0]
        conn.close()
        
        cabezal_val = p_info["cabezal"] if p_info["cabezal"] and p_info["cabezal"] != "None" else "S/D"
        l_poste = p_info["longitud_poste"] if p_info["longitud_poste"] and p_info["longitud_poste"] != "None" else "N/A"
        cant_aisl_actual = int(p_info["cantidad_aisladores"]) if pd.notna(p_info["cantidad_aisladores"]) else 0
        metros_tendido_actual = float(p_info["metros_tendido"]) if pd.notna(p_info["metros_tendido"]) else 0.0
        m3_excavacion_actual = float(p_info["m3_excavacion"]) if pd.notna(p_info["m3_excavacion"]) else 0.0
        
        st.info(f"🏗️ **ESTRUCTURA:** {p_info['tipo_estructura']} | 🧩 **CABEZAL:** {cabezal_val} | 📏 **LONGITUD POSTE:** {l_poste} | 🔌 **AISLADORES:** {cant_aisl_actual} ud | 📏 **TENDIDO:** {metros_tendido_actual} m | ⛏️ **EXCAVACIÓN:** {m3_excavacion_actual:.2f} m³")
        
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
            doc_red_actual = p_info["idi"] if p_info["idi"] and p_info["idi"] != "None" else None
            if doc_red_actual and os.path.exists(os.path.join(CARPA_ARCHIVOS, doc_red_actual)):
                st.write(f"🗺️ ** IDI Activo:** `{doc_red_actual}`")
                with open(os.path.join(CARPA_ARCHIVOS, doc_red_actual), "rb") as file:
                    st.download_button(label="📥 Descargar IDI", data=file, file_name=doc_red_actual, mime="application/octet-stream", key="dl_redline")
            else:
                st.warning("⚠️ No hay ningún IDI cargado.")
                
        st.markdown("---")
        
        with st.form("form_trazabilidad_avanzado"):
            def convertir_a_fecha(val):
                if val and str(val).strip().upper() not in ["NONE", "NAN", "N/A", "NO APLICA", "NA", "-", ""]:
                    try: 
                        return pd.to_datetime(val).date()
                    except: 
                        return None
                return None

            st.markdown("##### ⚙️ Configuración, Mediciones y Volúmenes del Piquete")
            
            c_aisl1, c_aisl2, c_aisl3, c_aisl4 = st.columns(4)
            with c_aisl1:
                cabezal_input = st.text_input("🧩 Cabezal:", value=cabezal_val)
            with c_aisl2:
                cant_aisladores_input = st.number_input("🔌 Cantidad de Aisladores:", min_value=0, max_value=20, value=cant_aisl_actual)
            with c_aisl3:
                metros_tendido_input = st.number_input("📏 Metros de Tendido / Vano (m):", min_value=0.0, step=1.0, value=metros_tendido_actual, format="%.2f")
            with c_aisl4:
                m3_excavacion_input = st.number_input("⛏️ Volumen de Excavación (m³):", min_value=0.0, step=0.1, value=m3_excavacion_actual, format="%.2f")

            st.markdown("---")
            st.markdown("##### 🛠️ Carga de Fechas para los 9 Hitos de Campo (Marque 'N/A' si el hito no aplica)")
            
            def input_hito_con_na(label, clave_db):
                val_db = str(p_info.get(clave_db, "")).strip()
                es_na_previo = val_db.upper() in ["N/A", "NO APLICA", "NA"]
                c_f, c_na = st.columns([3, 1])
                with c_na:
                    st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                    marcado_na = st.checkbox("N/A", value=es_na_previo, key=f"chk_na_{clave_db}")
                with c_f:
                    if marcado_na:
                        st.text_input(label, value="N/A", disabled=True, key=f"inp_na_{clave_db}")
                        return "N/A"
                    else:
                        fecha_val = st.date_input(label, value=convertir_a_fecha(val_db), key=f"inp_f_{clave_db}")
                        return str(fecha_val) if fecha_val else None

            col1, col2 = st.columns(2)
            with col1:
                f_excav = input_hito_con_na("1. EXCAV PIQUETES", "excavacion")
                f_vert = input_hito_con_na("2. VERTICALIZADO", "verticalizado")
                f_desfile = input_hito_con_na("3. DESFILE DE POSTE", "desfile_de_poste")
                f_riendas = input_hito_con_na("4. MONTAJE RIENDAS", "montaje_riendas")
            with col2:
                f_crucetas = input_hito_con_na("5. ARMADO DE CRUCETAS", "armado_de_crucetas")
                f_aislador = input_hito_con_na("6. MONTAJE DE AISLADOR (FECHA)", "montaje_aislador")
                f_tendido = input_hito_con_na("7. TENDIDO DE CONDUCTOR", "tendido")
                f_flechado = input_hito_con_na("8. FLECHADO", "flechado")
                f_engramp = input_hito_con_na("9. ENGRAMPADO", "engrampado")

            st.markdown("---")
            f_montaje = st.date_input("Fecha Montaje / Liberación Final", value=convertir_a_fecha(p_info["fecha_montaje"]))

            st.markdown("##### 📥 Carga / Actualización de Documentos Técnicos")
            col_arch1, col_arch2 = st.columns(2)
            with col_arch1:
                archivo_anexo = st.file_uploader("Subir Nuevo PLANIALTIMETRIA", type=["docx", "xlsx", "pdf", "xls"])
            with col_arch2:
                archivo_redline = st.file_uploader("Subir Nuevo IDI", type=["docx", "xlsx", "pdf", "xls"])

            if st.form_submit_button("💾 Actualizar Historial de Trazabilidad y Archivos"):
                nombre_anexo = p_info["anexo_montaje"]
                nombre_redline = p_info["idi"]
                
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
                    UPDATE piquetes SET cabezal=?, cantidad_aisladores=?, metros_tendido=?, m3_excavacion=?, excavacion=?, verticalizado=?, desfile_de_poste=?, montaje_riendas=?, armado_de_crucetas=?, montaje_aislador=?, tendido=?, flechado=?, engrampado=?, fecha_montaje=?, anexo_montaje=?, idi=?
                    WHERE piquete=?
                """, (cabezal_input, int(cant_aisladores_input), float(metros_tendido_input), float(m3_excavacion_input), 
                      f_excav, f_vert, f_desfile, f_riendas, f_crucetas, f_aislador, f_tendido, f_flechado, f_engramp, 
                      str(f_montaje) if f_montaje else None, str(nombre_anexo) if nombre_anexo else None, 
                      str(nombre_redline) if nombre_redline else None, piquete_sel))
                conn.commit()
                conn.close()
                
                upload_db_to_drive()
                st.session_state.proyecto_activo = tramo_sel
                st.success(f"✔️ Historial de {piquete_sel} actualizado correctamente.")
                st.rerun()

# -------------------------------------------------------------------------
# MÓDULO 1: ANALÍTICA AVANZADA Y KPIS
# -------------------------------------------------------------------------
else:
    conn = conectar_db()
    df_obra = pd.read_sql_query("SELECT * FROM piquetes", conn)
    df_cronogramas = pd.read_sql_query("SELECT * FROM cronogramas", conn)
    df_metas = pd.read_sql_query("SELECT * FROM metas_ritmo", conn)
    conn.close()
    
    if df_obra.empty:
        st.info("No existen registros de obra para procesar analíticas. Vaya al módulo de Migración.")
    else:
        df_obra["Avance_%"] = df_obra.apply(calcular_avance_piquete, axis=1)

        tramos_validos = [t for t in df_obra["tramo"].dropna().unique() if str(t).strip().lower() != "nan" and str(t).strip() != ""]
        
        idx_defecto = 0
        if st.session_state.proyecto_activo in tramos_validos:
            idx_defecto = tramos_validos.index(st.session_state.proyecto_activo)
            
        tramo_sel = st.selectbox("Frente Operativo / Proyecto Seleccionado:", tramos_validos, index=idx_defecto)
        st.session_state.proyecto_activo = tramo_sel
        df_tramo = df_obra[df_obra["tramo"] == tramo_sel]

        c_actual = df_cronogramas[df_cronogramas["tramo"] == tramo_sel]
        val_ini = pd.to_datetime(c_actual["inicio"].iloc[0]).date() if not c_actual.empty else datetime.date.today()
        val_ent = pd.to_datetime(c_actual["entrega"].iloc[0]).date() if not c_actual.empty else (datetime.date.today() + datetime.timedelta(days=60))

        m_actual = df_metas[df_metas["tramo"] == tramo_sel]
        val_ritmo_guardado = float(m_actual["ritmo_objetivo"].iloc[0]) if not m_actual.empty and pd.notna(m_actual["ritmo_objetivo"].iloc[0]) else None

        total_piquetes_tramo = len(df_tramo)
        piquetes_completados_100 = len(df_tramo[df_tramo["Avance_%"] >= 99.9])
        piquetes_pendientes = max(0, total_piquetes_tramo - piquetes_completados_100)

        inicio_base = pd.to_datetime(val_ini)
        entrega_base = pd.to_datetime(val_ent)
        hoy = pd.to_datetime(datetime.date.today())

        dias_restantes = max(0, (entrega_base - hoy).days)
        dias_transcurridos = max(1, (hoy - inicio_base).days)
        avance_promedio = df_tramo["Avance_%"].mean() if not df_tramo.empty else 0

        if piquetes_pendientes == 0:
            ritmo_requerido_piquetes = 0.0
        elif dias_restantes <= 0:
            ritmo_requerido_piquetes = 999.0
        else:
            ritmo_requerido_piquetes = round(piquetes_pendientes / dias_restantes, 2)

        with st.expander("⚙️ Ajustes Contractuales y Metas de Ritmo / Fecha (Grabación en BD)"):
            col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 2, 1.5])
            with col_f1: 
                f_inicio = st.date_input("Fecha Inicio Contractual", val_ini, key=f"i_{tramo_sel}")
            with col_f2: 
                f_entrega = st.date_input("Fecha Fin Contractual", val_ent, key=f"e_{tramo_sel}")
            with col_f3:
                ritmo_objetivo_input = st.number_input(
                    "Meta Ritmo (piquetes/día):", 
                    value=val_ritmo_guardado if val_ritmo_guardado is not None else float(ritmo_requerido_piquetes), 
                    min_value=0.0, step=0.1, format="%.2f"
                )
            with col_f4:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Guardar Parametrización en BD", type="primary", use_container_width=True):
                    conn = conectar_db()
                    conn.execute("INSERT OR REPLACE INTO cronogramas (tramo, inicio, entrega) VALUES (?, ?, ?)", (tramo_sel, str(f_inicio), str(f_entrega)))
                    conn.execute("INSERT OR REPLACE INTO metas_ritmo (tramo, ritmo_objetivo, fecha_meta) VALUES (?, ?, ?)", (tramo_sel, ritmo_objetivo_input, str(f_entrega)))
                    conn.commit()
                    conn.close()
                    upload_db_to_drive()
                    st.session_state.proyecto_activo = tramo_sel
                    st.success("✔️ Parámetros contractuales y metas de ritmo guardados con éxito.")
                    st.rerun()

        if avance_promedio <= 0.5:
            ritmo_diario = 0.0
            fin_proyectado = entrega_base
            desviacion_dias = 0
        else:
            ritmo_diario = avance_promedio / dias_transcurridos
            dias_proyectados_totales = int(100 / ritmo_diario) if ritmo_diario > 0 else 1825
            dias_proyectados_totales = min(dias_proyectados_totales, 1825)
            fin_proyectado = inicio_base + pd.Timedelta(days=dias_proyectados_totales)
            desviacion_dias = (fin_proyectado - entrega_base).days

        hitos_completados = 0
        for hito in HITOS_OBRA:
            hitos_completados += df_tramo[df_tramo[hito].notna() & (~df_tramo[hito].astype(str).str.upper().isin(["N/A", "NO APLICA"]))].shape[0]

        productividad_media = round(hitos_completados / dias_transcurridos, 2)

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        
        with kpi1:
            st.markdown(f"""
                <div class='kpi-card' style='border-left-color: #3b82f6;'>
                    <div class='kpi-title'>Avance Físico Consolidado</div>
                    <div class='kpi-value'>{int(avance_promedio)}%</div>
                    <div class='kpi-delta' style='color: #60a5fa;'>⚡ Ritmo global: {round(ritmo_diario, 2)}% / día</div>
                </div>
            """, unsafe_allow_html=True)

        with kpi2:
            st.markdown(f"""
                <div class='kpi-card' style='border-left-color: #10b981;'>
                    <div class='kpi-title'>Productividad de Campo</div>
                    <div class='kpi-value'>{productividad_media} <span style='font-size:13px;color:#94a3b8;'>hitos/día</span></div>
                    <div class='kpi-delta' style='color: #34d399;'>Total hitos logrados: {hitos_completados}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with kpi3:
            color_desv = "#ef4444" if desviacion_dias > 0 else "#10b981"
            txt_desv = f"+ {desviacion_dias} días de retraso" if desviacion_dias > 0 else f"{abs(desviacion_dias)} días adelantado"
            st.markdown(f"""
                <div class='kpi-card' style='border-left-color: {color_desv};'>
                    <div class='kpi-title'>Desviación Contractual</div>
                    <div class='kpi-value'>{abs(desviacion_dias)} <span style='font-size:13px;color:#94a3b8;'>días</span></div>
                    <div class='kpi-delta' style='color: {color_desv};'>{txt_desv}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with kpi4:
            st.markdown(f"""
                <div class='kpi-card' style='border-left-color: #f59e0b;'>
                    <div class='kpi-title'>Proyección Cierre Real</div>
                    <div class='kpi-value' style='font-size:22px;'>{fin_proyectado.strftime('%d/%m/%Y')}</div>
                    <div class='kpi-delta' style='color: #94a3b8;'>Contrato: {entrega_base.strftime('%d/%m/%Y')}</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<h3 style='color:#f8fafc; font-size:17px; margin-top:15px; margin-bottom:10px;'>🛠️ Desglose de Avance por Hito Operativo (Los 9 Hitos)</h3>", unsafe_allow_html=True)
        
        datos_hitos = []
        for hito in HITOS_OBRA:
            piquetes_aplicables = df_tramo[~df_tramo[hito].astype(str).str.upper().isin(["N/A", "NO APLICA", "NA", "N/D"])]
            tot_ap = len(piquetes_aplicables)
            cant = piquetes_aplicables[hito].notna().sum()
            pct = (cant / tot_ap * 100) if tot_ap > 0 else 0
            datos_hitos.append({
                "Hito": NOMBRES_HITOS[hito],
                "Piquetes Executados": cant,
                "Total Aplicable": tot_ap,
                "Porcentaje": round(pct, 1)
            })
        
        df_grafico_hitos = pd.DataFrame(datos_hitos)

        col_h1, col_h2 = st.columns([2, 1])
        with col_h1:
            fig_hitos_barras = px.bar(
                df_grafico_hitos, 
                x="Porcentaje", 
                y="Hito", 
                orientation='h',
                text="Porcentaje", 
                color="Porcentaje",
                color_continuous_scale="Blues", 
                template="plotly_dark",
                labels={"Porcentaje": "% de Cumplimiento"}
            )
            fig_hitos_barras.update_traces(texttemplate='%{text}%', textposition='outside')
            fig_hitos_barras.update_layout(
                margin=dict(l=10, r=10, t=10, b=10), height=340, 
                showlegend=False, coloraxis_showscale=False,
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig_hitos_barras, use_container_width=True)

        with col_h2:
            st.markdown("##### 📌 Avance Unitario")
            for h in datos_hitos:
                st.markdown(f"**{h['Hito']}:** `{h['Piquetes Executados']}/{h['Total Aplicable']}` ({h['Porcentaje']}%)")

        st.markdown("<h3 style='color:#f8fafc; font-size:17px; margin-top:15px; margin-bottom:10px;'>🎯 Metas de Ritmo, Flujo de Campo y Estado</h3>", unsafe_allow_html=True)

        if piquetes_pendientes == 0:
            alerta_color = "#10b981"
            txt_meta = "0 piquetes/día"
            txt_sub = "¡Frente 100% Finalizado!"
        elif dias_restantes <= 0:
            alerta_color = "#ef4444"
            txt_meta = f"{piquetes_pendientes} piquetes"
            txt_sub = "⚠️ Plazo contractual vencido."
        else:
            alerta_color = "#ef4444" if ritmo_requerido_piquetes > 2.0 else "#f59e0b"
            txt_meta = f"{ritmo_requerido_piquetes} piq/día"
            txt_sub = f"Faltan {piquetes_pendientes} piquetes ({dias_restantes} días rest.)"

        col_mid1, col_mid2, col_mid3 = st.columns([1.1, 1.3, 1.1])

        with col_mid1:
            st.markdown(f"""
                <div class='status-card' style='border-left: 5px solid {alerta_color};'>
                    <div class='kpi-title' style='color:{alerta_color};'>🎯 Meta Requerida (Contrato)</div>
                    <div class='kpi-value' style='color:{alerta_color}; font-size:24px;'>{txt_meta}</div>
                    <div style='color: #94a3b8; font-size: 13px; margin-top:6px;'>{txt_sub}</div>
                    <hr style='border-color: #1f2937; margin: 12px 0;'>
                    <div style='font-size: 12px; color: #cbd5e1;'>
                        <b>Ritmo Meta BD:</b> {val_ritmo_guardado if val_ritmo_guardado else ritmo_requerido_piquetes} piq/día
                    </div>
                </div>
            """, unsafe_allow_html=True)

        with col_mid2:
            fig_frentes = px.bar(
                df_grafico_hitos.head(4), x="Piquetes Executados", y="Hito", orientation='h',
                text="Piquetes Executados", color="Piquetes Executados", 
                color_continuous_scale="Tealgrn", template="plotly_dark"
            )
            fig_frentes.update_layout(
                margin=dict(l=10, r=10, t=10, b=10), height=145, 
                showlegend=False, coloraxis_showscale=False,
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig_frentes, use_container_width=True)

        with col_mid3:
            st.markdown(f"""
                <div class='status-card' style='border-left: 5px solid #8b5cf6;'>
                    <div class='kpi-title' style='color:#a78bfa;'>📊 Estado Operativo de Frente</div>
                    <div style='color: #f8fafc; font-size: 14px; margin-top:8px;'>
                        <b>Piquetes Totales:</b> {total_piquetes_tramo}<br>
                        <b>Concluidos 100%:</b> {piquetes_completados_100}<br>
                        <b>Pendientes:</b> {piquetes_pendientes}
                    </div>
                    <div style='margin-top:10px; font-size:12px; color:#94a3b8;'>
                        Progreso General del Frente: <b>{int(avance_promedio)}%</b>
                    </div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<h3 style='color:#f8fafc; font-size:17px; margin-top:20px; margin-bottom:10px;'>📈 Trayectoria Proyectada vs Ritmo Real de Campo</h3>", unsafe_allow_html=True)

        df_historico = df_tramo[df_tramo["Avance_%"] >= 99.9].copy()
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
            template="plotly_dark",
            color_discrete_map={
                "Meta Necesaria (Contrato)": "#10b981", 
                f"Tendencia Real ({ritmo_real_diario} piq/día)": "#ef4444" if ritmo_real_diario < ritmo_requerido_piquetes else "#3b82f6"
            }
        )
        
        fig_proy.add_hline(y=total_piquetes_tramo, line_dash="dash", line_color="#64748b", annotation_text=f"Total Objetivo: {total_piquetes_tramo} piq.")
        fig_proy.update_layout(
            margin=dict(l=10, r=10, t=20, b=10), height=300,
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
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

        st.markdown("<h3 style='color:#f8fafc; font-size:17px; margin-top:20px; margin-bottom:10px;'>📊 Simulación de Plazo Contractual vs Proyección de Ritmo Actual</h3>", unsafe_allow_html=True)
        df_gantt = pd.DataFrame([
            {"Línea de Tiempo": "Plazo Comprometido por Contrato", "Inicio": inicio_base, "Fin": entrega_base, "Condición": "Contrato Base"},
            {"Línea de Tiempo": "Proyección por Avance de Campo", "Inicio": inicio_base, "Fin": fin_proyectado, "Condición": "Proyección Real de Obra"}
        ])
        fig = px.timeline(
            df_gantt, x_start="Inicio", x_end="Fin", y="Línea de Tiempo", color="Condición", 
            template="plotly_dark",
            color_discrete_map={"Contrato Base": "#1d4ed8", "Proyección Real de Obra": "#d97706"}
        )
        fig.update_yaxes(autorange="reversed", title="")
        fig.update_layout(
            margin=dict(l=20, r=20, t=10, b=20), height=170,
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("<h3 style='color:#f8fafc; font-size:17px; margin-top:20px;'>📋 Matriz Completa de Trazabilidad</h3>", unsafe_allow_html=True)
        df_mostrar = df_tramo.copy()

        renombrar_columnas_export = {
            "excavacion": "1. EXCAV PIQUETES",
            "verticalizado": "2. VERTICALIZADO",
            "desfile_de_poste": "3. DESFILE DE POSTE",
            "montaje_riendas": "4. MONTAJE RIENDAS",
            "armado_de_crucetas": "5. ARMADO DE CRUCETAS",
            "montaje_aislador": "6. MONTAJE DE AISLADOR (FECHA)",
            "tendido": "7. TENDIDO DE CONDUCTOR",
            "flechado": "8. FLECHADO",
            "engrampado": "9. ENGRAMPADO",
            "fecha_montaje": "10. FECHA MONTAJE / LIBERACIÓN FINAL"
        }

        for hito in HITOS_OBRA:
            df_mostrar[hito] = df_mostrar[hito].apply(lambda x: "N/A" if str(x).upper() in ["N/A", "NO APLICA"] else (pd.to_datetime(x).strftime('%d/%m/%Y') if pd.notna(pd.to_datetime(x, errors='coerce')) else "-"))
            
        df_exportar = df_mostrar.rename(columns=renombrar_columnas_export)
        
        columnas_export_orden = [
            "tramo", 
            "piquete", 
            "tipo_estructura", 
            "cabezal", 
            "longitud_poste", 
            "cantidad_aisladores", 
            "metros_tendido", 
            "m3_excavacion", 
            "Avance_%", 
            "anexo_montaje", 
            "idi"
        ] + list(renombrar_columnas_export.values())
        
        st.markdown("---")
        st.markdown("### 📥 Exportar y Notificar Reportes de Trazabilidad")
        
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
            df_exportar[columnas_export_orden].to_excel(writer, sheet_name=f"Progreso_{tramo_sel}", index=False)
        
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
            components.html(
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
                <button class="btn-print" onclick="window.parent.print()">📄 Guardar Reporte / KPIs (PDF)</button>
                """,
                height=45
            )

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
