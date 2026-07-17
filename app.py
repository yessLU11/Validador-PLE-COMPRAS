# app.py - VERSIÓN DEFINITIVA SIN BUCLES Y CON ELIMINACIÓN FUNCIONAL
# ============================================================================
# APLICACIÓN PRINCIPAL VALIDADOR PLE COMPRAS
# ============================================================================

import streamlit as st
import pandas as pd
import os
import re
import time
import gc
from datetime import datetime

# ============================================================================
# IMPORTS DE MÓDULOS
# ============================================================================
from config import MESES_A_MANTENER
from database import (
    init_db, obtener_meses_existentes, agregar_mes,
    cargar_historico_completo, eliminar_ultimo_mes,
    eliminar_toda_base_datos, cargar_multiples_archivos,
    ordenar_meses, mes_existe, registrar_duplicados_reportados
)
from excel_reader import leer_hoja_unificada
from validator import detectar_duplicados
from report_generator import generar_reporte_excel
from duplicate_detector_internal import (
    leer_excel_todas_hojas, detectar_duplicados_internos,
    generar_auditoria_duplicados
)
from duplicate_report_generator_internal import generar_reporte_duplicados_interno
from conciliador import (
    generar_reporte_presentes_no_presentes,
    generar_reporte_presentes_no_presentes_sire_sunat,
    leer_todas_hojas_conciliacion,
    conciliar_archivos,
    generar_reporte_conciliacion
)

# ============================================================================
# CONFIGURACIÓN DE LA PÁGINA
# ============================================================================
st.set_page_config(
    page_title="Validador PLE Compras",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# INICIALIZACIÓN DE SESIÓN
# ============================================================================
def init_session_state():
    defaults = {
        'df_nuevo': None,
        'mes_nuevo': None,
        'ruta_archivo': None,
        'nombre_archivo': None,
        'duplicados': None,
        'resumen_df': None,
        'reporte_path': None,
        'validacion_realizada': False,
        'df_interno_raw': None,
        'df_interno_duplicados': None,
        'auditoria_interna': None,
        'reporte_interno_path': None,
        'archivo_interno_nombre': None,
        'df_conc1': None,
        'df_conc2': None,
        'nombre_conc1': None,
        'nombre_conc2': None,
        'fila_conc1': 1,
        'fila_conc2': 1,
        'mostrar_resumen_bd': False,
        # Bandera para evitar ejecuciones múltiples
        'mostrar_confirmacion_eliminar': False,  # <--- UNIFICADO
        '_batch_key': 'batch_uploader_default', 
        '_cargando_batch': False,
        '_eliminando_mes': False,
        '_subiendo_mes': False,
        '_eliminando_todo': False,
    }
    for key, default_value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_value

init_session_state()

# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================
def resetear_validacion():
    st.session_state.duplicados = None
    st.session_state.resumen_df = None
    st.session_state.reporte_path = None
    st.session_state.validacion_realizada = False

def resetear_mes_cargado():
    st.session_state.df_nuevo = None
    st.session_state.mes_nuevo = None
    st.session_state.ruta_archivo = None
    st.session_state.nombre_archivo = None
    resetear_validacion()

def descartar_archivo_cargado():
    """Descarta el archivo cargado y limpia el estado."""
    ruta = st.session_state.ruta_archivo
    resetear_mes_cargado()
    if ruta and os.path.exists(ruta):
        try:
            os.remove(ruta)
        except PermissionError:
            pass

def limpiar_uploads():
    """Limpia archivos temporales en la carpeta uploads."""
    contador = 0
    for archivo in os.listdir("uploads"):
        ruta = os.path.join("uploads", archivo)
        try:
            os.remove(ruta)
            contador += 1
        except PermissionError:
            pass
    return contador

# ============================================================================
# CSS PERSONALIZADO
# ============================================================================
st.markdown("""
<style>
    :root {
        --primary: #0F172A;
        --primary-light: #1E293B;
        --accent: #2563EB;
        --accent-light: #60A5FA;
        --success: #10B981;
        --warning: #F59E0B;
        --danger: #EF4444;
        --bg-light: #F8FAFC;
        --border: #E2E8F0;
        --text: #334155;
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
    }
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #000000;
        margin-bottom: 0.5rem;
        text-align: center;
    }
    .sub-header {
        font-size: 0.95rem;
        color: #64748B;
        margin-top: 0;
        margin-bottom: 1.5rem;
        text-align: center;
    }
    .sidebar-title {
        color: #F1F5F9;
        font-size: 1.1rem;
        font-weight: 600;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
        border-bottom: 2px solid #2563EB;
        padding-bottom: 0.5rem;
    }
    .sidebar-subtitle {
        color: #94A3B8;
        font-size: 0.85rem;
        margin-bottom: 0.5rem;
    }
    .card {
        background-color: #FFFFFF;
        border-radius: 0.875rem;
        padding: 1.5rem;
        border: 1px solid #E2E8F0;
        margin-bottom: 1.5rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
        transition: box-shadow 0.2s ease;
    }
    .card:hover {
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12);
    }
    .stat-box {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 0.875rem;
        text-align: center;
        margin-bottom: 1rem;
    }
    .stat-number {
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0.25rem;
    }
    .stat-label {
        font-size: 0.85rem;
        opacity: 0.9;
    }
    .doc-list-item {
        background: #F1F5F9;
        padding: 0.75rem 1rem;
        margin-bottom: 0.5rem;
        border-radius: 0.5rem;
        font-size: 0.9rem;
        font-weight: 500;
        color: #1E293B;
        border-left: 3px solid #2563EB;
    }
    .divider {
        margin: 1.5rem 0;
        border-top: 2px solid #E2E8F0;
    }
    .stButton button {
        border-radius: 0.5rem;
        font-weight: 500;
        transition: all 0.2s ease;
    }
    .stAlert {
        border-radius: 0.75rem;
        border-left: 4px solid;
    }
    .stExpander {
        border: 1px solid #E2E8F0;
        border-radius: 0.5rem;
        margin-bottom: 0.75rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# INICIALIZACIÓN DE BASE DE DATOS
# ============================================================================
init_db()
os.makedirs("reportes", exist_ok=True)
os.makedirs("uploads", exist_ok=True)

# ============================================================================
# ENCABEZADO PRINCIPAL
# ============================================================================
st.markdown('<p class="main-header">🔍 Validador PLE Compras</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Detecta duplicados comparando con los últimos 12 meses</p>', unsafe_allow_html=True)

# ============================================================================
# PESTAÑAS PRINCIPALES
# ============================================================================
tab_validar, tab_duplicados_internos, tab_conciliacion, tab_instrucciones, tab_info = st.tabs([
    "🔎 Validación",
    "🔍 Duplicados Internos",
    "🔁 Conciliación",
    "📖 ¿Cómo usar?",
    "ℹ️ Información"
])

# ============================================================================
# SIDEBAR - PANEL DE CONTROL PRINCIPAL
# ============================================================================
# ============================================================================
# SIDEBAR - PANEL DE CONTROL PRINCIPAL (SOLO DOCUMENTOS Y AYUDA)
# ============================================================================
with st.sidebar:
    st.markdown('# <span style="color:white">📊 PLE COMPRAS</span>', unsafe_allow_html=True)
    st.markdown('### <span style="color:white">Validador de Duplicados</span>', unsafe_allow_html=True)
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    
    # ================================================================
    # SECCIÓN 1: DOCUMENTOS EN BD (SIEMPRE VISIBLE)
    # ================================================================
    st.markdown('<p class="sidebar-title">📦 Documentos en BD</p>', unsafe_allow_html=True)
    meses = obtener_meses_existentes()
    
    if meses:
        meses_ordenados = ordenar_meses(meses, reverse=True)
        st.markdown(
            f'<div class="stat-box"><div class="stat-number">{len(meses)}</div>'
            f'<div class="stat-label">de {MESES_A_MANTENER} meses</div></div>',
            unsafe_allow_html=True
        )
        
        if len(meses) == MESES_A_MANTENER:
            st.success("✅ Historial completo: se mantienen los últimos 12 meses")
        else:
            st.warning(f"⚠️ Historial incompleto: {len(meses)} de {MESES_A_MANTENER} meses")
        
        st.markdown(f'<p class="sidebar-subtitle">Últimos meses registrados:</p>', unsafe_allow_html=True)
        for mes in meses_ordenados[:5]:
            mes_formateado = f"{mes[:2]}/{mes[2:]}"
            st.markdown(f'<div class="doc-list-item">📅 {mes_formateado}</div>', unsafe_allow_html=True)
        
        if len(meses_ordenados) > 5:
            with st.expander(f"Ver {len(meses_ordenados) - 5} meses más"):
                for mes in meses_ordenados[5:]:
                    mes_formateado = f"{mes[:2]}/{mes[2:]}"
                    st.markdown(f'<div class="doc-list-item">📅 {mes_formateado}</div>', unsafe_allow_html=True)
    else:
        st.info("📭 Sin documentos aún")
    
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    
    # ================================================================
    # SECCIÓN 2: AYUDA
    # ================================================================
    st.markdown('<p class="sidebar-title">ℹ️ Ayuda</p>', unsafe_allow_html=True)
    if st.button("📖 Ver instrucciones completas", use_container_width=True):
        st.info("Ve a la pestaña '📖 ¿Cómo usar?' para ver la guía detallada.")

# ============================================================================
# PESTAÑA 1: VALIDACIÓN (CON TODOS LOS CONTROLES - SIN BUCLES)
# ============================================================================
with tab_validar:
    # Título dentro de la pestaña
    st.markdown("## 🔎 Validador de Duplicados")
    st.markdown("Compara el archivo cargado con los últimos 12 meses almacenados.")
    
    # ============================================================
    # SECCIÓN: EVALUAR NUEVO MES (SOLO LECTURA)
    # ============================================================
    st.markdown("---")
    st.markdown("### 📂 Evaluar nuevo mes")
    st.caption("Carga un archivo para validar contra la BD (no se guarda automáticamente).")
    
    archivo = st.file_uploader(
        "Selecciona archivo PLE_COMPRAS_MMYYYY.xlsx",
        type=["xlsx"],
        key="evaluar_uploader",
        help="Formato: MMYYYY en el nombre (ej: 032025)"
    )
    
    if archivo is not None:
        nombre = archivo.name
        match = re.search(r"(\d{6})", nombre)
        if not match:
            st.error("❌ El nombre debe incluir 6 dígitos MMYYYY (ej: 032025)")
        else:
            mes = match.group(1)
            temp_path = f"uploads/evaluar_{nombre}"
            
            with open(temp_path, "wb") as f:
                f.write(archivo.getbuffer())
            
            with st.spinner("Leyendo archivo..."):
                df = leer_hoja_unificada(temp_path, mes)
            
            if df.empty:
                st.error("❌ No hay datos en las hojas del archivo")
            else:
                st.session_state.df_nuevo = df
                st.session_state.mes_nuevo = mes
                st.session_state.ruta_archivo = temp_path
                st.session_state.nombre_archivo = nombre
                resetear_validacion()
                st.success(f"✅ {mes} cargado para evaluación ({len(df)} registros)")
                
                if mes_existe(mes):
                    st.warning(f"⚠️ El mes {mes} ya existe en la BD. Si lo subes, reemplazará la versión anterior.")
    
    # Botones de acción
    if st.session_state.df_nuevo is not None:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col1:
            if st.button("🔍 Validar duplicados", use_container_width=True, type="primary"):
                with st.spinner("Comparando con histórico..."):
                    df_historico = cargar_historico_completo()
                    df_historico = df_historico[df_historico["mes_archivo"] != st.session_state.mes_nuevo]
                    
                    dups = detectar_duplicados(
                        st.session_state.df_nuevo,
                        df_historico,
                        st.session_state.mes_nuevo
                    )
                    
                    st.session_state.duplicados = dups
                    st.session_state.validacion_realizada = True
                    
                    if dups.empty:
                        st.success("✅ Sin duplicados encontrados")
                        st.session_state.resumen_df = pd.DataFrame({"Mensaje": ["Sin duplicados"]})
                    else:
                        if 'ya_reportado' in dups.columns:
                            nuevos = len(dups[dups['ya_reportado'] == False])
                            reportados = len(dups[dups['ya_reportado'] == True])
                            st.warning(f"⚠️ {len(dups)} duplicados ({nuevos} nuevos, {reportados} ya reportados)")
                        else:
                            st.warning(f"⚠️ {len(dups)} duplicados")
                        
                        resumen_col = "Mes(es)_donde_hay_duplicados" if "Mes(es)_donde_hay_duplicados" in dups.columns else "Mes_donde_ya_existia"
                        if resumen_col in dups.columns and "tipo_comprobante" in dups.columns:
                            resumen = dups.groupby([resumen_col, "tipo_comprobante"]).size().reset_index(name="cantidad_duplicados")
                            st.session_state.resumen_df = resumen
                        else:
                            st.session_state.resumen_df = pd.DataFrame()
                        
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        reporte_nombre = f"reportes/reporte_duplicados_{st.session_state.mes_nuevo}_{timestamp}.xlsx"
                        st.session_state.reporte_path = generar_reporte_excel(dups, reporte_nombre)
                        registrar_duplicados_reportados(dups, st.session_state.mes_nuevo)
        
        with col2:
            if st.session_state.reporte_path and os.path.exists(st.session_state.reporte_path):
                with open(st.session_state.reporte_path, "rb") as f:
                    st.download_button(
                        "📥 Descargar reporte",
                        f,
                        file_name=os.path.basename(st.session_state.reporte_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
        with col3:
            if st.button("❌ Descartar archivo", use_container_width=True):
                descartar_archivo_cargado()
                st.success("Archivo descartado. Puedes cargar otro.")
                # ✅ SIN st.rerun() - Streamlit actualiza automáticamente
    
    # Mostrar resultados de validación
    if st.session_state.df_nuevo is not None and st.session_state.validacion_realizada:
        st.markdown("---")
        st.markdown("### 📊 Resultados de validación")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("📅 Mes cargado", st.session_state.mes_nuevo)
        with col2:
            st.metric("📊 Registros", len(st.session_state.df_nuevo))
        with col3:
            if st.session_state.duplicados is not None and not st.session_state.duplicados.empty:
                st.metric("⚠️ Duplicados", len(st.session_state.duplicados), delta="⚠️ Revisar")
            else:
                st.metric("✅ Estado", "Limpio")
        
        if st.session_state.duplicados is not None:
            if st.session_state.duplicados.empty:
                st.success("🎉 **Sin duplicados** – Este mes está limpio.")
                st.info("Si quieres guardarlo en la BD, usa la sección **'📂 Subir nuevo mes'**.")
            else:
                if 'ya_reportado' in st.session_state.duplicados.columns:
                    nuevos = len(st.session_state.duplicados[st.session_state.duplicados['ya_reportado'] == False])
                    reportados = len(st.session_state.duplicados[st.session_state.duplicados['ya_reportado'] == True])
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("🆕 Nuevos duplicados", nuevos)
                    with col2:
                        st.metric("📌 Ya reportados", reportados)
                
                if st.session_state.resumen_df is not None and not st.session_state.resumen_df.empty:
                    st.markdown("#### 📈 Resumen de duplicados")
                    st.dataframe(st.session_state.resumen_df, use_container_width=True, hide_index=True)
                
                st.markdown("#### 📋 Detalle de duplicados")
                st.dataframe(st.session_state.duplicados.head(20), use_container_width=True)
                if len(st.session_state.duplicados) > 20:
                    st.caption(f"Mostrando 20 de {len(st.session_state.duplicados)} duplicados.")
                
                st.warning("⚠️ **Este mes tiene duplicados.** Revisa el reporte y corrige el archivo antes de subirlo.")
    
    # ============================================================
    # SECCIÓN: SUBIR NUEVO MES (ESCRITURA EN BD)
    # ============================================================
    st.markdown("---")
    st.markdown("### 📂 Subir nuevo mes")
    st.caption("Guarda el archivo validado en la base de datos (reemplaza si ya existe).")
    
    if st.session_state.df_nuevo is not None:
        st.info(f"📄 {st.session_state.nombre_archivo} listo para subir ({st.session_state.mes_nuevo})")
        
        meses_actuales = obtener_meses_existentes()
        if len(meses_actuales) >= MESES_A_MANTENER:
            meses_ordenados = ordenar_meses(meses_actuales)
            mes_a_eliminar = meses_ordenados[0]
            st.warning(f"⚠️ Hay 12 meses. Se eliminará el mes más antiguo: **{mes_a_eliminar}**")
        
        if st.button("✅ Subir a la base de datos", use_container_width=True, type="primary"):
            if not st.session_state.get("_subiendo_mes", False):
                st.session_state._subiendo_mes = True
                try:
                    exito, mes_eliminado = agregar_mes(
                        st.session_state.df_nuevo,
                        st.session_state.mes_nuevo,
                        os.path.basename(st.session_state.ruta_archivo)
                    )
                    if mes_eliminado:
                        st.success(f"✅ Mes {st.session_state.mes_nuevo} subido. Se eliminó {mes_eliminado} (el más antiguo)")
                    else:
                        st.success(f"✅ Mes {st.session_state.mes_nuevo} subido correctamente")
                    
                    descartar_archivo_cargado()
                    # ✅ SIN st.rerun() - Streamlit actualiza automáticamente
                except Exception as e:
                    st.error(f"❌ Error al subir mes: {str(e)}")
                finally:
                    st.session_state._subiendo_mes = False
            else:
                st.info("Subida en proceso, espera...")
    else:
        st.info("📭 Carga un archivo en 'Evaluar nuevo mes' primero.")
    
    # ============================================================
    # SECCIÓN: GESTIÓN
    # ============================================================
    st.markdown("---")
    st.markdown("### ⚙️ Gestión")
    
    # --- Cargar múltiples archivos ---
    with st.expander("📤 Cargar múltiples archivos"):
        st.caption("Carga varios archivos para llenar el historial inicial (máx 12).")
        
        batch_key = st.session_state.get('_batch_key', 'batch_uploader_default')
        archivos_batch = st.file_uploader(
            "Selecciona archivos",
            type=["xlsx"],
            accept_multiple_files=True,
            key=batch_key
        )
        
        if st.button("Cargar batch", key="btn_cargar_batch"):
            if not st.session_state.get("_cargando_batch", False):
                st.session_state._cargando_batch = True
                
                if not archivos_batch:
                    st.warning("⚠️ Por favor selecciona al menos un archivo.")
                    st.session_state._cargando_batch = False
                else:
                    temp_paths = []
                    for f in archivos_batch:
                        temp = f"uploads/batch_{f.name}"
                        with open(temp, "wb") as fp:
                            fp.write(f.getbuffer())
                        temp_paths.append((temp, f.name))
                    
                    with st.spinner("Procesando..."):
                        exito, mensaje, agregados, eliminados = cargar_multiples_archivos(temp_paths)
                    
                    gc.collect()
                    time.sleep(0.3)
                    for p, _ in temp_paths:
                        if os.path.exists(p):
                            try:
                                os.remove(p)
                            except PermissionError:
                                pass
                    
                    if exito:
                        st.success(mensaje)
                        if agregados:
                            st.info(f"Agregados: {', '.join(agregados)}")
                        if eliminados:
                            st.warning(f"Eliminados: {', '.join(eliminados)}")
                    else:
                        st.error(mensaje)
                    
                    # Cambiar la clave para limpiar el file_uploader
                    st.session_state._batch_key = f"batch_uploader_{int(time.time())}"
                    st.session_state._cargando_batch = False
                    # ✅ SIN st.rerun() - Streamlit actualiza automáticamente
            else:
                st.warning("Carga en proceso, por favor espera...")
    
    # --- Eliminar último mes ---
    if st.button("🗑️ Eliminar último mes", use_container_width=True, help="Elimina el mes más reciente de la BD"):
        if not st.session_state.get("_eliminando_mes", False):
            st.session_state._eliminando_mes = True
            ultimo = eliminar_ultimo_mes()
            if ultimo:
                st.success(f"Mes {ultimo} eliminado")
                if st.session_state.mes_nuevo == ultimo:
                    descartar_archivo_cargado()
            else:
                st.info("Sin meses para eliminar")
            st.session_state._eliminando_mes = False
            # ✅ SIN st.rerun() - Streamlit actualiza automáticamente
        else:
            st.info("Ya se está eliminando un mes, espera...")
    
    # --- Eliminar toda BD (con confirmación) ---
    if st.button("🔥 Eliminar toda BD", use_container_width=True, key="btn_eliminar_toda_bd"):
        st.session_state['mostrar_confirmacion_eliminar'] = True
    
    if st.session_state.get('mostrar_confirmacion_eliminar', False):
        st.warning("⚠️ IRREVERSIBLE - ¿Confirmas?")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Sí, eliminar", key="btn_confirmar_eliminar_todo", use_container_width=True):
                st.session_state['mostrar_confirmacion_eliminar'] = False
                try:
                    eliminar_toda_base_datos()
                    resetear_mes_cargado()
                    st.session_state.duplicados = None
                    st.session_state.resumen_df = None
                    st.session_state.reporte_path = None
                    st.success("✅ Base de datos eliminada correctamente")
                    # ✅ SIN st.rerun() - Streamlit actualiza automáticamente
                except Exception as e:
                    st.error(f"❌ Error al eliminar: {str(e)}")
        with col2:
            if st.button("Cancelar", key="btn_cancelar_eliminar_todo", use_container_width=True):
                st.session_state['mostrar_confirmacion_eliminar'] = False
                # ✅ SIN st.rerun() - Streamlit actualiza automáticamente
    
    # --- Botón: Ver resumen BD ---
    if st.button("📋 Ver resumen BD", use_container_width=True):
        meses = obtener_meses_existentes()
        if meses:
            st.info(f"📦 Total de meses: {len(meses)} de {MESES_A_MANTENER}")
            st.write("Meses almacenados:", ", ".join(ordenar_meses(meses)))
        else:
            st.info("📭 Base de datos vacía")
