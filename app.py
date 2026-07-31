# ============================================================================
# PARTE 1: IMPORTS Y CONFIGURACIÓN INICIAL
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
# PARTE 2: INICIALIZACIÓN DE SESIÓN Y FUNCIONES AUXILIARES
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
        'mostrar_confirmacion_eliminar': False,
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
# PARTE 3: CSS Y ENCABEZADO PRINCIPAL
# ============================================================================

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
st.markdown('<p class="sub-header">Herramienta que ayuda a validar y gestionar los documentos del PLE</p>', unsafe_allow_html=True)

# ============================================================================
# PESTAÑAS PRINCIPALES
# ============================================================================
tab_validar, tab_duplicados_internos, tab_conciliacion, tab_convertidor, tab_instrucciones = st.tabs([
    "🔎 Validación",
    "🔍 Duplicados Internos",
    "🔁 Conciliación",
    "🗂️ Convertir TXT a EXCEL",
    "📖 ¿Cómo usar?"
])

# ============================================================================
# PARTE 4: SIDEBAR - PANEL DE CONTROL PRINCIPAL
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
# PARTE 5: PESTAÑA 1 - VALIDACIÓN
# ============================================================================

with tab_validar:
    st.markdown("## 🔎 Validador de Duplicados")
    st.markdown("Compara el archivo cargado con los últimos 12 meses almacenados.")
    st.markdown("**Resumen:** Valida el mes cargado buscando duplicados exactos frente al historial (últimos 12 meses).\n Permite subir un nuevo mes a la base de datos si no hay duplicados o si se desea reemplazar el mes existente. \n En la **Gestión** puedes cargar múltiples archivos para llenar el historial inicial, eliminar el último mes o eliminar toda la base de datos.")

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
    
    # Mostrar resultados de validación
    if st.session_state.df_nuevo is not None and st.session_state.validacion_realizada:
        st.markdown("---")
        st.markdown("### Resultados de validación")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Mes cargado", st.session_state.mes_nuevo)
        with col2:
            st.metric("Registros", len(st.session_state.df_nuevo))
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
                    st.markdown("#### Resumen de duplicados")
                    st.dataframe(st.session_state.resumen_df, use_container_width=True, hide_index=True)
                
                st.markdown("#### Detalle de duplicados")
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
        
        if st.button("Subir a la base de datos", use_container_width=True, type="primary"):
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
    with st.expander("📤 Cargar múltiples archivos (12 documentos xlsx que seran la base de datos a usar)"):
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
                    
                    st.session_state._batch_key = f"batch_uploader_{int(time.time())}"
                    st.session_state._cargando_batch = False
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
                except Exception as e:
                    st.error(f"❌ Error al eliminar: {str(e)}")
        with col2:
            if st.button("Cancelar", key="btn_cancelar_eliminar_todo", use_container_width=True):
                st.session_state['mostrar_confirmacion_eliminar'] = False
    
    # --- Botón: Ver resumen BD ---
    if st.button("📋 Ver resumen BD", use_container_width=True):
        meses = obtener_meses_existentes()
        if meses:
            st.info(f"📦 Total de meses: {len(meses)} de {MESES_A_MANTENER}")
            st.write("Meses almacenados:", ", ".join(ordenar_meses(meses)))
        else:
            st.info("📭 Base de datos vacía")
# ============================================================================
# PARTE 6: PESTAÑA 2 - DUPLICADOS INTERNOS + PESTAÑA 3 - CONCILIACIÓN
# ============================================================================

# ============================================================================
# PESTAÑA 2: DUPLICADOS INTERNOS
# ============================================================================
with tab_duplicados_internos:
    st.markdown("## Validador de Duplicados Internos")
    st.markdown("Detecta registros duplicados dentro del mismo archivo Excel.")
    st.markdown("**Estructura (.xlsx):** Debe contener columnas: tipo_comprobante, serie, numero, ruc_proveedor (RUC) e importe_total. Si su plantilla usa columnas por letra, H/J/M/Y suelen mapear a Serie/Numero/RUC/Importe. Lectura: primera hoja desde fila 8; hojas adicionales desde fila 2.")
    st.markdown("**Cómo funciona:** Se cargan todas las hojas, se validan las columnas clave y se detectan grupos de registros duplicados. Se genera un reporte Excel descargable y una auditoría resumida para revisión rápida.")
   
    
    # SECCIÓN: CARGAR ARCHIVO
    st.markdown("---")
    st.markdown("### 📂 Cargar archivo")
    st.caption("Selecciona un archivo Excel para analizar duplicados internos.")
    
    archivo_interno = st.file_uploader(
        "Selecciona un archivo Excel (.xlsx)",
        type=["xlsx"],
        key="file_uploader_interno",
        help="El sistema leerá TODAS las hojas del archivo y detectará duplicados basándose en las columnas clave."
    )
    
    if archivo_interno is not None:
        nombre_archivo = archivo_interno.name
        temp_path_interno = f"uploads/{nombre_archivo}_interno"
        
        with open(temp_path_interno, "wb") as f:
            f.write(archivo_interno.getbuffer())
        
        with st.spinner("📖 Leyendo archivo Excel..."):
            try:
                df_raw = leer_excel_todas_hojas(temp_path_interno)
                st.session_state.df_interno_raw = df_raw
                st.session_state.archivo_interno_nombre = nombre_archivo
                st.session_state.df_interno_duplicados = None
                st.session_state.auditoria_interna = None
                st.session_state.reporte_interno_path = None
                
                st.success(f"✅ Archivo cargado: **{nombre_archivo}**")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("📋 Total de hojas", df_raw['hoja_origen'].nunique() if not df_raw.empty else 0)
                with col2:
                    st.metric("📊 Total de registros", len(df_raw))
                with col3:
                    st.metric("📌 Columnas clave", "4 campos (H, J, M, Y)")
                
                if not df_raw.empty:
                    hojas = df_raw['hoja_origen'].unique().tolist()
                    st.caption(f"📄 Hojas procesadas: {', '.join(hojas)}")
                
            except Exception as e:
                st.error(f"❌ Error al leer archivo: {str(e)}")
                st.session_state.df_interno_raw = None
                if os.path.exists(temp_path_interno):
                    try:
                        os.remove(temp_path_interno)
                    except:
                        pass
    
    if st.session_state.df_interno_raw is not None and not st.session_state.df_interno_raw.empty:
        st.markdown("---")
        st.markdown("### 🔍 Analizar duplicados")
        st.caption("Detecta registros duplicados dentro del archivo basándose en las columnas clave.")
        
        if st.button("🔍 DETECTAR DUPLICADOS INTERNOS", use_container_width=True, type="primary"):
            with st.spinner("🔄 Analizando duplicados en todas las hojas..."):
                try:
                    df_dups, auditoria = detectar_duplicados_internos(st.session_state.df_interno_raw)
                    st.session_state.df_interno_duplicados = df_dups
                    st.session_state.auditoria_interna = auditoria
                    
                    if df_dups.empty:
                        st.success("✅ **¡Excelente!** No se encontraron duplicados en el archivo.")
                        st.balloons()
                    else:
                        st.warning(f"⚠️ Se encontraron **{len(df_dups)} registros duplicados** en **{auditoria.get('grupos_duplicados', 0)}** grupos.")
                        
                        if len(df_dups) > 100:
                            st.info(f"📊 Se encontraron {len(df_dups)} duplicados. Descarga el reporte Excel para ver todos los detalles.")
                   
                except Exception as e:
                    st.error(f"❌ Error al detectar duplicados: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
        
        if st.session_state.df_interno_duplicados is not None:
            st.markdown("---")
            st.markdown("### 📊 Resultados del análisis")
            
            if st.session_state.auditoria_interna:
                auditoria = st.session_state.auditoria_interna
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("📋 Total de filas", auditoria.get('total_filas', 0))
                with col2:
                    st.metric("⚠️ Duplicados", auditoria.get('total_duplicados', 0))
                with col3:
                    st.metric("📊 Grupos únicos", auditoria.get('grupos_duplicados', 0))
                with col4:
                    porcentaje = (auditoria.get('total_duplicados', 0) / auditoria.get('total_filas', 1)) * 100
                    st.metric("% Duplicados", f"{porcentaje:.1f}%")
                
                if auditoria.get('duplicados_por_hoja'):
                    st.markdown("#### 📄 Duplicados por hoja")
                    tabla_hojas = pd.DataFrame([
                        {"Hoja": hoja, "Cantidad de duplicados": cantidad}
                        for hoja, cantidad in auditoria['duplicados_por_hoja'].items()
                    ])
                    tabla_hojas = tabla_hojas.sort_values('Cantidad de duplicados', ascending=False)
                    st.dataframe(tabla_hojas, use_container_width=True, hide_index=True)
            
            if not st.session_state.df_interno_duplicados.empty:
                st.markdown("#### 📋 Registros duplicados encontrados")
                
                cols_mostrar = [col for col in st.session_state.df_interno_duplicados.columns
                               if not (isinstance(col, str) and col.startswith("_"))]
                
                df_mostrar = st.session_state.df_interno_duplicados[cols_mostrar].head(20)
                st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
                
                if len(st.session_state.df_interno_duplicados) > 20:
                    st.caption(f"📋 Mostrando 20 de {len(st.session_state.df_interno_duplicados)} duplicados. Descarga el reporte Excel para ver todos.")
                
                st.markdown("---")
                col_btn1, col_btn2 = st.columns([2, 1])
                with col_btn1:
                    if st.button("📥 Generar Reporte Excel", use_container_width=True, type="primary"):
                        with st.spinner("⏳ Generando reporte Excel..."):
                            try:
                                nombre_base = os.path.splitext(st.session_state.archivo_interno_nombre)[0]
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                nombre_salida = f"reportes/duplicados_internos_{nombre_base}.xlsx"
                                
                                generar_reporte_duplicados_interno(
                                    st.session_state.df_interno_raw,
                                    st.session_state.df_interno_duplicados,
                                    st.session_state.auditoria_interna,
                                    nombre_salida,
                                    st.session_state.archivo_interno_nombre
                                )
                                
                                st.session_state.reporte_interno_path = nombre_salida
                                st.success(f"✅ Reporte generado: **{os.path.basename(nombre_salida)}**")
                               
                            except Exception as e:
                                st.error(f"❌ Error generando reporte: {str(e)}")
                                import traceback
                                st.code(traceback.format_exc())
                
                with col_btn2:
                    if st.button("🔄 Limpiar resultados", use_container_width=True):
                        st.session_state.df_interno_duplicados = None
                        st.session_state.auditoria_interna = None
                        st.session_state.reporte_interno_path = None
                        st.rerun()
            
            if st.session_state.reporte_interno_path and os.path.exists(st.session_state.reporte_interno_path):
                st.markdown("---")
                with open(st.session_state.reporte_interno_path, "rb") as f:
                    st.download_button(
                        label="📥 Descargar Reporte Excel Completo",
                        data=f,
                        file_name=os.path.basename(st.session_state.reporte_interno_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
            
            with st.expander("📋 Ver auditoría detallada"):
                if st.session_state.auditoria_interna:
                    texto_auditoria = generar_auditoria_duplicados(
                        st.session_state.df_interno_duplicados,
                        st.session_state.auditoria_interna
                    )
                    st.code(texto_auditoria, language="text")
    
    if st.session_state.get('df_interno_raw') is not None:
        try:
            for archivo in os.listdir("uploads"):
                if archivo.endswith("_interno"):
                    ruta = os.path.join("uploads", archivo)
                    if os.path.exists(ruta):
                        tiempo_mod = os.path.getmtime(ruta)
                        if time.time() - tiempo_mod > 3600:
                            os.remove(ruta)
        except:
            pass

# ============================================================================
# PESTAÑA 3: CONCILIACIÓN
# ============================================================================
with tab_conciliacion:
    st.markdown("## 🔁 Conciliación de PLE Compras")
    st.markdown("Compara dos archivos PLE Compras para identificar diferencias entre SIRE_SUNAT y SIRE_BN.")
    st.markdown("**Estructura de los archivos a subir:** Ambos archivos (.xlsx) deben contener, como mínimo, columnas que identifiquen un comprobante: tipo_comprobante, serie, numero y ruc_proveedor (RUC). Si su plantilla tiene filas de encabezado adicionales, indique la fila de inicio. Se leerán todas las hojas y la comparación se realiza por la clave tipo+serie+numero+RUC.")
    st.markdown("""
        ### 📋 Requisitos del archivo
        
        | Columna | Letra | Campo |
        |---------|-------|-------| 
        | 7 | G | Tipo de Comprobante | 
        | 8 | H | Serie | 
        | 10 | J | Número | 
        | 13 | M | RUC | 
        
        **Importante:** La conciliación se realiza usando la combinación: **Tipo + Serie + Número + RUC**
        
        **Fila de inicio:**
        - `= 1` → Los datos comienzan en la FILA 1
        - `= 2` → Los datos comienzan en la FILA 2 (si hay encabezados)
        - `= 8` → Los datos comienzan en la FILA 8 (como en la hoja 8.1)
    """)
    
    # SECCIÓN: CONFIGURACIÓN DE ARCHIVOS
    st.markdown("---")
    st.markdown("### 📂 Cargar archivos")
    st.caption("Sube los dos archivos Excel y especifica la fila donde comienzan los datos.")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 📄 Archivo 1 (SIRE_SUNAT)")
        archivo1 = st.file_uploader(
            "Selecciona archivo SIRE_SUNAT",
            type=["xlsx"],
            key="conc1",
            help="Base de datos de SIRE_SUNAT (archivo maestro)"
        )
        fila1 = st.number_input(
            "Fila de inicio (Archivo 1)",
            min_value=1,
            value=1,
            step=1,
            key="fila1",
            help="Fila donde comienzan los datos (1-based)"
        )
    
    with col2:
        st.markdown("#### 📄 Archivo 2 (SIRE_BN)")
        archivo2 = st.file_uploader(
            "Selecciona archivo SIRE_BN",
            type=["xlsx"],
            key="conc2",
            help="Archivo de SIRE_BN para comparar"
        )
        fila2 = st.number_input(
            "Fila de inicio (Archivo 2)",
            min_value=1,
            value=1,
            step=1,
            key="fila2",
            help="Fila donde comienzan los datos (1-based)"
        )
    
    # INICIALIZACIÓN DE ESTADO
    if "df_conc1" not in st.session_state:
        st.session_state.df_conc1 = None
        st.session_state.df_conc2 = None
        st.session_state.nombre_conc1 = None
        st.session_state.nombre_conc2 = None
        st.session_state.fila_conc1 = 1
        st.session_state.fila_conc2 = 1
        st.session_state.conciliacion_realizada = False
        st.session_state.resumen_conciliacion = None
    
    if archivo1 and archivo2:
        if (st.session_state.nombre_conc1 != archivo1.name or
            st.session_state.nombre_conc2 != archivo2.name or
            st.session_state.fila_conc1 != fila1 or
            st.session_state.fila_conc2 != fila2):
            st.session_state.df_conc1 = None
            st.session_state.df_conc2 = None
            st.session_state.nombre_conc1 = archivo1.name
            st.session_state.nombre_conc2 = archivo2.name
            st.session_state.fila_conc1 = fila1
            st.session_state.fila_conc2 = fila2
            st.session_state.conciliacion_realizada = False
            st.session_state.resumen_conciliacion = None
    
    # BOTÓN CONCILIACIÓN GENERAL
    st.markdown("---")
    
    if archivo1 is None or archivo2 is None:
        st.info("📭 Sube ambos archivos para realizar la conciliación.")
    else:
        if st.button("📊 Conciliación General", use_container_width=True, type="primary"):
            with st.spinner("🔄 Procesando archivos..."):
                try:
                    path1 = f"uploads/conc_{archivo1.name}"
                    path2 = f"uploads/conc_{archivo2.name}"
                    
                    with open(path1, "wb") as f:
                        f.write(archivo1.getbuffer())
                    with open(path2, "wb") as f:
                        f.write(archivo2.getbuffer())
                    
                    df1 = leer_todas_hojas_conciliacion(path1, fila_inicio=fila1)
                    df2 = leer_todas_hojas_conciliacion(path2, fila_inicio=fila2)
                    
                    st.session_state.df_conc1 = df1
                    st.session_state.df_conc2 = df2
                    st.session_state.nombre_conc1 = archivo1.name
                    st.session_state.nombre_conc2 = archivo2.name
                    
                    resumen, solo1, solo2 = conciliar_archivos(
                        df1, df2,
                        nombre1=archivo1.name,
                        nombre2=archivo2.name
                    )
                    
                    st.session_state.resumen_conciliacion = resumen
                    st.session_state.conciliacion_realizada = True
                    
                    st.success("✅ Conciliación completada con éxito")
                    
                    st.markdown("### 📊 Resumen de Conciliación")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric(
                            "📄 Total registros",
                            f"{resumen['total_registros_1']:,} / {resumen['total_registros_2']:,}",
                            help=f"{archivo1.name}: {resumen['total_registros_1']:,}\n{archivo2.name}: {resumen['total_registros_2']:,}"
                        )
                    with col2:
                        st.metric(
                            "✅ Registros en común",
                            f"{resumen['comunes']:,}",
                            help="Registros que coinciden en ambos archivos"
                        )
                    with col3:
                        diferencia = resumen['diferencias_totales']
                        st.metric(
                            "⚠️ Diferencias totales",
                            f"{diferencia:,}",
                            delta="Revisar" if diferencia > 0 else "Sin diferencias",
                            delta_color="inverse" if diferencia > 0 else "off"
                        )
                    
                    st.markdown("#### Detalle por archivo")
                    df_resumen = pd.DataFrame({
                        "Concepto": [
                            "Total registros (serie)",
                            "IDs únicos",
                            "Registros solo en este archivo",
                            "Registros en común",
                            "Diferencias totales"
                        ],
                        archivo1.name: [
                            resumen['total_registros_1'],
                            resumen['ids_unicos_1'],
                            resumen['solo_en_1'],
                            resumen['comunes'],
                            resumen['diferencias_totales']
                        ],
                        archivo2.name: [
                            resumen['total_registros_2'],
                            resumen['ids_unicos_2'],
                            resumen['solo_en_2'],
                            resumen['comunes'],
                            resumen['diferencias_totales']
                        ]
                    })
                    st.dataframe(df_resumen, use_container_width=True, hide_index=True)
                    
                    nombre_salida = f"reportes/conciliacion_general.xlsx"
                    generar_reporte_conciliacion(
                        resumen, solo1, solo2,
                        df1, df2,
                        archivo1.name, archivo2.name,
                        nombre_salida
                    )
                    
                    with open(nombre_salida, "rb") as f:
                        st.download_button(
                            label="📥 Descargar Reporte de Conciliación General",
                            data=f,
                            file_name=os.path.basename(nombre_salida),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            type="primary"
                        )
                    
                    def eliminar_con_reintentos(ruta, intentos=3, espera=0.5):
                        for i in range(intentos):
                            try:
                                if os.path.exists(ruta):
                                    os.remove(ruta)
                                    return True
                            except PermissionError:
                                if i < intentos - 1:
                                    time.sleep(espera)
                                    continue
                        return False
                    
                    del df1
                    del df2
                    gc.collect()
                    time.sleep(0.3)
                    eliminar_con_reintentos(path1)
                    eliminar_con_reintentos(path2)
                    
                except Exception as e:
                    st.error(f"❌ Error en la conciliación: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
    
    # REPORTES DE PRESENCIA
    st.markdown("---")
    st.markdown("### 📋 Reportes de Presencia")
    st.caption("Genera reportes específicos para identificar registros presentes o ausentes en cada sistema.")
    
    if st.session_state.get('conciliacion_realizada', False) and st.session_state.df_conc1 is not None and st.session_state.df_conc2 is not None:
        
        with st.container():
            st.markdown("#### Reporte SIRE_BN")
            
            col_desc, col_btn = st.columns([3, 1])
            with col_desc:
                st.markdown("""
                Muestra qué registros del **SIRE_BN** están **presentes** o **no presentes** en **SIRE_SUNAT**.
                
                *Útil para identificar comprobantes declarados en SIRE_BN que no están registrados en SIRE_SUNAT.*
                """)
            with col_btn:
                if st.button("📋 SIRE_BN", use_container_width=True):
                    with st.spinner("Generando reporte SIRE_BN..."):
                        try:
                            df1 = st.session_state.df_conc1
                            df2 = st.session_state.df_conc2
                            nombre1 = st.session_state.nombre_conc1
                            nombre2 = st.session_state.nombre_conc2
                            
                            nombre_salida = f"reportes/reporte_SIRE_BN.xlsx"
                            generar_reporte_presentes_no_presentes(
                                df_sire_sunat=df1,
                                df_sire_bn=df2,
                                nombre_sire_sunat=nombre1,
                                nombre_sire_bn=nombre2,
                                ruta_salida=nombre_salida
                            )
                            
                            with open(nombre_salida, "rb") as f:
                                st.download_button(
                                    label="📥 Descargar Reporte SIRE_BN",
                                    data=f,
                                    file_name=os.path.basename(nombre_salida),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    type="primary"
                                )
                            st.success("✅ Reporte SIRE_BN generado correctamente")
                        except Exception as e:
                            st.error(f"❌ Error generando reporte SIRE_BN: {str(e)}")
                            import traceback
                            st.code(traceback.format_exc())
        
        st.markdown("---")
        
        with st.container():
            st.markdown("#### Reporte SIRE_SUNAT")
            
            col_desc, col_btn = st.columns([3, 1])
            with col_desc:
                st.markdown("""
                Muestra qué registros del **SIRE_SUNAT** están **presentes** o **no presentes** en **SIRE_BN**.
                
                *Útil para identificar comprobantes registrados en SIRE_SUNAT que NO han sido declarados en SIRE_BN.*
                """)
            with col_btn:
                if st.button("📋 SIRE_SUNAT", use_container_width=True):
                    with st.spinner("Generando reporte SIRE_SUNAT..."):
                        try:
                            from conciliador import generar_reporte_presentes_no_presentes_sire_sunat
                            
                            df1 = st.session_state.df_conc1
                            df2 = st.session_state.df_conc2
                            nombre1 = st.session_state.nombre_conc1
                            nombre2 = st.session_state.nombre_conc2
                            
                            nombre_salida = f"reportes/reporte_SIRE_SUNAT.xlsx"
                            generar_reporte_presentes_no_presentes_sire_sunat(
                                df_sire_sunat=df1,
                                df_sire_bn=df2,
                                nombre_sire_sunat=nombre1,
                                nombre_sire_bn=nombre2,
                                ruta_salida=nombre_salida
                            )
                            
                            with open(nombre_salida, "rb") as f:
                                st.download_button(
                                    label="📥 Descargar Reporte SIRE_SUNAT",
                                    data=f,
                                    file_name=os.path.basename(nombre_salida),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    type="primary"
                                )
                            st.success("✅ Reporte SIRE_SUNAT generado correctamente")
                        except Exception as e:
                            st.error(f"❌ Error generando reporte SIRE_SUNAT: {str(e)}")
                            import traceback
                            st.code(traceback.format_exc())
    else:
        st.info("ℹ️ Primero ejecuta la **Conciliación General** para habilitar los reportes individuales.")

# ============================================================================
# PARTE 7: PESTAÑA 4 - ¿CÓMO USAR? + PESTAÑA 5 - INFORMACIÓN + PIE DE PÁGINA
# ============================================================================

# ============================================================================
# PESTAÑA 4: ¿CÓMO USAR? - GUÍA COMPLETA
# ============================================================================
with tab_instrucciones:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    
    st.markdown("## 📖 Guía de uso paso a paso")
    st.markdown("Sigue esta guía para utilizar correctamente el Validador PLE Compras.")
    
    with st.expander("🟢 PASO 1: Preparar el archivo", expanded=True):
        st.markdown("""
        ### 📋 Requisitos del archivo
        
        | Requisito | Descripción |
        |-----------|-------------|
        | **Formato** | Excel (.xlsx) |
        | **Nombre** | Debe incluir 6 dígitos MMYYYY (ej: `PLE_COMPRAS_032025.xlsx`) |
        | **Hojas requeridas** | `8.1` (desde fila 8) y `PROGRAMAS SOCIALES` (desde fila 2) |
        | **Estructura** | Debe seguir el formato estándar del PLE Compras |
        
        ### 📌 Pasos a seguir
        
        1. **Verifica el nombre del archivo**: Asegúrate de que el nombre contenga exactamente 6 dígitos con el mes y año (MMYYYY).
           - ✅ Correcto: `PLE_COMPRAS_032025.xlsx`
           - ❌ Incorrecto: `PLE_COMPRAS_2025-03.xlsx`
        
        2. **Confirma las hojas**: Abre el archivo y verifica que contenga las hojas `8.1` y `PROGRAMAS SOCIALES`.
        
        3. **Revisa los datos**: Asegúrate de que los datos comiencen en la fila correcta (fila 8 para la hoja 8.1 y fila 2 para Programas Sociales).
        """)
    
    with st.expander("🟢 PASO 2: Cargar el archivo", expanded=False):
        st.markdown("""
        ### 📂 Cómo subir el archivo
        
        1. Ve al **panel izquierdo** y busca la sección **"📂 Subir nuevo mes"**.
        2. Haz clic en **"Selecciona archivo..."**.
        3. Elige el archivo Excel de tu computadora.
        4. Espera a que se cargue (verás un mensaje de confirmación).
        
        ### ⚙️ Qué sucede automáticamente
        
        - El sistema **lee automáticamente** las hojas requeridas.
        - Te mostrará cuántos registros se cargaron.
        - Si hay error en el nombre o las hojas, te lo indicará con un mensaje de error ❌.
        """)

    with st.expander("🟢 PASO 3: Validar duplicados", expanded=False):
        st.markdown("""
        ### 🔍 Proceso de validación

        1. Una vez subido el archivo, verás el botón **"🔍 VALIDAR DUPLICADOS"** en el panel izquierdo.
        2. Haz clic en el botón para iniciar la validación.
        3. El sistema comparará tu archivo con los **últimos 12 meses** almacenados en la base de datos.

        ### 📊 Campos que se comparan

        | Campo | Descripción |
        |-------|-------------|
        | Periodo (AAAAMM00) | Periodo del comprobante |
        | Código Único de la Operación (CUO) | CUO del comprobante |
        | Fecha de emisión | Fecha del comprobante |
        | Tipo de comprobante | Tipo (Factura, Boleta, etc.) |
        | Serie del comprobante | Serie del comprobante |
        | Número del comprobante | Número del comprobante |
        | RUC del proveedor | RUC del proveedor |
        | Razón social | Nombre del proveedor |
        | Base imponible | Base imponible del comprobante |
        | IGV | Monto del IGV |
        | Importe total | Importe total del comprobante |

        ### 📈 Interpretación de resultados

        - **Sin duplicados** ✅: Puedes agregar el mes sin problema.
        - **Con duplicados** ⚠️: Verás un resumen y detalle para revisar.
        """)

    with st.expander("🟢 PASO 4: Revisar resultados", expanded=False):
        st.markdown("""
        ### 📋 Información mostrada en la pestaña "🔎 Validación"

        **1. Métricas principales**
        - 📅 Mes cargado
        - 📊 Registros totales
        - ⚠️ Estado (Limpio / Con duplicados)

        **2. Resumen de duplicados** (si los hay)
        - Tabla agrupada por mes y tipo de comprobante

        **3. Detalle de duplicados**
        - Lista de los primeros 20 duplicados encontrados

        **4. Descarga de reporte**
        - Botón para descargar el reporte Excel completo

        ### 🎯 Decisiones a tomar

        | Situación | Acción recomendada |
        |-----------|-------------------|
        | **Sin duplicados** | Presiona **"✅ Subir a la base de datos"** |
        | **Con duplicados** | 1. Descarga el Excel<br>2. Revisa y corrige tu archivo<br>3. Sube de nuevo el archivo corregido |
        """)

    with st.expander("🟢 PASO 5: Agregar al historial", expanded=False):
        st.markdown("""
        ### 💾 Confirmar y agregar

        1. Presiona **"✅ Subir a la base de datos"** en el panel izquierdo.
        2. El sistema guardará todos los registros en la base de datos local.

        ### 🔄 Gestión automática del historial

        - **Límite de meses**: Se mantienen los últimos **12 meses**.
        - **Rotación automática**: Si ya hay 12 meses, el más antiguo se elimina automáticamente.
        - **Reemplazo**: Si el mes ya existe, se reemplaza con la nueva versión.

        ### 📦 Cargar múltiples archivos

        Si necesitas llenar el historial inicial:
        1. Ve a **"⚙️ Gestión" → "📤 Cargar múltiples archivos"**
        2. Selecciona los archivos (máximo 12)
        3. Presiona "Cargar batch"
        """)

    with st.expander("🟢 PASO 6: Corregir errores", expanded=False):
        st.markdown("""
        ### 🛠️ Herramientas de gestión

        **Eliminar último mes** 🗑️
        - Si agregaste el mes equivocado, puedes eliminarlo.
        - Solo elimina el mes más reciente.

        **Eliminar toda BD** 🔥
        - Reinicia completamente la base de datos.
        - **Acción irreversible**: Se eliminan TODOS los meses.

        ### ⚠️ Advertencias importantes

        > **¡Cuidado!** Las acciones de gestión son **irreversibles**.
        > Siempre verifica antes de eliminar datos.
        """)

    with st.expander("🟢 PASO 7: Validar duplicados internos", expanded=False):
        st.markdown("""
        ### 🔍 Detectar duplicados dentro del mismo archivo

        La pestaña **"🔍 Duplicados Internos"** permite detectar registros duplicados **dentro de un mismo archivo Excel**.

        **Columnas clave:** H (Serie), J (Número), M (RUC), Y (Importe total)

        ### 📋 Proceso

        1. Sube un archivo Excel en la pestaña "Duplicados Internos"
        2. Haz clic en **"🔍 DETECTAR DUPLICADOS INTERNOS"**
        3. Revisa los resultados y descarga el reporte si es necesario
        """)

    with st.expander("🟢 PASO 8: Conciliación de archivos", expanded=False):
        st.markdown("""
        ### 🔁 Comparar dos archivos PLE Compras

        La pestaña **"🔁 Conciliación"** permite comparar dos archivos Excel.

        **Proceso:**

        1. **Carga los archivos**: SIRE_SUNAT (Base) y SIRE_BN (Comparar)
        2. **Especifica la fila de inicio** para cada archivo
        3. **Ejecuta "Conciliación General"** para obtener el resumen
        4. **Genera reportes específicos**:
           - 📋 Reporte SIRE_BN
           - 📋 Reporte SIRE_SUNAT

        **Campos de comparación:** Tipo + Serie + Número + RUC
        """)

    st.markdown("</div>", unsafe_allow_html=True)

# ============================================================================
# PESTAÑA CONVERTIDOR SIRE (TXT a Excel)
# ============================================================================
with tab_convertidor:
    st.markdown("## 🔄 Convertidor SIRE/PLE - SUNAT")
    st.markdown("Convierte archivos TXT del formato SIRE a Excel con múltiples hojas.")
    
    # ============================================================
    # CSS personalizado para el convertidor
    # ============================================================
    st.markdown("""
    <style>
    .opcion-card {
        background: #f8f9fa;
        border: 2px solid #dee2e6;
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        transition: all 0.3s ease;
        height: 100%;
    }
    .opcion-card:hover {
        border-color: #2563EB;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.15);
    }
    .opcion-title {
        font-size: 1.1rem;
        font-weight: 600;
        color: #0F172A;
        margin-bottom: 8px;
    }
    .opcion-desc {
        font-size: 0.9rem;
        color: #475569;
    }
    .info-box {
        background: #f0f4ff;
        border-left: 4px solid #2563EB;
        padding: 15px;
        border-radius: 8px;
        margin: 10px 0;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # ============================================================
    # IMPORTAR MÓDULOS DEL CONVERTIDOR
    # ============================================================
    try:
        from src.sire_core import (
            SIREValidator,
            TXTProcessorConEncabezado,
            TXTProcessorSinEncabezado,
            ExcelGenerator
        )
        MODULOS_DISPONIBLES = True
    except ImportError as e:
        st.error(f"❌ Error al cargar los módulos del convertidor: {str(e)}")
        st.info("Asegúrate de que la carpeta 'src/' existe y contiene el archivo 'sire_core.py'")
        MODULOS_DISPONIBLES = False
    
    # ============================================================
    # INICIALIZAR ESTADO DEL CONVERTIDOR
    # ============================================================
    if 'formato_seleccionado' not in st.session_state:
        st.session_state.formato_seleccionado = None
    if 'mostrar_paso2' not in st.session_state:
        st.session_state.mostrar_paso2 = False
    if 'archivo_convertir' not in st.session_state:
        st.session_state.archivo_convertir = None
    
    # ============================================================
    # PASO 1: SELECCIÓN DE FORMATO
    # ============================================================
    if MODULOS_DISPONIBLES:
        st.markdown("### 📂 Paso 1: Selecciona el Formato de tu Archivo TXT")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            <div class="opcion-card">
                <div class="opcion-title">📄 Formato CON Encabezados</div>
                <div class="opcion-desc">Primera línea contiene nombres de columnas</div>
                <div class="opcion-desc" style="margin-top:10px;color:#666;">Ej: Ruc|Razón Social|Total CP|...</div>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button("✅ CON ENCABEZADOS", use_container_width=True, key="btn_convertir_formato1"):
                st.session_state.formato_seleccionado = "CON_ENCABEZADOS"
                st.session_state.mostrar_paso2 = True
                st.session_state.archivo_convertir = None
        
        with col2:
            st.markdown("""
            <div class="opcion-card">
                <div class="opcion-title">📄 Formato SIN Encabezados</div>
                <div class="opcion-desc">Archivo inicia directamente con datos</div>
                <div class="opcion-desc" style="margin-top:10px;color:#666;">Ej: 1|20260500|651-18264653-14|...</div>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button("❌ SIN ENCABEZADOS", use_container_width=True, key="btn_convertir_formato2"):
                st.session_state.formato_seleccionado = "SIN_ENCABEZADOS"
                st.session_state.mostrar_paso2 = True
                st.session_state.archivo_convertir = None
        
        # ============================================================
        # PASO 2: CARGA DE ARCHIVO
        # ============================================================
        if st.session_state.mostrar_paso2 and st.session_state.formato_seleccionado:
            formato = st.session_state.formato_seleccionado
            
            st.markdown("---")
            st.markdown(f"### 📂 Paso 2: Cargar archivo ({formato})")
            
            if st.button("🔄 Cambiar formato", use_container_width=False):
                st.session_state.formato_seleccionado = None
                st.session_state.mostrar_paso2 = False
                st.session_state.archivo_convertir = None
                st.rerun()
            
            col1, col2 = st.columns([3, 1])
            with col1:
                uploaded_file = st.file_uploader(
                    f"Selecciona el archivo TXT ({formato})",
                    type=["txt"],
                    help=f"Archivo en formato SIRE - {formato}",
                    key="file_uploader_convertir"
                )
            
            # ============================================================
            # PASO 3: PROCESAMIENTO
            # ============================================================
            if uploaded_file is not None:
                st.session_state.archivo_convertir = uploaded_file
                
                # Info del archivo
                tamano_mb = uploaded_file.size / (1024 * 1024)
                st.markdown(f"""
                <div class="info-box">
                    <b>📁 Archivo:</b> {uploaded_file.name}<br>
                    <b>📊 Tamaño:</b> {tamano_mb:.2f} MB ({uploaded_file.size:,} bytes)<br>
                    <b>📋 Formato:</b> {formato}
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("---")
                st.markdown("### ⚙️ Paso 3: Procesar y Convertir")
                
                if st.button("🚀 CONVERTIR A EXCEL", use_container_width=True, type="primary"):
                    start_time = time.perf_counter()
                    progress_bar = st.progress(0.0)
                    status_text = st.empty()
                    
                    def update_progress(value, message=''):
                        try:
                            progress_bar.progress(value)
                        except Exception:
                            pass
                        if message:
                            status_text.info(message)
                    
                    with st.spinner('⏳ Procesando datos, por favor espere...'):
                        try:
                            # Crear carpetas si no existen
                            os.makedirs("input_files", exist_ok=True)
                            os.makedirs("output_files", exist_ok=True)
                            
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            update_progress(0.05, 'Guardando archivo temporal...')
                            
                            # Guardar temporalmente
                            temp_path = os.path.join("input_files", "temp_upload.txt")
                            with open(temp_path, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                            
                            # Elegir procesador según formato
                            if formato == "CON_ENCABEZADOS":
                                processor = TXTProcessorConEncabezado(max_size_mb=2000)
                            else:
                                processor = TXTProcessorSinEncabezado(max_size_mb=2000)
                            
                            update_progress(0.20, 'Leyendo el archivo TXT...')
                            df, error = processor.read_txt(temp_path)
                            
                            if error:
                                st.error(f"❌ Error al leer: {error}")
                            else:
                                update_progress(0.45, 'Validando datos...')
                                validator = SIREValidator()
                                validation = validator.validate_all(df)
                                
                                # Mostrar advertencias
                                if validation['warnings']:
                                    for warn in validation['warnings']:
                                        st.warning(f"⚠️ {warn}")
                                
                                if not validation['is_valid']:
                                    st.error("❌ Validación fallida:")
                                    for err in validation['errors']:
                                        st.error(f"  • {err}")
                                else:
                                    # Generar Excel
                                    nombre_original = uploaded_file.name
                                    if nombre_original.lower().endswith('.txt'):
                                        nombre_original = nombre_original[:-4]
                                    
                                    output_filename = f"output_files/{nombre_original}.xlsx"
                                    generator = ExcelGenerator()
                                    success, message = generator.create_excel(
                                        df,
                                        output_filename,
                                        progress_callback=update_progress
                                    )
                                    
                                    elapsed = time.perf_counter() - start_time
                                    update_progress(1.0, f'Conversión completada en {elapsed:.2f} segundos')
                                    
                                    if success:
                                        # Limpiar temporal
                                        if os.path.exists(temp_path):
                                            os.remove(temp_path)
                                        
                                        # Mostrar estadísticas
                                        st.markdown("---")
                                        st.markdown("### 📊 Estadísticas de Conversión")
                                        
                                        col1, col2, col3, col4 = st.columns(4)
                                        with col1:
                                            st.metric("📝 Registros", f"{len(df):,}")
                                        with col2:
                                            st.metric("📊 Columnas", len(df.columns))
                                        with col3:
                                            st.metric("📄 Hojas Excel", generator.sheets_created)
                                        with col4:
                                            st.metric("⏱️ Tiempo", f"{elapsed:.2f} s")
                                        
                                        # Buscar columna de monto
                                        monto_col = None
                                        for col in ['MontoTotal', 'Total CP', 'Total']:
                                            if col in df.columns:
                                                monto_col = col
                                                break
                                        if monto_col:
                                            try:
                                                total = pd.to_numeric(df[monto_col], errors='coerce').sum()
                                                st.metric("💰 Total General", f"{total:,.2f}")
                                            except:
                                                st.metric("💰 Total", "N/A")
                                        
                                        # Vista previa
                                        st.markdown("---")
                                        st.markdown("### 👀 Vista Previa de Datos")
                                        st.dataframe(df.head(10), use_container_width=True, height=300)
                                        
                                        st.success("✅ ¡Conversión completada exitosamente!")
                                        
                                        # Botón descarga
                                        with open(output_filename, "rb") as f:
                                            st.download_button(
                                                label="📥 DESCARGAR EXCEL GENERADO",
                                                data=f,
                                                file_name=f"docExcel_{timestamp}.xlsx",
                                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                                use_container_width=True,
                                                type="primary"
                                            )
                                    else:
                                        st.error(f"❌ {message}")
                            
                        except Exception as e:
                            st.error(f"❌ Error inesperado: {str(e)}")
                            with st.expander("Ver detalles técnicos"):
                                st.code(traceback.format_exc())
        
        else:
            if not st.session_state.mostrar_paso2:
                st.info("👆 Selecciona un formato arriba para continuar con la carga del archivo.")
    
    # ============================================================
    # FOOTER DEL CONVERTIDOR
    # ============================================================
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #9e9e9e; padding: 10px;'>
        <p>🏦 <b>Área de Tributación - BN</b></p>
        <p style='font-size: 11px;'>Convertidor SIRE/PLE v2.1 | Automatización de Declaraciones Tributarias</p>
    </div>
    """, unsafe_allow_html=True)
    
    # ============================================================
    # SOPORTE Y CONTACTO
    # ============================================================
    st.markdown("### Soporte y contacto")
    st.markdown("""
    - **Soporte técnico:** Yessly Poma de la Cruz
    - **Área:** Tributación
    - **Versión:** 2.2
    """)

    st.markdown("</div>", unsafe_allow_html=True)


    # ============================================================================
    # PIE DE PÁGINA
    # ============================================================================
    st.markdown("---")
    st.caption("Seguridad: Datos locales en SQLite. Sin conexiones externas. | v2.2 | Soporte: Yessly Poma de la Cruz - Área de Tributación")