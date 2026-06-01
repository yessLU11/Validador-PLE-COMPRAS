# ============================================================================
# app.py - APLICACIÓN PRINCIPAL VALIDADOR PLE COMPRAS
# ============================================================================
# Descripción: Aplicación Streamlit para validación de duplicados en archivos
#              PLE de compras. Permite cargar nuevos meses, compararlos contra
#              el histórico (últimos 12 meses) y detectar registros duplicados.
# ============================================================================

import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from excel_reader import leer_hoja_unificada
from database import (
    init_db, obtener_meses_existentes, agregar_mes,
    cargar_historico_completo, eliminar_ultimo_mes,
    eliminar_toda_base_datos, cargar_multiples_archivos
)
from validator import detectar_duplicados
from report_generator import generar_reporte_excel
from config import MESES_A_MANTENER, COLUMNAS_CLAVE

# ============================================================================
# CONFIGURACIÓN DE LA PÁGINA
# ============================================================================
# Configura la apariencia y comportamiento básico de la aplicación Streamlit
st.set_page_config(
    page_title="# Validador PLE Compras",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# ESTILOS CSS PERSONALIZADOS
# ============================================================================
# Define la apariencia visual de la aplicación: colores, espaciados, 
# efectos hover, etc. Se aplica mediante inyección HTML/CSS.
st.markdown("""
<style>
    /* Variables de color - Paleta profesional */
    :root {
        --primary: #0F172A;              /* Azul oscuro (fondos)     */
        --primary-light: #1E293B;        /* Azul oscuro claro        */
        --accent: #2563EB;               /* Azul principal (acciones)*/
        --accent-light: #60A5FA;         /* Azul claro               */
        --success: #10B981;              /* Verde (éxito)            */
        --warning: #F59E0B;              /* Ámbar (advertencia)      */
        --danger: #EF4444;               /* Rojo (error)             */
        --bg-light: #F8FAFC;             /* Fondo claro              */
        --border: #E2E8F0;               /* Borde gris               */
        --text: #334155;                 /* Texto gris               */
    }
    
    /* Sidebar - Panel izquierdo con degradado */
    [data-testid="stSidebar"] {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
    }
    
    /* Headers principales */
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #F1F5F9;
        margin-bottom: 0.5rem;
        text-align: center;
    }
    
    /* Subtítulos */
    .sub-header {
        font-size: 0.95rem;
        color: #64748B;
        margin-top: 0;
        margin-bottom: 1.5rem;
        text-align: center;
    }
    
    /* Títulos en sidebar */
    .sidebar-title {
        color: #F1F5F9;
        font-size: 1.1rem;
        font-weight: 600;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
        border-bottom: 2px solid #2563EB;
        padding-bottom: 0.5rem;
    }
    
    /* Subtítulos en sidebar */
    .sidebar-subtitle {
        color: #94A3B8;
        font-size: 0.85rem;
        margin-bottom: 0.5rem;
    }
    
    /* Tarjetas (cards) - Contenedores visuales */
    .card {
        background-color: #FFFFFF;
        border-radius: 0.875rem;
        padding: 1.5rem;
        border: 1px solid #E2E8F0;
        margin-bottom: 1.5rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
        transition: box-shadow 0.2s ease;
    }
    
    /* Efecto hover en cards */
    .card:hover {
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.12);
    }
    
    /* Encabezado de tarjeta */
    .card-header {
        font-size: 1.15rem;
        font-weight: 600;
        color: #0F172A;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* Cajas de estadísticas - Colores llamativos */
    .stat-box {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 0.875rem;
        text-align: center;
        margin-bottom: 1rem;
    }
    
    /* Números en estadísticas */
    .stat-number {
        font-size: 2.2rem;
        font-weight: 700;
        margin-bottom: 0.25rem;
    }
    
    /* Etiqueta en estadísticas */
    .stat-label {
        font-size: 0.85rem;
        opacity: 0.9;
    }
    
    /* Elemento de lista en sidebar */
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
    
    /* Divisores/separadores */
    .divider {
        margin: 1.5rem 0;
        border-top: 2px solid #E2E8F0;
    }
    
    /* Tabla de previsualización */
    .preview-table {
        font-size: 0.9rem;
    }
    
    /* Estilos para botones */
    .stButton button {
        border-radius: 0.5rem;
        font-weight: 500;
        transition: all 0.2s ease;
    }
    
    /* Tabs/pestañas */
    .stTabs [data-baseweb="tab-list"] {
        gap: 1rem;
        border-bottom: 2px solid #E2E8F0;
    }
    
    /* Alertas y mensajes */
    .stAlert {
        border-radius: 0.75rem;
        border-left: 4px solid;
    }
    
    /* Expandibles */
    .stExpander {
        border: 1px solid #E2E8F0;
        border-radius: 0.5rem;
        margin-bottom: 0.75rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# INICIALIZACIÓN DE LA BASE DE DATOS Y CARPETAS
# ============================================================================
# Crea la estructura inicial: BD SQLite y carpetas para reportes/uploads
init_db()
os.makedirs("reportes", exist_ok=True)
os.makedirs("uploads", exist_ok=True)

# ============================================================================
# INICIALIZACIÓN DEL ESTADO DE SESIÓN DE STREAMLIT
# ============================================================================
# Streamlit ejecuta el script completo cada vez que el usuario interactúa.
# El estado de sesión persiste variables entre ejecuciones.
# Definimos todas las variables que necesitaremos a lo largo de la app.
if "df_nuevo" not in st.session_state:
    # DataFrame del mes actual cargado
    st.session_state.df_nuevo = None
    # Identificador del mes en formato YYYYMM (ej: 202503)
    st.session_state.mes_nuevo = None
    # Ruta temporal del archivo subido
    st.session_state.ruta_archivo = None
    # DataFrame con duplicados encontrados
    st.session_state.duplicados = None
    # Ruta del archivo Excel de reporte generado
    st.session_state.reporte_path = None
    # DataFrame con resumen de duplicados
    st.session_state.resumen_df = None

# ============================================================================
# ENCABEZADO PRINCIPAL
# ============================================================================
# Muestra el título de la aplicación 
# principal. Se utiliza HTML personalizado para aplicar estilos CSS.
st.markdown('<p class="main-header">🔍 Validador </p>', unsafe_allow_html=True)

st.markdown('<p class="sub-header">Detecta duplicados comparando con los últimos 12 meses</p>', unsafe_allow_html=True)

# ============================================================================
# PESTAÑAS (TABS) PRINCIPALES
# ============================================================================
# Divide la interfaz en 3 secciones principales:
# 1. Validación - Flujo principal de carga y validación
# 2. Instrucciones - Guía de uso paso a paso
# 3. Información - Detalles técnicos y métricas
tab_validar, tab_instrucciones, tab_info = st.tabs(["🔎 Validación", "📖 ¿Cómo usar?", "ℹ️ Información"])

# ============================================================================
# PESTAÑA 1: VALIDACIÓN
# ============================================================================
# Esta es la pestaña principal donde se muestran los resultados de la
# validación y las acciones post-validación. Se muestra un panel con el
# mes cargado, registros, estado y opción de agregar al historial.
with tab_validar:
    if st.session_state.df_nuevo is None:
        # Si no hay datos cargados, mostrar instrucción inicial
        st.info("👈 Usa el panel izquierdo para subir un archivo y validar duplicados")
    else:
        # Panel de información: Mes, registros y estado de validación
        st.markdown('<div class="card">', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("#### 📅 Mes cargado")
            st.markdown(f'<p class="stat-number">{st.session_state.mes_nuevo}</p>', unsafe_allow_html=True)
        
        with col2:
            st.markdown("#### 📊 Registros")
            st.markdown(f'<p class="stat-number">{len(st.session_state.df_nuevo)}</p>', unsafe_allow_html=True)
        
        with col3:
            st.markdown("#### ⚠️ Estado")
            if st.session_state.duplicados is None:
                # Aún no se ha validado
                st.markdown('<p style="color: #F59E0B; font-size: 1.5rem; font-weight: 700;">Pendiente</p>', unsafe_allow_html=True)
            elif st.session_state.duplicados.empty:
                # Validado y sin duplicados
                st.markdown('<p style="color: #10B981; font-size: 1.5rem; font-weight: 700;">✅ Limpio</p>', unsafe_allow_html=True)
            else:
                # Validado y CON duplicados
                st.markdown(f'<p style="color: #EF4444; font-size: 1.5rem; font-weight: 700;">{len(st.session_state.duplicados)}</p>', unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # ====== Sección de resultados (solo si se ha validado) ======
        if st.session_state.duplicados is not None:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            
            if st.session_state.duplicados.empty:
                # SIN duplicados - se puede agregar sin problemas
                st.success("🎉 **Sin duplicados** – Este mes está listo para agregar al historial")
            else:
                # CON duplicados - mostrar resumen y detalle
                # 1. Mostrar resumen (agrupado por mes y tipo comprobante)
                if st.session_state.resumen_df is not None and not st.session_state.resumen_df.empty:
                    st.markdown("### 📈 Resumen de duplicados")
                    st.dataframe(st.session_state.resumen_df, use_container_width=True, hide_index=True)
                
                # 2. Mostrar detalle (primeros 20 registros)
                st.markdown("### 📋 Detalle de duplicados")
                st.dataframe(st.session_state.duplicados.head(20), use_container_width=True)
                if len(st.session_state.duplicados) > 20:
                    st.caption(f"Mostrando 20 de {len(st.session_state.duplicados)} duplicados. Descarga el Excel para ver todos.")
                
                # 3. Botón de descarga del reporte completo
                if st.session_state.reporte_path and os.path.exists(st.session_state.reporte_path):
                    st.markdown("---")
                    with open(st.session_state.reporte_path, "rb") as f:
                        st.download_button(
                            label="📥 Descargar Reporte Excel Completo",
                            data=f,
                            file_name=os.path.basename(st.session_state.reporte_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            type="primary"
                        )
            
            st.markdown("</div>", unsafe_allow_html=True)
            
            # ====== Botón para confirmar y agregar al historial ======
            # Este sirve para agregar el mes al historial incluso si tiene duplicados, asumiendo que el usuario ya revisó el reporte y decidió proceder.
            if st.session_state.duplicados is not None:
                st.markdown('<div class="card">', unsafe_allow_html=True)
                if st.button("✅ CONFIRMAR Y AGREGAR AL HISTORIAL", use_container_width=True, type="primary"):
                    try:
                        # Intentar agregar el mes a la BD
                        exito, mes_eliminado = agregar_mes(
                            st.session_state.df_nuevo,
                            st.session_state.mes_nuevo,
                            os.path.basename(st.session_state.ruta_archivo)
                        )
                        
                        # Informar si se eliminó un mes antiguo (rolling de 12 meses)
                        if mes_eliminado:
                            st.info(f"🗑️ Se eliminó el mes {mes_eliminado} (el más antiguo)")
                        
                        # Confirmar éxito
                        st.success(f"✅ Mes {st.session_state.mes_nuevo} agregado correctamente")
                        
                        # Limpiar estado para siguiente validación
                        st.session_state.df_nuevo = None
                        st.session_state.mes_nuevo = None
                        st.session_state.ruta_archivo = None
                        st.session_state.duplicados = None
                        st.session_state.reporte_path = None
                        st.session_state.resumen_df = None
                        
                        # Reexecutar la app para actualizar interfaz
                        st.rerun()
                    except Exception as e:
                        # Capturar errores en la inserción a BD
                        st.error(f"❌ Error al agregar mes: {str(e)}")
                
                st.markdown("</div>", unsafe_allow_html=True)


# ============================================================================
# PESTAÑA 2: GUÍA DE INSTRUCCIONES
# ============================================================================
# Proporciona un manual completo del uso de la aplicación con pasos
# detallados, requisitos y advertencias. Utiliza expandibles para
# mantener la interfaz limpia y organizada.
with tab_instrucciones:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    
    st.markdown("## 📖 Guía de uso paso a paso")
    
    # Paso 1: Preparación del archivo
    with st.expander("🟢 PASO 1: Preparar el archivo", expanded=True):
        st.markdown("""
        ### Requisitos del archivo
        - **Formato**: Excel (.xlsx)
        - **Nombre**: Debe incluir 6 dígitos MMYYYY (ej: `PLE_COMPRAS_032025.xlsx`)
        - **Contenido**: Debe tener las hojas:
          - `8.1` (desde fila 8)
          - `PROGRAMAS SOCIALES` (desde fila 2)
        
        ### Pasos
        1. Asegúrate de tener el archivo Excel correcto
        2. El nombre debe seguir el formato: `PLE_COMPRAS_MMYYYY.xlsx`
        3. Verifica que contenga los datos en las hojas especificadas
        """)
    
    # Paso 2: Subir el archivo
    with st.expander("🟢 PASO 2: Subir el archivo", expanded=False):
        st.markdown("""
        ### Cómo subir
        1. Mira el panel izquierdo bajo **"📂 Subir nuevo mes"**
        2. Haz clic en **"Selecciona archivo..."**
        3. Elige el archivo Excel de tu computadora
        4. Espera a que se cargue (verás un mensaje de confirmación)
        
        ### Qué sucede
        - El sistema **lee automáticamente** las hojas requeridas
        - Te mostrará cuántos registros se cargaron
        - Si hay error en el nombre o hojas, te lo indicará con ❌
        """)
    
    # Paso 3: Validar duplicados
    with st.expander("🟢 PASO 3: Validar duplicados", expanded=False):
        st.markdown("""
        ### Validación automática
        1. Una vez subido el archivo, verás el botón **"🔍 VALIDAR DUPLICADOS"** en el panel izquierdo
        2. Haz clic en él
        3. El sistema comparará tu archivo con los últimos 12 meses almacenados
        
        ### Qué se compara
        - Se buscan coincidencias exactas en estos campos:
          - Periodo (AAAAMM00)
          - Código Único de la Operación (CUO)  
          - Fecha de emisión
          - Tipo de comprobante
          - Serie del comprobante
          - Número del comprobante
          - RUC del proveedor
          - Razón social
          - Base imponible
          - IGV
          - Importe total
        
        ### Resultados
        - **Sin duplicados**: Puedes agregar el mes sin problema
        - **Con duplicados**: Verás un resumen y detalle para revisar
        """)
    
    # Paso 4: Revisar resultados
    with st.expander("🟢 PASO 4: Revisar resultados", expanded=False):
        st.markdown("""
        ### En la pestaña "🔎 Validación"
        1. **Resumen**: Tabla con duplicados agrupados por mes y tipo de comprobante
        2. **Detalle**: Lista de los primeros 20 duplicados encontrados
        3. **Excel**: Botón para descargar el reporte completo con TODOS los duplicados
        
        ### Decisiones
        - **Sin duplicados**: Presiona **"✅ CONFIRMAR Y AGREGAR AL HISTORIAL"**
        - **Con duplicados**: 
          - Descarga el Excel para ver todos
          - Ajusta tu archivo original (elimina registros duplicados)
          - Sube de nuevo el archivo corregido
        """)
    
    # Paso 5: Agregar al historial
    with st.expander("🟢 PASO 5: Agregar al historial", expanded=False):
        st.markdown("""
        ### Confirmar y agregar
        1. Una vez revisado el resultado, presiona **"✅ CONFIRMAR Y AGREGAR AL HISTORIAL"**
        2. El sistema guardará todos los registros en la base de datos
        
        ### Gestión automática
        - Si ya hay 12 meses: El más antiguo se elimina automáticamente
        - Se mantiene un historial de últimos 12 meses
        - Verás confirmación al terminar
        
        ### Cargar múltiples archivos
        - Si necesitas llenar el historial inicial (11 meses):
          - Abre **"⚙️ Gestión" → "📂 Cargar múltiples archivos"** en el panel izquierdo
          - Sube los 11 archivos a la vez
          - Presiona "Cargar"
        """)
    
    # Paso 6: Corregir errores
    with st.expander("🟢 PASO 6: Corregir errores", expanded=False):
        st.markdown("""
        ### Si agregaste el mes equivocado
        1. Abre **"⚙️ Gestión"** en el panel izquierdo
        2. Presiona **"🗑️ Eliminar último mes"**
        3. El último mes agregado se eliminará
        4. Puedes subir uno nuevo
        
        ### Reiniciar completamente
        1. En **"⚙️ Gestión"**, presiona **"🔥 Eliminar toda BD"**
        2. Confirma cuando se te pida
        3. Toda la base de datos se vaciará
        4. Puedes empezar desde cero
        
        ⚠️ **ADVERTENCIA**: Estas acciones son irreversibles
        """)
    
    st.markdown("</div>", unsafe_allow_html=True)

# ======================= TAB: INFORMACIÓN =======================
with tab_info:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📊 Resultados generados")
        st.markdown("""
        Cuando validas duplicados, se generan automáticamente:
        
        #### 📊 Reporte Excel con 2 hojas:
        
        **1. Duplicados_Detalle**
        - Todas las filas duplicadas encontradas
        - Información completa de cada comprobante
        
        **2. Resumen_por_Mes**
        - Conteo de duplicados agrupado por:
          - Mes donde ya existía el comprobante
          - Tipo de comprobante
          - Cantidad de duplicados en esa combinación
        """)
    
    with col2:
        st.markdown("### 🔐 Seguridad y almacenamiento")
        st.markdown("""
        #### ✅ Características de seguridad:
        - **Base de datos local**: SQLite guardada en tu computadora
        - **Sin conexión externa**: Los datos NO se envían a internet
        - **Privacidad total**: Todo queda en tu máquina
        
        #### 📁 Ubicación de archivos:
        - `data/ple_history.db` - Base de datos principal
        - `reportes/` - Archivos Excel generados
        - `uploads/` - Archivos temporales cargados
        """)
    
    st.markdown("---")
    
    st.markdown("### 📈 Flujo de trabajo recomendado")
    st.markdown("""
    ```
    1. INICIO
       ↓
    2. Subir archivo del nuevo mes
       ↓
    3. Validar duplicados
       ↓
    4. ¿Hay duplicados?
       ├─ NO → Agregar al historial ✅
       └─ SÍ → Descargar Excel, corregir, volver a paso 2
       ↓
    5. Mes agregado al historial
       ↓
    6. Repetir para siguientes meses
    ```
    """)
    
    st.markdown("---")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📦 Máx. meses en historial", f"{MESES_A_MANTENER} meses")
    with col2:
        st.metric("🔄 Meses actuales", len(obtener_meses_existentes()))
    with col3:
        st.metric("💾 Base de datos", "SQLite Local")
    
    st.markdown("</div>", unsafe_allow_html=True)






# ============================================================================
# PESTAÑA 3: INFORMACIÓN Y MÉTRICAS
# ============================================================================
# Proporciona información sobre seguridad, almacenamiento de datos,
# flujo de trabajo recomendado y métricas actuales del sistema.
with tab_info:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    
    # Sección 1: Resultados generados
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📊 Resultados generados")
        st.markdown("""
        Cuando validas duplicados, se generan automáticamente:
        
        #### 📊 Reporte Excel con 2 hojas:
        
        **1. Duplicados_Detalle**
        - Todas las filas duplicadas encontradas
        - Información completa de cada comprobante
        
        **2. Resumen_por_Mes**
        - Conteo de duplicados agrupado por:
          - Mes donde ya existía el comprobante
          - Tipo de comprobante
          - Cantidad de duplicados en esa combinación
        """)
    
    with col2:
        st.markdown("### 🔐 Seguridad y almacenamiento")
        st.markdown("""
        #### ✅ Características de seguridad:
        - **Base de datos local**: SQLite guardada en tu computadora
        - **Sin conexión externa**: Los datos NO se envían a internet
        - **Privacidad total**: Todo queda en tu máquina
        
        #### 📁 Ubicación de archivos:
        - `data/ple_history.db` - Base de datos principal
        - `reportes/` - Archivos Excel generados
        - `uploads/` - Archivos temporales cargados
        """)
    
    st.markdown("---")
    
    # Flujo de trabajo recomendado
    st.markdown("### 📈 Flujo de trabajo recomendado")
    st.markdown("""
    ```
    1. INICIO
       ↓
    2. Subir archivo del nuevo mes
       ↓
    3. Validar duplicados
       ↓
    4. ¿Hay duplicados?
       ├─ NO → Agregar al historial ✅
       └─ SÍ → Descargar Excel, corregir, volver a paso 2
       ↓
    5. Mes agregado al historial
       ↓
    6. Repetir para siguientes meses
    ```
    """)
    
    st.markdown("---")
    
    # Métricas actuales del sistema
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📦 Máx. meses en historial", f"{MESES_A_MANTENER} meses")
    with col2:
        st.metric("🔄 Meses actuales", len(obtener_meses_existentes()))
    with col3:
        st.metric("💾 Base de datos", "SQLite Local")
    
    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================================
# SIDEBAR: PANEL DE CONTROL PRINCIPAL
# ============================================================================
# El sidebar contiene toda la funcionalidad de carga, validación y gestión.
# Este es el panel principal donde el usuario interactúa con la aplicación.
with st.sidebar:
    st.markdown("# 📊 PLE COMPRAS")
    st.markdown("### Validador de Duplicados")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    
    # =====================================================================
    # SECCIÓN 1: Documentos en BD (estado actual del historial)
    # =====================================================================
    st.markdown('<p class="sidebar-title">📦 Documentos en BD</p>', unsafe_allow_html=True)
    meses = obtener_meses_existentes()
    
    if meses:
        # Ordenar meses de forma descendente (más reciente primero)
        meses_ordenados = sorted(meses, reverse=True)
        
        # Mostrar estadística: cuántos meses de los 12 máximos tenemos
        st.markdown(
            f'<div class="stat-box"><div class="stat-number">{len(meses)}</div><div class="stat-label">de {MESES_A_MANTENER} meses</div></div>',
            unsafe_allow_html=True
        )
        
        # Listar los últimos 5 meses
        st.markdown(f'<p class="sidebar-subtitle">Últimos meses registrados:</p>', unsafe_allow_html=True)
        for mes in meses_ordenados[:5]:
            st.markdown(f'<div class="doc-list-item">📅 {mes}</div>', unsafe_allow_html=True)
        
        # Si hay más de 5, mostrar expandible con el resto
        if len(meses_ordenados) > 5:
            with st.expander(f"Ver {len(meses_ordenados) - 5} meses más"):
                for mes in meses_ordenados[5:]:
                    st.markdown(f'<div class="doc-list-item">📅 {mes}</div>', unsafe_allow_html=True)
        
        # Indicador de historial completo
        if len(meses) == MESES_A_MANTENER:
            st.success("✅ Historial completo")
    else:
        # Base de datos vacía
        st.info("📭 Sin documentos aún")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    
    # Sección: Upload nuevo mes
    st.markdown('<p class="sidebar-title">📂 Subir nuevo mes</p>', unsafe_allow_html=True)
    archivo = st.file_uploader("Selecciona archivo PLE_COMPRAS_MMYYYY.xlsx", type=["xlsx"], key="file_uploader", help="Formato: MMYYYY en el nombre")
    
    if archivo is not None:
        nombre = archivo.name
        match = re.search(r"(\d{6})", nombre)
        if not match:
            st.error("❌ El nombre debe incluir 6 dígitos MMYYYY (ej: 032025)")
        else:
            mes = match.group(1)
            temp_path = f"uploads/{nombre}"
            with open(temp_path, "wb") as f:
                f.write(archivo.getbuffer())
            with st.spinner("Leyendo archivo..."):
                df = leer_hoja_unificada(temp_path, mes)
            if df.empty:
                st.error("❌ No hay datos en hojas 8.1 / PROGRAMAS SOCIALES")
            else:
                st.session_state.df_nuevo = df
                st.session_state.mes_nuevo = mes
                st.session_state.ruta_archivo = temp_path
                st.success(f"✅ {mes} ({len(df)} registros)")
                st.session_state.duplicados = None
                st.session_state.reporte_path = None
                st.session_state.resumen_df = None
    
    if st.session_state.df_nuevo is not None:
        if st.button("🔍 VALIDAR DUPLICADOS", use_container_width=True, type="primary"):
            with st.spinner("Comparando con histórico..."):
                df_historico = cargar_historico_completo()
                df_historico = df_historico[df_historico["mes_archivo"] != st.session_state.mes_nuevo]
                dups = detectar_duplicados(st.session_state.df_nuevo, df_historico)
                st.session_state.duplicados = dups
                if dups.empty:
                    st.success("✅ Sin duplicados")
                    st.session_state.resumen_df = pd.DataFrame({"Mensaje": ["Sin duplicados"]})
                else:
                    st.warning(f"⚠️ {len(dups)} duplicados encontrados")
                    resumen_col = "Mes(es)_donde_hay_duplicados" if "Mes(es)_donde_hay_duplicados" in dups.columns else "Mes_donde_ya_existia"
                    if resumen_col in dups.columns and "tipo_comprobante" in dups.columns:
                        resumen = dups.groupby([resumen_col, "tipo_comprobante"]).size().reset_index(name="cantidad_duplicados")
                        st.session_state.resumen_df = resumen
                    else:
                        st.session_state.resumen_df = pd.DataFrame()
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    reporte_nombre = f"reportes/reporte_duplicados_{st.session_state.mes_nuevo}_{timestamp}.xlsx"
                    st.session_state.reporte_path = generar_reporte_excel(dups, reporte_nombre)
            st.rerun()
    #Agregar un ✅ boton Descargar Reporte
    if st.session_state.reporte_path:
        with open(st.session_state.reporte_path, "rb") as f:
            st.download_button("✅ Descargar Reporte", f, file_name=os.path.basename(st.session_state.reporte_path), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    
    # Sección: Gestión
    st.markdown('<p class="sidebar-title">⚙️ Gestión</p>', unsafe_allow_html=True)
    
    if st.button("🗑️ Eliminar último mes", use_container_width=True, help="Quita el mes más reciente agregado"):
        ultimo = eliminar_ultimo_mes()
        if ultimo:
            st.success(f"Mes {ultimo} eliminado")
        else:
            st.info("Sin meses para eliminar")
        st.rerun()
    
    with st.expander("📂 Cargar múltiples archivos"):
        st.caption("Carga 11 archivos para llenar el historial inicial")
        archivos_batch = st.file_uploader("Selecciona archivos", type=["xlsx"], accept_multiple_files=True, key="batch_uploader")
        if archivos_batch and st.button("Cargar"):
            temp_paths = []
            for f in archivos_batch:
                temp = f"uploads/batch_{f.name}"
                with open(temp, "wb") as fp:
                    fp.write(f.getbuffer())
                temp_paths.append((temp, f.name))
            with st.spinner("Procesando..."):
                exito, mensaje, agregados, eliminados = cargar_multiples_archivos(temp_paths)
            if exito:
                st.success(mensaje)
                st.info(f"Agregados: {', '.join(agregados)}")
                if eliminados:
                    st.warning(f"Eliminados: {', '.join(eliminados)}")
            else:
                st.error(mensaje)
            for p, _ in temp_paths:
                if os.path.exists(p):
                    os.remove(p)
            st.rerun()
    
    if st.button("🔥 Eliminar toda BD", use_container_width=True):
        st.warning("⚠️ IRREVERSIBLE - ¿Confirmas?")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Sí, eliminar"):
                eliminar_toda_base_datos()
                st.success("Base de datos eliminada")
                st.session_state.df_nuevo = None
                st.session_state.duplicados = None
                st.session_state.reporte_path = None
                st.session_state.resumen_df = None
                st.rerun()
        with col2:
            st.button("Cancelar")

# Pie de página
st.markdown("---")
st.caption("🔐 **Seguridad**: Datos locales en SQLite. Sin conexiones externas. | v2.1 - UI Mejorada | soporte técnico: Yessly Poma de la cruz Área de Tributación")