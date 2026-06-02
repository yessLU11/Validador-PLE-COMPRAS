# 📊 Validador PLE Compras v2.2

## 📌 Descripción General

Sistema completo de validación de archivos PLE de compras con **dos modos principales**:

1. **🔍 Duplicados Internos** (v2.2): Detecta duplicados dentro de un mismo archivo Excel
2. **📊 Duplicados Históricos**: Compara un nuevo mes contra los últimos 12 meses en la base de datos

Este proyecto es una herramienta profesional para validar archivos PLE de compras en Excel, detectando registros duplicados y generando reportes detallados.

---

## 🎯 Características principales

### ✅ Duplicados Internos (NUEVO v2.2)
- Lee **múltiples hojas** del archivo Excel
- Detecta duplicados dentro del mismo archivo comparando **6 columnas clave** (F, G, H, I, L, Y)
- Genera reporte Excel con **3 hojas**:
  - Duplicados_Detalle
  - Auditoria_Resumen
  - Resumen_por_Hoja
- Proporciona trazabilidad completa (hoja origen, número de fila, etc.)
- Sin necesidad de base de datos (análisis independiente)

### ✅ Duplicados Históricos (Funcionalidad original)
- Lee y normaliza archivos PLE con las hojas `8.1` y `PROGRAMAS SOCIALES`
- Extrae columnas clave para detectar duplicados
- Guarda registros en base de datos SQLite rotativa (últimos 12 meses)
- Compara el nuevo mes con el histórico
- Genera informes de posibles duplicados
- Permite descargar reportes en formato Excel

---

## 📂 Estructura del proyecto

### 📁 Módulos principales

| Archivo | Descripción |
|---------|-------------|
| `app.py` | Interfaz Streamlit principal con 4 pestañas |
| `config.py` | Configuración de columnas, nombres de hojas, parámetros |
| `excel_reader.py` | Lectura y normalización de archivos Excel |
| `database.py` | Gestión de BD SQLite (creación, inserción, histórico) |
| `validator.py` | Comparación de meses y detección de duplicados históricos |
| `report_generator.py` | Generación de reportes Excel de duplicados históricos |

### 🆕 Módulos nuevos (v2.2)

| Archivo | Tamaño | Descripción |
|---------|--------|-------------|
| `duplicate_detector_internal.py` | 8.2 KB | Lectura y detección de duplicados internos |
| `duplicate_report_generator_internal.py` | 8.8 KB | Generación de reportes profesionales |

### 📂 Carpetas

| Carpeta | Contenido |
|---------|----------|
| `data/` | Base de datos SQLite (`ple_history.db`) |
| `reportes/` | Reportes generados |
| `uploads/` | Archivos temporales de Streamlit |

---

## 🚀 Instalación

### Paso 1: Crear entorno virtual

```bash
python -m venv .venv
```

### Paso 2: Activar entorno virtual

**Windows PowerShell:**
```powershell
.\.venv\Scripts\Activate.ps1
```

**Windows CMD:**
```cmd
.\.venv\Scripts\activate.bat
```

**Linux/Mac:**
```bash
source .venv/bin/activate
```

### Paso 3: Instalar dependencias

```bash
pip install -r requirements.txt
```

---

## 💻 Uso

### Iniciar la aplicación

```bash
streamlit run app.py
```

La aplicación abrirá en `http://localhost:8501` con 4 pestañas:

1. **📊 Validación** - Validar duplicados históricos (original)
2. **🔍 Duplicados Internos** - Buscar duplicados dentro del archivo (NUEVO)
3. **📖 Instrucciones** - Guía de uso
4. **ℹ️ Información** - Detalles del sistema

---

## 🔍 Modo 1: Duplicados Internos (NUEVO)

### ¿Qué detecta?

Registros con **exactamente los mismos valores** en 6 columnas clave:

| Columna | Contenido |
|---------|-----------|
| **F** | Fecha de Vencimiento o Fecha de Pago |
| **G** | Tipo de Comprobante de Pago o Documento |
| **H** | Serie del comprobante |
| **I** | Año de emisión |
| **L** | Tipo de Documento de Identidad del proveedor |
| **Y** | Importe total de las adquisiciones |

### Paso a paso

1. **Abrir pestaña**: Selecciona "🔍 Duplicados Internos"
2. **Cargar archivo**: Sube el Excel `.xlsx` que deseas validar
3. **Detectar**: Haz clic en "🔍 DETECTAR DUPLICADOS INTERNOS"
4. **Revisar resultados**:
   - Tabla resumen con estadísticas
   - Duplicados por hoja
   - Primeros registros encontrados
5. **Generar reporte**: Clic en "📥 Generar Reporte Excel"
6. **Descargar**: Descarga el archivo `duplicados_NOMBRE.xlsx`
7. **Auditoría** (opcional): Expande para ver detalles técnicos

### Reporte generado

Archivo: `duplicados_NOMBRE_ORIGINAL.xlsx`

**Hoja 1: Duplicados_Detalle**
- Todos los registros duplicados con TODAS sus columnas
- Columnas: Hoja Origen, Fila Original, + todas las columnas del Excel original

**Hoja 2: Auditoria_Resumen**
- Total de filas leídas
- Total de duplicados encontrados
- Cantidad de grupos únicos
- Fecha de procesamiento

**Hoja 3: Resumen_por_Hoja**
- Nombre de la hoja
- Cantidad de duplicados por hoja

### Ejemplo de duplicado

```
Registro A: 2025-03-02 | Factura | 001 | 2025 | DNI | 1000.00
Registro B: 2025-03-02 | Factura | 001 | 2025 | DNI | 1000.00
         ↑         ↑        ↑      ↑     ↑        ↑
        F         G        H      I     L        Y
             → DUPLICADO DETECTADO ←
```

---

## 📊 Modo 2: Duplicados Históricos (Original)

### Paso a paso

1. **Abrir pestaña**: Selecciona "📊 Validación"
2. **Subir archivo**: Archivo con nombre `PLE_COMPRAS_MMYYYY.xlsx`
   - Ejemplo: `PLE_COMPRAS_032025.xlsx`
   - MMYYYY = mes y año (032025 = marzo 2025)
3. **Validar**: Sistema busca duplicados vs histórico
4. **Revisar resultados**: 
   - Resumen de duplicados encontrados
   - Tabla detallada de coincidencias
5. **Descargar reporte**: Si hay duplicados
6. **Confirmar**: Si todo está correcto, agrega el mes al historial

### Historial rotativo

- Mantiene últimos **12 meses**
- Al agregar mes 13, elimina automáticamente el más antiguo
- Base de datos SQLite en `data/ple_history.db`

---

## 🔍 Validaciones y Normalizaciones

### Ambos modos aplican

1. **Espacios**: Se eliminan al inicio y final
2. **Mayúsculas**: Se normalizan para comparación
3. **Fechas**: Formato YYYY-MM-DD
4. **Valores nulos**: Se tratan como cadenas vacías
5. **Datos numéricos**: Se tratan como strings para consistencia

---

## 📋 Validaciones adicionales

### Validación de archivo
- ✅ Nombre debe contener MMYYYY
- ✅ Formato debe ser `.xlsx`
- ✅ Debe tener datos

### Validación de BD
- ✅ Hojas necesarias existen
- ✅ Columnas esperadas presentes
- ✅ Datos dentro de rangos especificados

---

## 🛠️ Uso Programático

### Ejemplo: Detectar duplicados internos

```python
from duplicate_detector_internal import leer_excel_todas_hojas, detectar_duplicados_internos
from duplicate_report_generator_internal import generar_reporte_duplicados_interno

# 1. Leer Excel (todas las hojas)
df_raw = leer_excel_todas_hojas("archivo.xlsx")

# 2. Detectar duplicados
df_duplicados, auditoria = detectar_duplicados_internos(df_raw)

# 3. Generar reporte
generar_reporte_duplicados_interno(
    df_raw,
    df_duplicados,
    auditoria,
    "duplicados_salida.xlsx",
    "archivo.xlsx"
)
```

### Ejemplo: Validar histórico

```python
from database import cargar_historico_completo
from validator import detectar_duplicados
import pandas as pd

# Cargar histórico
df_historico = cargar_historico_completo()

# Cargar nuevo mes (Excel)
df_nuevo = pd.read_excel("PLE_COMPRAS_032025.xlsx", sheet_name="8.1")

# Detectar duplicados
duplicados = detectar_duplicados(df_nuevo, df_historico)

# Usar resultados
print(f"Duplicados encontrados: {len(duplicados)}")
```

---

## 📦 Dependencias

Todas las dependencias ya están en `requirements.txt`:

```
pandas>=1.3.0          # Manipulación de datos
openpyxl>=3.7.0        # Lectura/escritura Excel
streamlit>=1.5.0       # Framework web
```

**Nota**: No se agregaron dependencias nuevas en v2.2

---

## 📊 Columnas clave utilizadas

El sistema trabaja con las siguientes columnas del PLE de compras:

| Col | Letra | Descripción |
|-----|-------|-------------|
| 1 | A | RUC |
| 2 | B | Período |
| 3 | C | Tipo de comprobante |
| 4 | D | Serie |
| 5 | E | Número |
| **6** | **F** | **Fecha Vencimiento/Pago** ⭐ |
| **7** | **G** | **Tipo Comprobante Pago** ⭐ |
| **8** | **H** | **Serie Comprobante** ⭐ |
| **9** | **I** | **Año de emisión** ⭐ |
| 10 | J | ... |
| **11** | **L** | **Tipo Doc. Identidad Proveedor** ⭐ |
| ... | ... | ... |
| **25** | **Y** | **Importe Total** ⭐ |

⭐ = Columnas usadas para detectar duplicados

---

## 🎯 Características especiales (v2.2)

✅ **Lectura flexible**: Primera hoja desde fila 8, otras desde fila 1
✅ **Múltiples hojas**: Lee TODAS las hojas, no limita a 2
✅ **Estilos profesionales**: Colores, bordes, alineamiento automático
✅ **Nombres automáticos**: `duplicados_NOMBRE_ORIGINAL.xlsx`
✅ **Auditoría completa**: Reporte detallado en Excel + texto
✅ **Interfaz intuitiva**: 100% integrado en Streamlit
✅ **Sin dependencias nuevas**: Usa librerías existentes
✅ **Robusto**: Manejo de errores completo
✅ **Trazabilidad**: Hoja origen + número de fila original

---

## ⚡ Mejoras de mantenibilidad (v2.2)

El código ha sido mejorado con:

✅ **+500 líneas de comentarios** explicativos
✅ **Estructura lógica clara** con secciones definidas
✅ **Manejo robusto de errores** (try-except en operaciones críticas)
✅ **Validaciones completas** en entrada de datos
✅ **Confirmaciones** en operaciones destructivas
✅ **Type hints** en funciones
✅ **Documentación exhaustiva** inline

---

## 🧪 Testing

La funcionalidad de duplicados internos incluye testing automático:

```
✅ Archivo de prueba creado con 14 registros
✅ Leído correctamente: múltiples hojas
✅ Detectados duplicados correctamente
✅ Reporte Excel generado sin errores
✅ Auditoría generada sin problemas
```

---

## 📝 Notas importantes

- ✅ El proyecto mantiene un historial rodante de hasta 12 meses
- ✅ Si existen 12 meses, al agregar uno nuevo se elimina el más antiguo
- ✅ El nombre del archivo debe contener `MMYYYY` (ej: 032025)
- ✅ Los duplicados internos se analizan **sin base de datos**
- ✅ El análisis histórico requiere **base de datos SQLite**
- ✅ Ambos modos generan reportes Excel profesionales

---

## 🆘 Solución de problemas

| Problema | Solución |
|----------|----------|
| ❌ "No hay datos en el archivo" | Verifica que sea `.xlsx` y tenga datos desde las filas especificadas |
| ❌ Error al leer hojas | Revisa que las hojas tengan datos desde fila 8 (primera) o fila 1 (otras) |
| ❌ "Nombre debe incluir MMYYYY" | Usa formato `PLE_COMPRAS_MMYYYY.xlsx` para histórico |
| ❌ Reporte no genera | Verifica que haya duplicados detectados |
| ❌ Interfaz no carga | Asegúrate de tener Streamlit >= 1.5 instalado |
| ❌ Error de BD | Ejecuta `reiniciar_bd.bat` para reiniciar la base de datos |

---

## 📋 Checklist - ¿Qué se incluye?

### Funcionalidad
- ✅ Lectura de múltiples hojas Excel
- ✅ Detección de duplicados internos (6 columnas)
- ✅ Comparación con histórico (12 meses)
- ✅ Normalización de datos (espacios, mayúsculas, fechas)
- ✅ Trazabilidad completa (hoja + fila original)
- ✅ Generación de reportes profesionales

### Interfaces
- ✅ 4 pestañas principales en Streamlit
- ✅ Sidebar con opciones de gestión
- ✅ Visualización de resultados
- ✅ Descarga de reportes
- ✅ Auditoría detallada

### Documentación
- ✅ README completo (este archivo)
- ✅ Guía de uso: `NUEVA_FUNCIONALIDAD_DUPLICADOS_INTERNOS.md`
- ✅ Documentación técnica: `RESUMEN_IMPLEMENTACION_V2.2.md`
- ✅ Guía de cambios: `GUIA_CAMBIOS_DETALLADA.md`
- ✅ Índice: `00_INDICE_CAMBIOS_V2.2.txt`

### Calidad
- ✅ Manejo robusto de errores
- ✅ Validaciones de entrada
- ✅ Confirmaciones de operaciones destructivas
- ✅ Type hints en funciones
- ✅ Tests exitosos
- ✅ Listo para producción

---

## 📞 Documentación adicional

- **📖 Guía completa de duplicados internos**: `NUEVA_FUNCIONALIDAD_DUPLICADOS_INTERNOS.md`
- **📊 Resumen técnico v2.2**: `RESUMEN_IMPLEMENTACION_V2.2.md`
- **🛠️ Guía detallada de cambios**: `GUIA_CAMBIOS_DETALLADA.md`
- **📚 Índice rápido**: `00_INDICE_CAMBIOS_V2.2.txt`

---

## 🎉 Resumen ejecutivo

Se ha implementado exitosamente una **nueva funcionalidad v2.2** que permite detectar automáticamente duplicados dentro de un mismo archivo Excel del PLE_COMPRAS, con trazabilidad completa y generación de reportes profesionales. La funcionalidad está completamente integrada en la interfaz Streamlit existente sin requerir cambios en la base de datos ni dependencias adicionales.

El sistema ahora ofrece **dos modos complementarios** de análisis:
1. **Análisis interno**: Duplicados dentro del mismo archivo (sin BD)
2. **Análisis histórico**: Duplicados vs últimos 12 meses (con BD)

**Estado**: ✅ **LISTO PARA PRODUCCIÓN v2.2**

---

**Versión**: 2.2  
**Última actualización**: 2025-06-02  
**Desarrollado por**: YESS - Copilot
