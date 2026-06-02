# 🔍 NUEVA FUNCIONALIDAD: Validador de Duplicados Internos

## 📋 Resumen de lo implementado

Se ha agregado una **nueva pestaña** en la aplicación Streamlit para detectar duplicados **dentro de un mismo archivo Excel**. Esta función:

1. **Lee múltiples hojas** del archivo Excel
   - Primera hoja: desde fila 8
   - Resto de hojas: desde fila 1

2. **Detecta duplicados** comparando columnas clave:
   - Columna F: Fecha de Vencimiento o Fecha de Pago
   - Columna G: Tipo de Comprobante de Pago o Documento
   - Columna H: Serie del comprobante de pago o documento
   - Columna I: Año de emisión de la DUA o DSI
   - Columna L: Tipo de Documento de Identidad del proveedor
   - Columna Y: Importe total de las adquisiciones

3. **Genera reportes Excel** con:
   - Hoja "Duplicados_Detalle": todas las filas duplicadas con TODAS sus columnas
   - Hoja "Auditoria_Resumen": estadísticas generales
   - Hoja "Resumen_por_Hoja": cantidad de duplicados por hoja

4. **Proporciona trazabilidad completa**:
   - Hoja origen
   - Número de fila original
   - Cantidad total de duplicados
   - Grupos de duplicados encontrados

---

## 🚀 ¿CÓMO USAR?

### PASO 1: Abrir la nueva pestaña

1. Ejecuta la aplicación: `streamlit run app.py`
2. En las pestañas principales, selecciona **"🔍 Duplicados Internos"**

### PASO 2: Cargar archivo Excel

1. En la sección "📂 Cargar archivo", haz clic en **"Selecciona un archivo Excel"**
2. Elige el archivo `.xlsx` que quieres validar
3. El sistema leerá automáticamente todas las hojas
4. Verás un mensaje: ✅ "Archivo cargado: NOMBRE_ARCHIVO.xlsx (X registros leídos)"

### PASO 3: Detectar duplicados

1. Haz clic en el botón **"🔍 DETECTAR DUPLICADOS INTERNOS"**
2. El sistema analizará el archivo

### PASO 4: Revisar resultados

El sistema mostrará:

**Estadísticas principales:**
- 📋 Total de filas (registros leídos)
- ⚠️ Duplicados encontrados (cantidad total)
- 📊 Grupos únicos (cuántos conjuntos de duplicados hay)

**Tabla por hoja:**
- Hoja origen
- Cantidad de duplicados en esa hoja

**Primeros registros:**
- Muestra los primeros 20 duplicados encontrados
- Si hay más, indica que descargues el Excel para verlos todos

### PASO 5: Descargar reporte

1. Haz clic en **"📥 Generar Reporte Excel"**
2. Se creará un archivo: `duplicados_NOMBRE_ORIGINAL.xlsx`
3. Haz clic en **"📥 Descargar Reporte Excel Completo"**
4. El archivo se descargará a tu computadora

### PASO 6 (Opcional): Ver auditoría detallada

- Haz clic en "📋 Ver auditoría detallada" (expandible)
- Verás un reporte texto con:
  - Resumen de cantidad de duplicados
  - Duplicados agrupados por hojas
  - Detalles de cada grupo con número de fila

---

## 📊 EJEMPLO DE SALIDA

### Archivo generado: `duplicados_PLE_COMPRAS_032025.xlsx`

**Hoja 1: "Duplicados_Detalle"**
| Hoja Origen | Fila Original | ColA | ColB | ColC | ... ColY |
|-----------|---------|------|------|------|---------|
| Hoja1     | 8       | 123  | ABC  | 20   | 1000.00 |
| Hoja1     | 12      | 123  | ABC  | 20   | 1000.00 |
| Hoja2     | 5       | 123  | ABC  | 20   | 1000.00 |
| Hoja3     | 14      | 456  | DEF  | 25   | 2500.00 |
| ...       | ...     | ...  | ...  | ...  | ...     |

**Hoja 2: "Auditoria_Resumen"**
| Métrica                        | Valor            |
|-------------------------------|------------------|
| Total de filas leídas          | 500              |
| Total de registros duplicados  | 45               |
| Grupos únicos de duplicados    | 12               |
| Archivo original               | PLE_COMPRAS.xlsx |
| Fecha de procesamiento         | 2025-03-02 10:30 |

**Hoja 3: "Resumen_por_Hoja"**
| Hoja          | Cantidad de Duplicados |
|---------------|----------------------|
| Hoja1 (8.1)   | 25                   |
| Hoja2 (PROG)  | 20                   |

---

## 🔧 ARCHIVOS CREADOS/MODIFICADOS

### Nuevos archivos creados:

1. **`duplicate_detector_internal.py`** (8.2 KB)
   - Función: `leer_excel_todas_hojas()` - Lee múltiples hojas
   - Función: `detectar_duplicados_internos()` - Detecta duplicados
   - Función: `obtener_todas_columnas_originales()` - Extrae datos
   - Función: `generar_auditoria_duplicados()` - Crea auditoría

2. **`duplicate_report_generator_internal.py`** (8.8 KB)
   - Función: `generar_reporte_duplicados_interno()` - Crea Excel
   - Función: `_aplicar_estilos()` - Estilos profesionales
   - Función: `generar_archivo_descarga()` - Nombre automático

### Archivos modificados:

1. **`app.py`**
   - Agregó importes de nuevos módulos
   - Agregó variables de sesión para duplicados internos
   - Agregó nueva pestaña "🔍 Duplicados Internos"
   - Actualizó numeración de pestañas (ahora 4 en total)

---

## 💾 ALMACENAMIENTO

Los reportes generados se guardan en:
- **Carpeta:** `reportes/`
- **Nombre:** `duplicados_NOMBRE_ORIGINAL.xlsx`
- **Ejemplo:** `duplices_PLE_COMPRAS_032025.xlsx`

Los archivos cargados se guardan temporalmente en:
- **Carpeta:** `uploads/`

---

## 🎯 LÓGICA DE DETECCIÓN

### ¿Cómo define "duplicado"?

Se considera duplicado un registro que tiene **exactamente** los mismos valores en las 6 columnas clave:

```
F (Fecha) + G (Tipo Comprobante) + H (Serie) + I (Año) + L (Doc. Proveedor) + Y (Importe)
```

**Ejemplo:**
- Registro 1: 2025-03-02 | Factura | 001 | 2025 | DNI | 1000.00 ✓
- Registro 2: 2025-03-02 | Factura | 001 | 2025 | DNI | 1000.00 ✓ **← DUPLICADO**
- Registro 3: 2025-03-02 | Factura | 001 | 2025 | DNI | 2000.00 ✗ (Diferente importe)

---

## ⚙️ CONFIGURACIÓN

### Ubicación de lectura por hoja:

- **Primera hoja (Hoja1):** Comienza lectura desde fila 8 (índice 7 en Python)
- **Otras hojas:** Comienzan lectura desde fila 1 (índice 0 en Python)

Esto es configurable en `duplicate_detector_internal.py`, línea ~40:
```python
fila_inicio = 7 if idx_hoja == 0 else 0  # Cambiar según necesidad
```

### Columnas comparadas:

Están definidas en `duplicate_detector_internal.py`, línea ~56:
```python
columnas_duplicado = {
    "F": 5,    # Fecha
    "G": 6,    # Tipo Comprobante
    "H": 7,    # Serie
    "I": 8,    # Año
    "L": 11,   # Documento Proveedor
    "Y": 24    # Importe
}
```

---

## 📝 NOTAS IMPORTANTES

1. **Normalización:** Los valores se normalizan (espacios eliminados, mayúsculas) antes de comparar
2. **Múltiples hojas:** Lee TODAS las hojas del Excel, no solo las primeras
3. **Reporte profesional:** Los Excel generados tienen estilos de colores y formatos
4. **Sin conexión externa:** Todo ocurre en tu computadora, sin enviar datos a internet
5. **Reutilizable:** Puedes cargar múltiples archivos secuencialmente

---

## 🧪 PRUEBA

Se incluye un script de prueba: `test_duplicate_detector.py`

Ejecutar:
```bash
python test_duplicate_detector.py
```

Este script:
- Crea un archivo Excel de prueba
- Detecta duplicados
- Genera un reporte
- Muestra auditoría detallada

---

## 📞 SOPORTE

Si encuentras problemas:

1. **Verifica que el archivo sea `.xlsx`** (no `.xls`)
2. **Verifica que las hojas tengan datos** desde las filas especificadas
3. **Revisa la consola** para mensajes de error detallados
4. **Ejecuta el test** para verificar que todo funciona

---

## ✅ CHECKLIST DE FUNCIONALIDADES

- ✅ Lee múltiples hojas
- ✅ Detecta duplicados basándose en 6 columnas clave
- ✅ Genera Excel con 3 hojas (Detalle, Auditoría, Resumen)
- ✅ Trazabilidad completa (hoja, fila)
- ✅ Auditoría detallada
- ✅ Interfaz Streamlit integrada
- ✅ Estilos profesionales en reportes
- ✅ Nombres de archivo automáticos

---

**Versión:** 2.2 con nueva funcionalidad de Duplicados Internos
**Fecha:** 2025-06-02
