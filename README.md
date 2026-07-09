# 📊 PLE COMPRAS

Una aplicación web robusta para validación, auditoría y conciliación de archivos PLE (Programa de Libros Electrónicos) de compras según la normativa SUNAT de Perú.

## 🎯 Características Principales

- **Detección de Duplicados**: Identifica registros duplicados dentro del mismo archivo y entre meses
- **Auditoría Interna**: Valida consistencia e integridad de datos en archivos PLE
- **Conciliación de Archivos**: Compara dos archivos PLE para identificar discrepancias
- **Reconciliación SIRE-SUNAT**: Valida datos contra registros SIRE de SUNAT
- **Histórico de 12 Meses**: Mantiene base de datos con últimos 12 meses para comparativas
- **Reportes Profesionales**: Genera reportes en Excel con formato y validaciones

## 🚀 Inicio Rápido

### Requisitos Previos

- Python 3.8+
- pip (gestor de paquetes de Python)

### Instalación

1. **Clonar el repositorio**
```bash
git clone https://github.com/yessLU11/Validador-PLE-COMPRAS.git
cd Validador-PLE-COMPRAS
```

2. **Crear un entorno virtual (recomendado)**
```bash
python -m venv venv
# En Windows:
venv\Scripts\activate
# En Linux/Mac:
source venv/bin/activate
```

3. **Instalar dependencias**
```bash
pip install -r requirements.txt
```

### Ejecutar la Aplicación

```bash
streamlit run app.py
```

La aplicación se abrirá en `http://localhost:8501`

## 📁 Estructura del Proyecto

```
Validador-PLE-COMPRAS/
├── app.py                              # Aplicación principal (Streamlit)
├── config.py                           # Configuración de parámetros
├── validator.py                        # Lógica de comparación y detección de duplicados
├── conciliador.py                      # Lógica de conciliación de archivos
├── duplicate_detector_internal.py      # Detección de duplicados internos
├── duplicate_report_generator_internal.py # Generador de reportes de duplicados
├── database.py                         # Gestión de base de datos SQLite
├── excel_reader.py                     # Lectura de archivos Excel
├── report_generator.py                 # Generación de reportes profesionales
├── requirements.txt                    # Dependencias del proyecto
├── data/                               # Carpeta para almacenamiento de datos
├── uploads/                            # Carpeta para archivos cargados temporalmente
├── reportes/                           # Carpeta de reportes generados
└── README.md                           # Este archivo
```

## 🔧 Funcionalidades Detalladas

### 1. Validación de Duplicados

**¿Qué hace?**
- Compara datos del mes actual con archivos de meses anteriores
- Detecta registros que ya existen en la base de datos histórica
- Identifica duplicados dentro del mismo archivo

**¿Cómo funciona?**
- Utiliza merge (inner join) en columnas clave definidas en `config.py`
- Compara: fecha emisión, tipo comprobante, serie, número, RUC proveedor, razón social, montos
- Genera reporte detallado con información de meses duplicados

**Columnas Clave para Comparación:**
```python
- fecha_emision
- tipo_comprobante
- serie_comprobante
- numero_comprobante
- ruc_proveedor
- razon_social
- base_imponible
- igv
- importe_total
```

### 2. Auditoría Interna

**¿Qué hace?**
- Valida la consistencia e integridad de datos dentro de un archivo
- Detecta duplicados internos antes de cargar en base de datos
- Identifica registros problemáticos

**¿Cómo funciona?**
- Lectura de todas las hojas del archivo Excel
- Normalización de datos (tipos de documentos, formatos, etc.)
- Análisis de duplicados usando columnas identificadoras

### 3. Conciliación de Archivos

**¿Qué hace?**
- Compara dos archivos PLE completos
- Identifica registros presentes en uno pero no en otro
- Genera reportes de discrepancias

**¿Cómo funciona?**
- Lee ambos archivos y normaliza sus datos
- Crea IDs únicos basados en: tipo_comprobante + serie + número
- Compara sets de registros
- Genera reporte con registros no coincidentes

### 4. Reconciliación SIRE-SUNAT

**¿Qué hace?**
- Valida datos PLE contra registros SIRE de SUNAT
- Identifica discrepancias en montos y datos de operaciones
- Genera reporte de inconsistencias

## 📊 Configuración (config.py)

Modifica los siguientes parámetros según tus necesidades:

```python
# Columnas clave para comparación
COLUMNAS_CLAVE = [
    "fecha_emision", "tipo_comprobante", "serie_comprobante",
    "numero_comprobante", "ruc_proveedor", "razon_social",
    "base_imponible", "igv", "importe_total"
]

# Hojas del archivo Excel a procesar
HOJA_PRINCIPAL = "8.1"
HOJA_SOCIALES = "Programas Sociales"

# Filas de inicio de datos
FILA_INICIO_PRINCIPAL = 8
FILA_INICIO_SOCIALES = 1

# Meses de histórico a mantener
MESES_A_MANTENER = 12
```

## 💾 Base de Datos

La aplicación utiliza SQLite para almacenar:
- Registros de todos los meses cargados
- Metadatos (fecha carga, período, hoja origen)
- Histórico completo para comparativas

**Archivo de base de datos:** `data/database.db`

### Operaciones Disponibles

```python
# Inicializar base de datos
init_db()

# Obtener meses existentes
obtener_meses_existentes()

# Agregar nuevo mes
agregar_mes(df, periodo, mes_archivo)

# Cargar histórico completo
cargar_historico_completo(excluir_mes=None)

# Eliminar último mes cargado
eliminar_ultimo_mes()

# Eliminar toda la base de datos
eliminar_toda_base_datos()
```

## 📝 Formato de Archivos

### Archivo PLE Compras Esperado

El archivo Excel debe contener:
- **Hoja "8.1"**: Registros de compras principales (a partir de fila 8)
- **Hoja "Programas Sociales"**: Compras de programas sociales (a partir de fila 1)

**Columnas Requeridas:**
- B: Período (AAAAMM00)
- C: Código Único de la Operación (CUO)
- E: Fecha de emisión
- G: Tipo de comprobante
- H: Serie del comprobante
- J: Número del comprobante
- M: RUC del proveedor
- N: Razón social del proveedor
- Q: Base imponible
- R: IGV
- Y: Importe total

## 🎨 Interfaz Streamlit

La aplicación incluye:
- **Sidebar**: Navegación entre funcionalidades principales
- **Carga de Archivos**: Interface para subir archivos PLE
- **Gestión de Base de Datos**: Ver, agregar, eliminar meses
- **Generación de Reportes**: Descargar reportes en Excel
- **Visualización de Datos**: Tablas interactivas con resultados

### Secciones Disponibles

1. **Cargar Nuevo Mes**: Sube y procesa un archivo PLE
2. **Detectar Duplicados**: Compara mes actual con histórico
3. **Auditoría Interna**: Valida consistencia del archivo
4. **Conciliación**: Compara dos archivos PLE
5. **Gestión BD**: Visualiza y gestiona datos en base de datos

## 📊 Reportes Generados

### Tipos de Reportes

1. **Reporte de Duplicados**: 
   - Registros duplicados entre meses
   - Información de meses donde aparecen

2. **Reporte de Auditoría Interna**:
   - Duplicados dentro del mismo archivo
   - Problemas de integridad de datos

3. **Reporte de Conciliación**:
   - Registros presentes/no presentes en cada archivo
   - Análisis de discrepancias

4. **Reporte SIRE-SUNAT**:
   - Validación contra registros SUNAT
   - Inconsistencias identificadas

### Formato de Reportes

- **Formato**: Excel (.xlsx)
- **Estilos**: Colores, bordes, alineación profesional
- **Filtros**: Activados para facilitar análisis
- **Ubicación**: Carpeta `reportes/`

## 🔌 Dependencias Principales

```
Flask-SocketIO==5.5.1
Streamlit==1.57.0
pandas==2.2.2
openpyxl==3.1.5
SQLAlchemy==2.0.49
reportlab==4.4.4
```

Ver `requirements.txt` para lista completa de dependencias.

## 🛠️ Desarrollo

### Extensiones Posibles

1. Agregar validación de impuestos
2. Integración con API SUNAT
3. Exportación a otros formatos (CSV, JSON)
4. Dashboard de análisis avanzado
5. Validación de reglas de negocio personalizadas

### Contribuciones

Las contribuciones son bienvenidas. Por favor:
1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto está bajo licencia MIT. Ver archivo `LICENSE` para más detalles.

## 👤 Autor

**Yessly Poma**

## 📞 Soporte

Para reportar errores, sugerencias o preguntas:
- Abre un issue en el repositorio
- Contacta al equipo de desarrollo

## 🔒 Privacidad y Seguridad

- Los archivos cargados se procesan localmente
- Los datos se almacenan en base de datos SQLite local
- No se envía información a servidores externos
- Recomendado usar en ambiente seguro con acceso restringido

## 📚 Referencia de Normativa

- **PLE (Programa de Libros Electrónicos)**: Regulado por SUNAT (Superintendencia Nacional de Aduanas y de Administración Tributaria)
- **Formato**: Conforme a Resolución de Superintendencia SUNAT
- **Validación**: Cumple estándares de formato y contenido SUNAT

---

**Versión**: 1.0.0  
**Última actualización**: 2026-07-09  
**Estado**: Producción ✅
