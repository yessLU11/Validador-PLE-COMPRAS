# Validador PLE Compras

Este proyecto es una herramienta para validar archivos PLE de compras en Excel, comparando un nuevo mes contra los últimos 12 meses almacenados en un historial SQLite.

## Qué hace
- Lee y normaliza archivos PLE con las hojas `8.1` y `PROGRAMAS SOCIALES`.
- Extrae las columnas clave necesarias para detectar duplicados.
- Guarda los registros en una base de datos SQLite rotativa que mantiene solo los últimos 12 meses.
- Compara el nuevo mes con el histórico y genera un informe de posibles duplicados.
- Permite descargar un reporte de duplicados en formato Excel.

## Estructura del proyecto
- `app.py`: interfaz Streamlit y punto de entrada principal.
- `config.py`: definiciones de columnas clave, nombres de hojas y parámetros del historial.
- `excel_reader.py`: lee y normaliza las hojas del archivo Excel.
- `database.py`: crea la base de datos SQLite, inserta registros y mantiene el histórico.
- `validator.py`: compara el mes nuevo con el histórico y detecta duplicados.
- `report_generator.py`: genera un reporte Excel de duplicados.
- `requirements.txt`: dependencias necesarias.
- `data/`: carpeta donde se almacena la base de datos `ple_history.db`.
- `uploads/`: carpeta temporal para subir archivos en Streamlit.

## Instalación
1. Crear un entorno virtual de Python:

```bash
python -m venv .venv
```

2. Activar el entorno virtual:

- Windows PowerShell:
  ```powershell
  .\.venv\Scripts\Activate.ps1
  ```
- Windows CMD:
  ```cmd
  .\.venv\Scripts\activate.bat
  ```

3. Instalar dependencias:

```bash
pip install -r requirements.txt
```

## Uso
1. Iniciar la app Streamlit:

```bash
streamlit run app.py
```

2. Subir un archivo PLE con nombre que contenga `MMYYYY`, por ejemplo `PLE_COMPRAS_032025.xlsx`.
3. Revisar los duplicados detectados en la interfaz.
4. Generar y descargar el reporte de duplicados si se encuentran coincidencias.
5. Confirmar la incorporación del mes al historial.

## Notas importantes
- El proyecto mantiene un historial rodante de hasta 12 meses.
- Si ya existen 12 meses, al agregar uno nuevo se elimina el mes más antiguo.
- El nombre del archivo debe contener un bloque de 6 dígitos con el mes y año (`MMYYYY`).

## Recomendaciones
- Revisar la estructura del archivo Excel para que las hojas y columnas esperadas existan.
- Usar un archivo de ejemplo real para validar la detección de duplicados antes de procesar datos definitivos.
