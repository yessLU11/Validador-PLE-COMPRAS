 # Manejo de SQLite (crear tablas, agregar mes, eliminar, consultar)
# database.py
# database.py
import sqlite3
import pandas as pd
import os
from config import MESES_A_MANTENER

DB_PATH = "data/ple_history.db"

def _is_valid_sqlite_db(path):
    """Verifica si el archivo es una base de datos SQLite válida."""
    if not os.path.exists(path):
        return False
    try:
        conn = sqlite3.connect(path)
        conn.execute("SELECT 1")
        conn.close()
        return True
    except sqlite3.DatabaseError:
        return False

def init_db():
    """Crea la tabla si no existe. Si el archivo es inválido, lo elimina y recrea.
    Si la tabla existe pero no tiene las columnas nuevas, las agrega."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    
    if os.path.exists(DB_PATH) and not _is_valid_sqlite_db(DB_PATH):
        print(f"Archivo {DB_PATH} corrupto o no es SQLite. Eliminando...")
        os.remove(DB_PATH)
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Crear tabla si no existe
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS registros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mes_archivo TEXT NOT NULL,
            hoja_origen TEXT,
            "Periodo (AAAAMM00)" TEXT,
            "Código Único de la Operación (CUO)" TEXT,
            fecha_emision TEXT,
            tipo_comprobante TEXT,
            serie_comprobante TEXT,
            numero_comprobante TEXT,
            ruc_proveedor TEXT,
            razon_social TEXT,
            base_imponible TEXT,
            igv TEXT,
            importe_total TEXT,
            fila_original INTEGER DEFAULT 0,
            nombre_archivo TEXT
        )
    """)
    
    # Verificar y agregar columnas faltantes si la tabla ya existía
    cursor.execute("PRAGMA table_info(registros)")
    columnas_existentes = {row[1] for row in cursor.fetchall()}
    
    columnas_faltantes = {
        "Periodo (AAAAMM00)": 'TEXT',
        "Código Único de la Operación (CUO)": 'TEXT'
    }
    
    for col_name, col_type in columnas_faltantes.items():
        if col_name not in columnas_existentes:
            try:
                cursor.execute(f'ALTER TABLE registros ADD COLUMN "{col_name}" {col_type}')
                print(f"Columna '{col_name}' agregada a la tabla registros")
            except sqlite3.OperationalError as e:
                print(f"Error agregando columna '{col_name}': {e}")
    
    # Índice con las columnas clave para búsqueda de duplicados
    columnas_idx = [
        "fecha_emision", "tipo_comprobante", "serie_comprobante", "numero_comprobante",
        "ruc_proveedor", "razon_social", "base_imponible", "igv", "importe_total"
    ]
    cursor.execute(f"CREATE INDEX IF NOT EXISTS idx_claves ON registros ({','.join(columnas_idx)})")
    conn.commit()
    conn.close()

def obtener_meses_existentes():
    """Devuelve lista de meses (YYYYMM) presentes en la BD ordenados."""
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT DISTINCT mes_archivo FROM registros ORDER BY mes_archivo", conn)
    conn.close()
    return df["mes_archivo"].tolist()

def agregar_mes(df_nuevo, mes_str, nombre_archivo=""):
    """
    Agrega un nuevo mes a la BD, eliminando el más antiguo si ya hay 12 meses.
    df_nuevo: DataFrame con las mismas columnas (incluyendo mes_archivo, hoja_origen)
    mes_str: YYYYMM (ej '202503')
    Retorna: (exito, mes_eliminado)
    """
    conn = sqlite3.connect(DB_PATH)
    # 1. Obtener meses actuales
    meses = pd.read_sql_query("SELECT DISTINCT mes_archivo FROM registros ORDER BY mes_archivo", conn)
    meses_list = meses["mes_archivo"].tolist()
    
    mes_eliminado = None
    if len(meses_list) >= MESES_A_MANTENER:
        mes_mas_antiguo = meses_list[0]
        conn.execute("DELETE FROM registros WHERE mes_archivo = ?", (mes_mas_antiguo,))
        mes_eliminado = mes_mas_antiguo
    
    # 2. Insertar el nuevo mes
    df_nuevo["nombre_archivo"] = nombre_archivo
    df_nuevo.to_sql("registros", conn, if_exists="append", index=False)
    conn.commit()
    conn.close()
    return True, mes_eliminado

def eliminar_ultimo_mes():
    """
    Elimina el mes más reciente (mayor YYYYMM) de la base de datos.
    Retorna el mes eliminado o None si no hay registros.
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT mes_archivo FROM registros ORDER BY mes_archivo DESC LIMIT 1")
    row = cursor.fetchone()
    if not row:
        conn.close()
        return None
    mes_eliminar = row[0]
    cursor.execute("DELETE FROM registros WHERE mes_archivo = ?", (mes_eliminar,))
    conn.commit()
    conn.close()
    return mes_eliminar

def cargar_historico_completo():
    """Retorna un DataFrame con todos los registros de los últimos 12 meses."""
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM registros", conn)
    conn.close()
    return df

def eliminar_toda_base_datos():
    """Elimina todos los registros de la tabla, dejando la BD vacía."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM registros")
    conn.commit()
    conn.close()

def cargar_multiples_archivos(lista_archivos):
    """
    Recibe una lista de tuplas (ruta_archivo, nombre_archivo) 
    o directamente rutas. Cada archivo se lee y se añade a la BD,
    manteniendo el rolling de 12 meses.
    Si ya existen meses, se combinan y luego se eliminan los más antiguos
    hasta dejar solo los 12 más recientes.
    Retorna (exito, mensaje, meses_agregados, meses_eliminados)
    """
    from excel_reader import leer_hoja_unificada
    import re
    
    todos_dfs = []
    meses_subidos = []
    
    for ruta, nombre in lista_archivos:
        match = re.search(r"(\d{6})", nombre)
        if not match:
            continue
        mes = match.group(1)
        df = leer_hoja_unificada(ruta, mes)
        if not df.empty:
            df["nombre_archivo"] = nombre
            todos_dfs.append(df)
            meses_subidos.append(mes)
    
    if not todos_dfs:
        return False, "No se pudo leer ningún archivo válido.", [], []
    
    df_total = pd.concat(todos_dfs, ignore_index=True)
    
    conn = sqlite3.connect(DB_PATH)
    # Primero añadimos los nuevos registros
    df_total.to_sql("registros", conn, if_exists="append", index=False)
    
    # Ahora aplicar rolling: obtener todos los meses distintos, ordenar
    meses_df = pd.read_sql_query("SELECT DISTINCT mes_archivo FROM registros ORDER BY mes_archivo", conn)
    meses_lista = meses_df["mes_archivo"].tolist()
    eliminados = []
    while len(meses_lista) > MESES_A_MANTENER:
        mes_viejo = meses_lista.pop(0)
        conn.execute("DELETE FROM registros WHERE mes_archivo = ?", (mes_viejo,))
        eliminados.append(mes_viejo)
    
    conn.commit()
    conn.close()
    return True, f"Procesados {len(todos_dfs)} archivos.", meses_subidos, eliminados
# Añadir al final de database.py

def eliminar_toda_base_datos():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM registros")
    conn.commit()
    conn.close()

def cargar_multiples_archivos(lista_archivos):
    from excel_reader import leer_hoja_unificada
    import re
    todos_dfs = []
    meses_subidos = []
    for ruta, nombre in lista_archivos:
        match = re.search(r"(\d{6})", nombre)
        if not match:
            continue
        mes = match.group(1)
        df = leer_hoja_unificada(ruta, mes)
        if not df.empty:
            df["nombre_archivo"] = nombre
            todos_dfs.append(df)
            meses_subidos.append(mes)
    if not todos_dfs:
        return False, "No se pudo leer ningún archivo válido.", [], []
    df_total = pd.concat(todos_dfs, ignore_index=True)
    conn = sqlite3.connect(DB_PATH)
    df_total.to_sql("registros", conn, if_exists="append", index=False)
    # Rolling
    meses_df = pd.read_sql_query("SELECT DISTINCT mes_archivo FROM registros ORDER BY mes_archivo", conn)
    meses_lista = meses_df["mes_archivo"].tolist()
    eliminados = []
    while len(meses_lista) > MESES_A_MANTENER:
        mes_viejo = meses_lista.pop(0)
        conn.execute("DELETE FROM registros WHERE mes_archivo = ?", (mes_viejo,))
        eliminados.append(mes_viejo)
    conn.commit()
    conn.close()
    return True, f"Procesados {len(todos_dfs)} archivos.", meses_subidos, eliminados