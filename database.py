# database.py - VERSIÓN REFACTORIZADA (SIN DEPENDENCIAS DE STREAMLIT)
import sqlite3
import pandas as pd
import os
from config import MESES_A_MANTENER

DB_PATH = "data/ple_history.db"

def _is_valid_sqlite_db(path):
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
    """Crea las tablas si no existen."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    
    if os.path.exists(DB_PATH) and not _is_valid_sqlite_db(DB_PATH):
        print(f"Archivo {DB_PATH} corrupto. Eliminando...")
        os.remove(DB_PATH)
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Tabla registros
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
    
    # Verificar columnas faltantes
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
            except sqlite3.OperationalError:
                pass
    
    # Índices
    columnas_idx = [
        "fecha_emision", "tipo_comprobante", "serie_comprobante", "numero_comprobante",
        "ruc_proveedor", "razon_social", "base_imponible", "igv", "importe_total"
    ]
    cursor.execute(f"CREATE INDEX IF NOT EXISTS idx_claves ON registros ({','.join(columnas_idx)})")
    
    # Tabla duplicados_reportados
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS duplicados_reportados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mes_archivo TEXT NOT NULL,
            fecha_emision TEXT NOT NULL,
            tipo_comprobante TEXT NOT NULL,
            serie_comprobante TEXT NOT NULL,
            numero_comprobante TEXT NOT NULL,
            ruc_proveedor TEXT NOT NULL,
            razon_social TEXT NOT NULL,
            base_imponible TEXT NOT NULL,
            igv TEXT NOT NULL,
            importe_total TEXT NOT NULL,
            fecha_reporte TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(mes_archivo, fecha_emision, tipo_comprobante, serie_comprobante, 
                   numero_comprobante, ruc_proveedor, razon_social, base_imponible, igv, importe_total)
        )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_reportados_mes ON duplicados_reportados(mes_archivo)")
    
    conn.commit()
    conn.close()

# ======= FUNCIONES AUXILIARES DE ORDENAMIENTO =======
def _mes_a_orden(mes_archivo):
    """Convierte MMYYYY a tupla (YYYY, MM)."""
    if not isinstance(mes_archivo, str) or len(mes_archivo) < 6 or not mes_archivo.isdigit():
        return (9999, 99)
    return (int(mes_archivo[2:6]), int(mes_archivo[0:2]))

def ordenar_meses(meses, reverse=False):
    """Ordena cronológicamente una lista de meses MMYYYY."""
    return sorted(list(dict.fromkeys(meses)), key=_mes_a_orden, reverse=reverse)

def mes_existe(mes_archivo):
    """Retorna True si el mes ya existe en la BD."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM registros WHERE mes_archivo = ? LIMIT 1", (mes_archivo,))
    existe = cursor.fetchone() is not None
    conn.close()
    return existe

# ======= OPERACIONES CRUD =======
def obtener_meses_existentes():
    """Devuelve lista de meses (MMYYYY) presentes en la BD ordenados cronológicamente."""
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query(
        "SELECT DISTINCT mes_archivo FROM registros ORDER BY CAST(SUBSTR(mes_archivo, 3, 4) AS INTEGER), CAST(SUBSTR(mes_archivo, 1, 2) AS INTEGER)",
        conn
    )
    conn.close()
    return df["mes_archivo"].tolist()

def agregar_mes(df_nuevo, mes_str, nombre_archivo=""):
    """
    Agrega un nuevo mes a la BD. Reemplaza si ya existe.
    Si hay más de 12 meses, elimina el más antiguo.
    """
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # 1. Reemplazar versión anterior del mismo mes (si existe)
    cursor.execute("DELETE FROM registros WHERE mes_archivo = ?", (mes_str,))
    
    # 2. Obtener meses actuales (ordenados)
    meses_df = pd.read_sql_query(
        "SELECT DISTINCT mes_archivo FROM registros ORDER BY CAST(SUBSTR(mes_archivo, 3, 4) AS INTEGER), CAST(SUBSTR(mes_archivo, 1, 2) AS INTEGER)",
        conn
    )
    meses_list = meses_df["mes_archivo"].tolist()
    
    mes_eliminado = None
    if len(meses_list) >= MESES_A_MANTENER:
        mes_mas_antiguo = meses_list[0]
        cursor.execute("DELETE FROM registros WHERE mes_archivo = ?", (mes_mas_antiguo,))
        mes_eliminado = mes_mas_antiguo
    
    # 3. Insertar el nuevo mes
    df_nuevo["nombre_archivo"] = nombre_archivo
    df_nuevo.to_sql("registros", conn, if_exists="append", index=False)
    conn.commit()
    conn.close()
    return True, mes_eliminado

def eliminar_ultimo_mes():
    """Elimina el mes más reciente de la base de datos."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT DISTINCT mes_archivo FROM registros ORDER BY CAST(SUBSTR(mes_archivo, 3, 4) AS INTEGER) DESC, CAST(SUBSTR(mes_archivo, 1, 2) AS INTEGER) DESC LIMIT 1"
    )
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
    """Retorna un DataFrame con todos los registros."""
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM registros", conn)
    conn.close()
    return df

def eliminar_toda_base_datos():
    import sys
    print("🔍 DEBUG: Iniciando eliminación de toda la base de datos...", flush=True)
    print(f"🔍 Ruta de la BD: {os.path.abspath(DB_PATH)}", flush=True)
    
    if not os.path.exists(DB_PATH):
        print("⚠️ El archivo de la BD no existe. No hay nada que eliminar.", flush=True)
        return
    
    try:
        conn = sqlite3.connect(DB_PATH)
        # Verificar cuántos registros hay antes de eliminar
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM registros")
        count = cursor.fetchone()[0]
        print(f"📊 Registros antes de eliminar: {count}", flush=True)
        
        conn.execute("DELETE FROM registros")
        conn.commit()
        conn.close()
        print("✅ Eliminación completada con éxito.", flush=True)
    except Exception as e:
        print(f"❌ Error durante la eliminación: {e}", flush=True)
        raise

    
def cargar_multiples_archivos(lista_archivos):
    """Carga múltiples archivos manteniendo rolling de 12 meses."""
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
    cursor = conn.cursor()
    
    # Reemplazar meses existentes que se están cargando de nuevo
    meses_unicos = sorted(set(meses_subidos), key=lambda mes: _mes_a_orden(mes))
    if meses_unicos:
        placeholders = ",".join(["?"] * len(meses_unicos))
        cursor.execute(
            f"DELETE FROM registros WHERE mes_archivo IN ({placeholders})",
            tuple(meses_unicos)
        )
    
    df_total.to_sql("registros", conn, if_exists="append", index=False)
    
    # Rolling: mantener solo 12 meses
    meses_df = pd.read_sql_query(
        "SELECT DISTINCT mes_archivo FROM registros ORDER BY CAST(SUBSTR(mes_archivo, 3, 4) AS INTEGER), CAST(SUBSTR(mes_archivo, 1, 2) AS INTEGER)",
        conn
    )
    meses_lista = meses_df["mes_archivo"].tolist()
    eliminados = []
    while len(meses_lista) > MESES_A_MANTENER:
        mes_viejo = meses_lista.pop(0)
        cursor.execute("DELETE FROM registros WHERE mes_archivo = ?", (mes_viejo,))
        eliminados.append(mes_viejo)
    
    conn.commit()
    conn.close()
    return True, f"Procesados {len(todos_dfs)} archivos.", meses_subidos, eliminados

# ======= FUNCIONES PARA DUPLICADOS REPORTADOS =======
def registrar_duplicados_reportados(duplicados_df, mes_archivo):
    """Registra duplicados reportados usando INSERT OR IGNORE."""
    if duplicados_df.empty:
        return
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    insert_query = """
        INSERT OR IGNORE INTO duplicados_reportados (
            mes_archivo, fecha_emision, tipo_comprobante, serie_comprobante,
            numero_comprobante, ruc_proveedor, razon_social, base_imponible,
            igv, importe_total
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    
    for _, row in duplicados_df.iterrows():
        try:
            cursor.execute(insert_query, (
                mes_archivo,
                str(row.get('fecha_emision', '')) if pd.notna(row.get('fecha_emision', '')) else '',
                str(row.get('tipo_comprobante', '')) if pd.notna(row.get('tipo_comprobante', '')) else '',
                str(row.get('serie_comprobante', '')) if pd.notna(row.get('serie_comprobante', '')) else '',
                str(row.get('numero_comprobante', '')) if pd.notna(row.get('numero_comprobante', '')) else '',
                str(row.get('ruc_proveedor', '')) if pd.notna(row.get('ruc_proveedor', '')) else '',
                str(row.get('razon_social', '')) if pd.notna(row.get('razon_social', '')) else '',
                str(row.get('base_imponible', '')) if pd.notna(row.get('base_imponible', '')) else '',
                str(row.get('igv', '')) if pd.notna(row.get('igv', '')) else '',
                str(row.get('importe_total', '')) if pd.notna(row.get('importe_total', '')) else ''
            ))
        except Exception as e:
            print(f"Error insertando registro: {e}")
            continue
    
    conn.commit()
    conn.close()

def obtener_duplicados_reportados(mes_archivo=None):
    conn = sqlite3.connect(DB_PATH)
    query = "SELECT * FROM duplicados_reportados"
    if mes_archivo:
        query += f" WHERE mes_archivo = '{mes_archivo}'"
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def marcar_duplicados_reportados(duplicados_df, mes_archivo):
    if duplicados_df.empty:
        return duplicados_df
    
    reportes = obtener_duplicados_reportados()
    if reportes.empty:
        duplicados_df['ya_reportado'] = False
        duplicados_df['meses_previos'] = ''
        return duplicados_df
    
    columnas_clave = ['fecha_emision', 'tipo_comprobante', 'serie_comprobante', 
                     'numero_comprobante', 'ruc_proveedor', 'razon_social',
                     'base_imponible', 'igv', 'importe_total']
    
    for col in columnas_clave:
        if col in duplicados_df.columns:
            duplicados_df[col] = duplicados_df[col].astype(str)
        if col in reportes.columns:
            reportes[col] = reportes[col].astype(str)
    
    duplicados_df['_clave_duplicado'] = duplicados_df[columnas_clave].apply(lambda row: '|'.join(row), axis=1)
    reportes['_clave_duplicado'] = reportes[columnas_clave].apply(lambda row: '|'.join(row), axis=1)
    meses_por_clave = reportes.groupby('_clave_duplicado')['mes_archivo'].apply(list).to_dict()
    
    duplicados_df['ya_reportado'] = duplicados_df['_clave_duplicado'].isin(reportes['_clave_duplicado'])
    duplicados_df['meses_previos'] = duplicados_df['_clave_duplicado'].map(
        lambda x: ', '.join(sorted(meses_por_clave.get(x, []))) if x in meses_por_clave else ''
    )
    
    duplicados_df.drop(columns=['_clave_duplicado'], inplace=True)
    return duplicados_df