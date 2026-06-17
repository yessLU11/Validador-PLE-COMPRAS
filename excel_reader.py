# Lectura y normalización de los archivos PLE
# excel_reader.py
# Este módulo se encarga de leer los archivos Excel, extraer las columnas relevantes,
# normalizar los datos (limpiar espacios, convertir a mayúsculas, formatear fechas), y unificar las hojas "8.1" y "PROGRAMAS SOCIALES" en un solo DataFrame.
# Se relaciona con database.py para cargar los datos a la base de datos, y con main.py para ser llamado desde el flujo principal de carga de archivos.
# ============================================================================
import pandas as pd
import re
from config import *

def normalizar_valor(valor):
    """Limpia espacios, convierte a string, mayúsculas para texto."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, (int, float)):
        # Para números (RUC, montos) lo dejamos como está, pero como string
        return str(valor).strip()
    # Para fechas (pandas Timestamp) la convertimos a YYYY-MM-DD
    if isinstance(valor, pd.Timestamp):
        return valor.strftime("%Y-%m-%d")
    # texto normal
    return str(valor).strip().upper()

def leer_hoja_unificada(archivo_excel, mes_archivo):
    """
    Lee ambas hojas (8.1 y PROGRAMAS SOCIALES), las unifica en un solo DataFrame
    añadiendo una columna 'hoja_origen' y 'mes_archivo' (YYYYMM).
    """
    dfs = []
    
    # Función auxiliar para convertir letra de columna a índice
    def letra_a_index(letra):
        letra = letra.upper()
        idx = 0
        for c in letra:
            idx = idx * 26 + (ord(c) - ord('A') + 1)
        return idx - 1
    
    # Obtener nombres de hojas disponibles UNA SOLA VEZ
    xlsx = pd.ExcelFile(archivo_excel)
    hojas_disponibles = xlsx.sheet_names
    
    # ===== LECTURA DE HOJA PRINCIPAL =====
    if HOJA_PRINCIPAL in hojas_disponibles:
        try:
            df_ppal = pd.read_excel(archivo_excel, sheet_name=HOJA_PRINCIPAL, header=None)
            # Tomamos desde la fila FILA_INICIO_PRINCIPAL (8) en adelante
            df_ppal = df_ppal.iloc[FILA_INICIO_PRINCIPAL-1:, :].copy()  # .copy() para evitar warnings
            
            columnas_interes = {col_letra: letra_a_index(col_letra) for col_letra in MAPEO_COLUMNAS.keys()}
            # Extraer solo esas columnas
            df_ppal_sub = df_ppal.iloc[:, list(columnas_interes.values())].copy()
            # Renombrar con nuestros nombres
            df_ppal_sub.columns = list(MAPEO_COLUMNAS.values())
            # Asignar hoja de origen
            df_ppal_sub["hoja_origen"] = HOJA_PRINCIPAL
            dfs.append(df_ppal_sub)
        except Exception as e:
            print(f"Error leyendo hoja {HOJA_PRINCIPAL}: {e}")
    else:
        print(f"Hoja '{HOJA_PRINCIPAL}' no encontrada. Verifica el archivo.")
    
    # ===== LECTURA DE HOJA DE PROGRAMAS SOCIALES =====
    if HOJA_SOCIALES in hojas_disponibles:
        try:
            df_social = pd.read_excel(archivo_excel, sheet_name=HOJA_SOCIALES, header=None)
            # Tomamos desde la fila FILA_INICIO_SOCIALES (1) en adelante
            df_social = df_social.iloc[FILA_INICIO_SOCIALES-1:, :].copy()
            
            columnas_interes = {col_letra: letra_a_index(col_letra) for col_letra in MAPEO_COLUMNAS.keys()}
            # Extraer solo esas columnas
            df_social_sub = df_social.iloc[:, list(columnas_interes.values())].copy()
            # Renombrar con nuestros nombres
            df_social_sub.columns = list(MAPEO_COLUMNAS.values())
            # Asignar hoja de origen
            df_social_sub["hoja_origen"] = HOJA_SOCIALES
            dfs.append(df_social_sub)
        except Exception as e:
            print(f"Error procesando hoja {HOJA_SOCIALES}: {e}")
    else:
        print(f"Hoja '{HOJA_SOCIALES}' no encontrada. Se omite.")
    
    # ===== UNIFICAR DATAFRAMES =====
    if not dfs:
        print("No se encontraron hojas válidas para procesar.")
        return pd.DataFrame()
    
    df_final = pd.concat(dfs, ignore_index=True)
    
    # Limpiar valores nulos: eliminar filas donde todas las columnas clave estén vacías
    df_final = df_final.dropna(how='all', subset=COLUMNAS_CLAVE)
    
    # Aplicar normalización a cada columna clave
    for col in COLUMNAS_CLAVE:
        if col in df_final.columns:
            df_final[col] = df_final[col].apply(normalizar_valor)
    
    # Agregar columna del mes (formato YYYYMM)
    df_final["mes_archivo"] = mes_archivo
    
    return df_final