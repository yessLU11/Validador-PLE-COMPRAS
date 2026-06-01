# Lectura y normalización de los archivos PLE
# excel_reader.py
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
    # Leer hoja principal
    try:
        df_ppal = pd.read_excel(archivo_excel, sheet_name=HOJA_PRINCIPAL, header=None)
        # Tomamos desde la fila FILA_INICIO_PRINCIPAL (8) en adelante
        df_ppal = df_ppal.iloc[FILA_INICIO_PRINCIPAL-1:, :]  # -1 porque pandas es 0-index
        # Seleccionamos solo las columnas que necesitamos (A=0, E=4, G=6, H=7, J=9, M=12, N=13, Q=16, R=17, Y=24)
        # Pero es mejor por nombre de letra: convertimos letra a índice
        def letra_a_index(letra):
            letra = letra.upper()
            idx = 0
            for c in letra:
                idx = idx * 26 + (ord(c) - ord('A') + 1)
            return idx - 1
        
        columnas_interes = {col_letra: letra_a_index(col_letra) for col_letra in MAPEO_COLUMNAS.keys()}
        # Extraer solo esas columnas
        df_ppal_sub = df_ppal.iloc[:, list(columnas_interes.values())]
        # Renombrar con nuestros nombres
        df_ppal_sub.columns = list(MAPEO_COLUMNAS.values())
        df_ppal_sub["hoja_origen"] = HOJA_PRINCIPAL
        dfs.append(df_ppal_sub)
    except Exception as e:
        print(f"Error leyendo hoja {HOJA_PRINCIPAL}: {e}")
    
    # Leer hoja de programas sociales
    try:
        df_social = pd.read_excel(archivo_excel, sheet_name=HOJA_SOCIALES, header=None)
        df_social = df_social.iloc[FILA_INICIO_SOCIALES-1:, :]
        df_social_sub = df_social.iloc[:, list(columnas_interes.values())]
        df_social_sub.columns = list(MAPEO_COLUMNAS.values())
        df_social_sub["hoja_origen"] = HOJA_SOCIALES
        dfs.append(df_social_sub)
    except Exception as e:
        print(f"Error leyendo hoja {HOJA_SOCIALES}: {e}")
    
    if not dfs:
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