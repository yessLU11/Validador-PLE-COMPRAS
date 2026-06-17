# ============================================================================
# duplicate_detector_internal.py
# ============================================================================
# Descripción: Detecta duplicados DENTRO DE UN MISMO ARCHIVO EXCEL.
# Compara registros basándose en columnas clave (F, G, H, I, L, Y).
# Retorna DataFrame con duplicados encontrados + auditoría.
# ============================================================================

import pandas as pd
from typing import Tuple, Dict, List
import re


def letra_a_index(letra: str) -> int:
    """Convierte letra de columna Excel (A, B, C, ..., Y, Z, AA, etc) a índice (0-based)"""
    # Esta parte hace que 'A' -> 0, 'B' -> 1, ..., 'Z' -> 25, 'AA' -> 26, etc. para convertir letras de columnas a índices numéricos. 
    # esto nos srive para mapear las letras de columnas a índices numéricos que pandas puede usar para acceder a las columnas por posición.
    letra = letra.upper()
    idx = 0
    for c in letra:
        idx = idx * 26 + (ord(c) - ord('A') + 1)
    return idx - 1


def normalizar_valor(valor) -> str:
    """Limpia espacios, convierte a string, mayúsculas para texto."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, (int, float)):
        return str(valor).strip()
    if isinstance(valor, pd.Timestamp):
        return valor.strftime("%Y-%m-%d")
    return str(valor).strip().upper()


def _make_unique_headers(headers):
    """Convierte una lista de nombres de columna potencialmente duplicados en nombres únicos internos."""
    seen = {}
    unique_headers = []
    for idx, header in enumerate(headers):
        if pd.isna(header):
            header = f"Unnamed: {idx}"
        else:
            header = str(header)
        if header in seen:
            seen[header] += 1
            header = f"{header}__dup_{seen[header]}"
        else:
            seen[header] = 0
        unique_headers.append(header)
    return unique_headers


def leer_excel_todas_hojas(archivo_excel: str) -> pd.DataFrame:
    """
    Lee TODAS las hojas del Excel.
    - Primera hoja: desde fila 8 (FILA_INICIO_PRINCIPAL)
    - Resto de hojas: desde fila 1
    
    Retorna un DataFrame unificado con columna 'hoja_origen' y 'numero_fila_original'.
    """
    dfs = []
    
    # Leer todos los nombres de hojas
    excel_file = pd.ExcelFile(archivo_excel)
    hojas = excel_file.sheet_names
    
    # Mapeo de columnas (F=5, G=6, H=7, I=8, L=11, Y=24, más todas las demás)
    columnas_duplicado = {
        #"F": 5,   # Fecha de Vencimiento o Fecha de Pago
        #"G": 6,   # Tipo de Comprobante de Pago o Documento
        "H": 7,   # Serie del comprobante de pago o documento
        "J": 9,   # Número del comprobante de pago o documento
        #"I": 8,   # Año de emisión de la DUA o DSI
        #"L": 11,  # Tipo de Documento de Identidad del proveedor
        "M": 12,  # RUC del proveedor
        "Y": 24   # Importe total de las adquisiciones
        
    }
    
    for idx_hoja, nombre_hoja in enumerate(hojas):
        try:
            # Determinar fila inicio según sea primera o no
            fila_inicio = 7 if idx_hoja == 0 else 0  # 7 = fila 8 (0-indexed), 0 = fila 1
            
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja, header=None)
            
            #Solo la primera hoja define los nombres de columnas, las demás se leen sin nombres (header=None) y se les asignan nombres únicos internos
            # Solo la primera hoja define los nombres de columna
            if idx_hoja == 0:
                if df.shape[0] > 1:
                    header_row = df.iloc[1].tolist()   # Guardamos los nombres
                else:
                    header_row = None
            else:
                # Para las demás hojas, usar el mismo header_row de la primera
                header_row = header_row if 'header_row' in locals() else None

            df = df.iloc[fila_inicio:, :].copy()
            if header_row is not None:
                # Asegurar que el número de columnas coincida (rellenar con nombres genéricos si faltan)
                num_cols_df = df.shape[1]
                if len(header_row) < num_cols_df:
                    header_row = header_row + [f"Col_{i}" for i in range(len(header_row), num_cols_df)]
                elif len(header_row) > num_cols_df:
                    header_row = header_row[:num_cols_df]
                df.columns = _make_unique_headers(header_row)

            if df.empty or df.shape[0] == 0:
                continue
            
            # Normalizar columnas clave de duplicado
            for col_letra, col_idx in columnas_duplicado.items():
                if col_idx < df.shape[1]:
                    df[f"_norm_{col_letra}"] = df.iloc[:, col_idx].apply(normalizar_valor)
            
            # Guardar referencias a filas originales (número de fila en el Excel original)
            df["hoja_origen"] = nombre_hoja
            df["numero_fila_original"] = range(fila_inicio + 1, fila_inicio + len(df) + 1)  # +1 porque Excel es 1-indexed
            
            # Guardar todas las columnas del Excel original (para el reporte final)
            dfs.append(df)
            
        except Exception as e:
            print(f"Error leyendo hoja '{nombre_hoja}': {e}")
    
    if not dfs:
        return pd.DataFrame()
    
    # Concatenar todas las hojas
    df_final = pd.concat(dfs, ignore_index=True)
    return df_final


def detectar_duplicados_internos(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict]:
    """
    Detecta duplicados DENTRO del DataFrame.
    
    Compara usando columnas: H, J, M, Y (normalizadas).
    Retorna:
    - DataFrame con todas las filas duplicadas (incluye TODAS las columnas originales)
    - Dict con auditoría (cantidad de duplicados, por hoja, etc.)
    """
    
    if df.empty:
        return pd.DataFrame(), {}
    
    # Crear "clave de duplicado" con las columnas normalizadas
    columnas_clave_norm = ["_norm_H", "_norm_J", "_norm_M", "_norm_Y"]
    
    # Verificar que existan las columnas normalizadas
    columnas_existentes = [col for col in columnas_clave_norm if col in df.columns]
    if not columnas_existentes:
        return pd.DataFrame(), {"error": "No se pudieron leer las columnas clave de duplicado"}
    
    # Crear serie de claves de duplicado
    df["_clave_duplicado"] = df[columnas_existentes].apply(
        lambda row: "|".join(row.astype(str)), axis=1
    )
    
    # Encontrar cuáles se repiten (value_counts > 1)
    duplicados_mask = df["_clave_duplicado"].duplicated(keep=False)
    df_duplicados = df[duplicados_mask].copy()
    
    # Ordenar por clave para agrupar duplicados juntos
    df_duplicados = df_duplicados.sort_values("_clave_duplicado").reset_index(drop=True)
    
    # Crear auditoría
    auditoria = {
        "total_filas": len(df),
        "total_duplicados": len(df_duplicados),
        "grupos_duplicados": len(df_duplicados["_clave_duplicado"].unique()) if not df_duplicados.empty else 0,
    }
    
    # Contar por hoja
    if not df_duplicados.empty:
        auditoria["duplicados_por_hoja"] = df_duplicados["hoja_origen"].value_counts().to_dict()
    
    return df_duplicados, auditoria


def obtener_todas_columnas_originales(df: pd.DataFrame, df_duplicados: pd.DataFrame) -> pd.DataFrame:
    """
    Extrae TODAS las columnas originales del Excel para los duplicados.
    """
    if df_duplicados.empty:
        return pd.DataFrame()
    
    # Obtener los índices originales de las filas duplicadas
    indices_duplicados = df_duplicados.index.tolist()
    
    # Crear nuevo DataFrame solo con columnas de datos originales (no las normalizadas ni de auditoría)
    columnas_no_originales = [col for col in df.columns if col.startswith("_") or col in ["hoja_origen", "numero_fila_original"]]
    columnas_originales = [col for col in df.columns if col not in columnas_no_originales]
    
    # Si no hay columnas originales (solo hay las nuevas), usar todas excepto las marcadas
    if not columnas_originales:
        columnas_originales = [col for col in df.columns if not col.startswith("_")]
    
    result = df_duplicados.copy()
    result = result[columnas_originales + ["hoja_origen", "numero_fila_original", "_clave_duplicado"]]
    
    return result


def generar_auditoria_duplicados(df_duplicados: pd.DataFrame, auditoria: Dict) -> str:
    """
    Genera un texto de auditoría detallado sobre los duplicados encontrados.
    """
    texto_auditoria = "=" * 80 + "\n"
    texto_auditoria += "AUDITORÍA DE DUPLICADOS - VALIDACIÓN INTERNA DE EXCEL\n"
    texto_auditoria += "=" * 80 + "\n\n"
    
    texto_auditoria += f"Total de filas leídas: {auditoria.get('total_filas', 0)}\n"
    texto_auditoria += f"Total de DUPLICADOS encontrados: {auditoria.get('total_duplicados', 0)}\n"
    texto_auditoria += f"Grupos únicos de duplicados: {auditoria.get('grupos_duplicados', 0)}\n\n"
    
    if auditoria.get('duplicados_por_hoja'):
        texto_auditoria += "Duplicados por hoja:\n"
        for hoja, cantidad in auditoria['duplicados_por_hoja'].items():
            texto_auditoria += f"  - {hoja}: {cantidad} registros\n"
    
    texto_auditoria += "\n" + "=" * 80 + "\n"
    
    if not df_duplicados.empty:
        # Agrupar por clave de duplicado
        texto_auditoria += "\nDETALLE DE GRUPOS DE DUPLICADOS:\n"
        texto_auditoria += "-" * 80 + "\n"
        
        grupos = df_duplicados.groupby("_clave_duplicado")
        for idx, (clave, grupo) in enumerate(grupos, 1):
            texto_auditoria += f"\nGrupo {idx} (repeticiones: {len(grupo)}):\n"
            
            for _, row in grupo.iterrows():
                hoja = row.get("hoja_origen", "N/A")
                fila = row.get("numero_fila_original", "N/A")
                texto_auditoria += f"  - Hoja: '{hoja}', Fila: {fila}\n"
    
    return texto_auditoria
