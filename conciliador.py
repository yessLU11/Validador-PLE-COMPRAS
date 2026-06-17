#conciliador.py
# Este archivo contiene toda la lógica de negocio:
# Lectura, limpieza, creación de ID, comparación, generación del reporte Excel con formato profesional.

# Módulo para conciliación de dos archivos PLE Compras (comparación de series y números)
# Autor: Yessly Poma
# Fecha: 2026-06-16

import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import time

# ------------------------------------------------------------
# 1. LECTURA Y NORMALIZACIÓN DE UN ARCHIVO
# ------------------------------------------------------------

def leer_archivo_conciliacion(ruta_archivo, hoja=None, fila_inicio=1):
    """
    Lee un archivo Excel y extrae las columnas H (Serie) y J (Número) por posición fija.
    Parámetros:
        ruta_archivo: str
        hoja: str o None (si None, toma la primera hoja)
        fila_inicio: int (1‑based) -> fila donde comienzan los datos (ej. 5 si la fila 5 es la primera de datos)
    Retorna DataFrame con columnas originales + Serie, Numero, ID_CONCILIACION, Hoja_Origen.
    """
    # skiprows = fila_inicio - 1 (porque pandas usa 0‑based)
    skip_rows = fila_inicio - 1

    if hoja is None:
        df = pd.read_excel(ruta_archivo, header=None, skiprows=skip_rows)
    else:
        df = pd.read_excel(ruta_archivo, sheet_name=hoja, header=None, skiprows=skip_rows)

    if df.empty:
        raise ValueError("No se encontraron datos en el archivo u hoja especificada.")

    if df.shape[1] < 10:
        raise ValueError("El archivo no tiene suficientes columnas (se necesitan H y J).")

    df_proc = df.copy()
    serie_raw = df_proc.iloc[:, 7].astype(str).str.strip()
    numero_raw = df_proc.iloc[:, 9].astype(str).str.strip()
    numero_clean = numero_raw.str.replace(r'\.0$', '', regex=True)
    numero_fill = numero_clean.str.zfill(8)

    df_proc['Serie'] = serie_raw
    df_proc['Numero'] = numero_fill
    df_proc['ID_CONCILIACION'] = serie_raw + numero_fill
    df_proc['Hoja_Origen'] = hoja if hoja else 'Hoja1'

    return df_proc

def leer_todas_hojas_conciliacion(ruta_archivo, fila_inicio=1):
    """
    Lee todas las hojas del archivo, aplicando el mismo fila_inicio a cada una.
    """
    xlsx = pd.ExcelFile(ruta_archivo)
    hojas = xlsx.sheet_names
    dfs = []
    for hoja in hojas:
        try:
            df_hoja = leer_archivo_conciliacion(ruta_archivo, hoja=hoja, fila_inicio=fila_inicio)
            dfs.append(df_hoja)
        except Exception as e:
            print(f"Error en hoja {hoja}: {e}")
            continue
    if not dfs:
        raise ValueError("No se pudo leer ninguna hoja válida.")
    return pd.concat(dfs, ignore_index=True)

# ------------------------------------------------------------
# 2. COMPARACIÓN DE DOS ARCHIVOS
# ------------------------------------------------------------

def conciliar_archivos(df1, df2, nombre1="Archivo 1", nombre2="Archivo 2"):
    """
    Compara dos DataFrames (ya procesados con ID_CONCILIACION) y devuelve:
        - resumen: dict con estadísticas
        - solo1: DataFrame con filas de df1 cuyo ID no está en df2
        - solo2: DataFrame con filas de df2 cuyo ID no está en df1
    """
    ids1 = set(df1['ID_CONCILIACION'].dropna().unique())
    ids2 = set(df2['ID_CONCILIACION'].dropna().unique())

    comunes = ids1 & ids2
    solo1_ids = ids1 - ids2
    solo2_ids = ids2 - ids1

    solo1 = df1[df1['ID_CONCILIACION'].isin(solo1_ids)].copy()
    solo2 = df2[df2['ID_CONCILIACION'].isin(solo2_ids)].copy()

    resumen = {
        'nombre_archivo_1': nombre1,
        'nombre_archivo_2': nombre2,
        'total_registros_1': len(df1),
        'total_registros_2': len(df2),
        'ids_unicos_1': len(ids1),
        'ids_unicos_2': len(ids2),
        'solo_en_1': len(solo1),
        'solo_en_2': len(solo2),
        'comunes': len(comunes),
        'diferencias_totales': len(solo1) + len(solo2)
    }

    return resumen, solo1, solo2


def contar_por_tipo_documento(df1, df2, nombre1, nombre2):
    """
    Cuenta los registros por tipo de documento (columna G, índice 6)
    en ambos DataFrames y genera un resumen comparativo.
    
    Retorna un DataFrame con:
        - Tipo de documento (descripción)
        - Total en archivo 1
        - Total en archivo 2
        - Diferencia (positiva)
        - Libro electrónico (cuál tiene más)
    """
    # Mapeo de códigos a descripciones
    TIPOS_DOC = {
        '01': 'FACTURAS',
        '03': 'BOLETAS',
        '05': 'PASAJES AEREOS',
        '12': 'TICKET MAQUINA REGISTRADORA',
        '14': 'RECIBOS DE SERVICIO PUBLICO',
        '15': 'TRANSPORTE TERRESTRE',
        '16': 'TRANSPORTE DE CARGA',
        '30': 'DOCUMENTOS AUTORIZADOS',
        '46': 'NO DOMICILIADO'
    }
    
    # Extraer tipos de la columna G (índice 6) de cada DataFrame
    # Asumimos que los DataFrames mantienen las columnas originales
    # La columna G es la posición 6 (0-based)
    
    def obtener_tipos(df):
        """Extrae los tipos de la columna G (índice 6)"""
        if df.shape[1] > 6:
            tipos = df.iloc[:, 6].astype(str).str.strip()
            # Reemplazar valores vacíos o nulos con 'DESCONOCIDO'
            tipos = tipos.replace('', 'DESCONOCIDO')
            tipos = tipos.fillna('DESCONOCIDO')
            return tipos
        else:
            return pd.Series(['SIN_COLUMNA'] * len(df))
    
    tipos1 = obtener_tipos(df1)
    tipos2 = obtener_tipos(df2)
    
    # Contar frecuencias
    conteo1 = tipos1.value_counts().to_dict()
    conteo2 = tipos2.value_counts().to_dict()
    
    # Obtener todos los tipos únicos (de ambos archivos)
    todos_tipos = set(conteo1.keys()) | set(conteo2.keys())
    
    # Construir DataFrame de resultados
    resultados = []
    for tipo in sorted(todos_tipos):
        total1 = conteo1.get(tipo, 0)
        total2 = conteo2.get(tipo, 0)
        diferencia = abs(total1 - total2)
        
        # Determinar qué archivo tiene más
        if total1 > total2:
            libro_electronico = nombre1
        elif total2 > total1:
            libro_electronico = nombre2
        else:
            libro_electronico = 'IGUALES'
        
        # Obtener descripción del tipo
        descripcion = TIPOS_DOC.get(tipo, tipo)
        
        resultados.append({
            'Tipo de Documento': descripcion,
            f'Total {nombre1}': total1,
            f'Total {nombre2}': total2,
            'Diferencia': diferencia,
            'Libro Electrónico': libro_electronico
        })
    
    # Ordenar por código de tipo (numéricamente)
    # Para orden correcto: 01, 03, 05, 12, 14, 15, 16, 30, 46
    orden_tipos = ['01', '03', '05', '12', '14', '15', '16', '30', '46', 'DESCONOCIDO', 'SIN_COLUMNA']
    df_resultado = pd.DataFrame(resultados)
    
    # Asignar orden
    df_resultado['_orden'] = df_resultado['Tipo de Documento'].apply(
        lambda x: orden_tipos.index(x) if x in orden_tipos else 999
    )
    df_resultado = df_resultado.sort_values('_orden').drop('_orden', axis=1)
    
    # Agregar fila de TOTAL
    total_fila = {
        'Tipo de Documento': 'TOTAL GENERAL',
        f'Total {nombre1}': sum(r[f'Total {nombre1}'] for r in resultados),
        f'Total {nombre2}': sum(r[f'Total {nombre2}'] for r in resultados),
        'Diferencia': abs(sum(r[f'Total {nombre1}'] for r in resultados) - 
                          sum(r[f'Total {nombre2}'] for r in resultados)),
        'Libro Electrónico': nombre1 if sum(r[f'Total {nombre1}'] for r in resultados) > 
                                   sum(r[f'Total {nombre2}'] for r in resultados) else nombre2
    }
    df_resultado = pd.concat([df_resultado, pd.DataFrame([total_fila])], ignore_index=True)
    
    return df_resultado

# ------------------------------------------------------------
# 3. GENERACIÓN DE REPORTE EXCEL
# ------------------------------------------------------------
def generar_reporte_conciliacion(resumen, solo1, solo2, df1, df2, nombre1, nombre2, ruta_salida):
    """
    Crea un archivo Excel con tres hojas:
        1. Resumen (dos tablas: resumen general + desglose por tipo)
        2. Solo en Archivo_1
        3. Solo en Archivo_2
    """
    wb = Workbook()
    wb.remove(wb.active)
    
    # ---------- Hoja 1: Resumen ----------
    ws1 = wb.create_sheet("Resumen")
    
    # ===== TABLA 1: Resumen General =====
    # Título
    ws1['A1'] = 'RESUMEN GENERAL DE CONCILIACIÓN'
    ws1.merge_cells('A1:C1')
    ws1['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws1['A1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')
    
    # Encabezados
    headers_resumen = ['Concepto', nombre1, nombre2]
    ws1.append(headers_resumen)
    
    datos_resumen = [
        ['Total registros (serie)', resumen['total_registros_1'], resumen['total_registros_2']],
        ['IDs únicos', resumen['ids_unicos_1'], resumen['ids_unicos_2']],
        ['Registros solo en este archivo', resumen['solo_en_1'], resumen['solo_en_2']],
        ['Registros en común', resumen['comunes'], resumen['comunes']],
        ['Diferencias totales', resumen['diferencias_totales'], resumen['diferencias_totales']]
    ]
    for fila in datos_resumen:
        ws1.append(fila)
    
    # Formato para la primera tabla
    _aplicar_formato_encabezado(ws1, row=2)  # Fila 2 contiene los encabezados
    
    # Ajustar ancho de columnas
    for col in ['A', 'B', 'C']:
        ws1.column_dimensions[col].width = 30
    
    # ===== Espacio entre tablas =====
    fila_actual = len(datos_resumen) + 4  # Dejar 2 filas de espacio
    
    # ===== TABLA 2: Desglose por Tipo de Documento =====
    # Título
    ws1.cell(row=fila_actual, column=1, value='DESGLOSE POR TIPO DE DOCUMENTO')
    ws1.merge_cells(start_row=fila_actual, start_column=1, end_row=fila_actual, end_column=5)
    ws1.cell(row=fila_actual, column=1).font = Font(bold=True, size=14, color="FFFFFF")
    ws1.cell(row=fila_actual, column=1).fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws1.cell(row=fila_actual, column=1).alignment = Alignment(horizontal='center', vertical='center')
    
    fila_actual += 1
    
    # Obtener desglose por tipo
    df_tipos = contar_por_tipo_documento(df1, df2, nombre1, nombre2)
    
    # Escribir encabezados
    headers_tipos = list(df_tipos.columns)
    for col_idx, header in enumerate(headers_tipos, 1):
        ws1.cell(row=fila_actual, column=col_idx, value=header)
    
    fila_actual += 1
    
    # Escribir datos
    for _, row in df_tipos.iterrows():
        for col_idx, header in enumerate(headers_tipos, 1):
            ws1.cell(row=fila_actual, column=col_idx, value=row[header])
        fila_actual += 1
    
    # Aplicar formato a encabezados de la segunda tabla
    _aplicar_formato_encabezado(ws1, row=fila_actual - len(df_tipos) - 1)
    
    # Ajustar ancho de columnas para la segunda tabla
    for col_idx in range(1, len(headers_tipos) + 1):
        col_letter = get_column_letter(col_idx)
        ws1.column_dimensions[col_letter].width = 25
    
    # Aplicar formato de número a las columnas de totales
    for row in range(fila_actual - len(df_tipos), fila_actual):
        # Formato para columnas numéricas (2, 3, 4)
        for col_idx in [2, 3, 4]:
            if col_idx <= len(headers_tipos):
                cell = ws1.cell(row=row, column=col_idx)
                cell.number_format = '#,##0'
    
    # ---------- Hoja 2: Solo en Archivo 1 ----------
    if not solo1.empty:
        ws2 = wb.create_sheet("Solo en Archivo_1")
        columnas_orden = ['ID_CONCILIACION', 'Serie', 'Numero', 'Hoja_Origen'] + \
                         [c for c in solo1.columns if c not in ['ID_CONCILIACION', 'Serie', 'Numero', 'Hoja_Origen']]
        df_export = solo1[columnas_orden]
        _escribir_dataframe_con_formato(ws2, df_export)
    else:
        ws2 = wb.create_sheet("Solo en Archivo_1")
        ws2.append(["No hay registros solo en este archivo"])

    # ---------- Hoja 3: Solo en Archivo 2 ----------
    if not solo2.empty:
        ws3 = wb.create_sheet("Solo en Archivo_2")
        columnas_orden = ['ID_CONCILIACION', 'Serie', 'Numero', 'Hoja_Origen'] + \
                         [c for c in solo2.columns if c not in ['ID_CONCILIACION', 'Serie', 'Numero', 'Hoja_Origen']]
        df_export = solo2[columnas_orden]
        _escribir_dataframe_con_formato(ws3, df_export)
    else:
        ws3 = wb.create_sheet("Solo en Archivo_2")
        ws3.append(["No hay registros solo en este archivo"])

    wb.save(ruta_salida)
    return ruta_salida

# ------------------------------------------------------------
# 4. FUNCIONES AUXILIARES DE FORMATO
# ------------------------------------------------------------

def _aplicar_formato_encabezado(ws, row=1):
    """Aplica formato de encabezado (negrita, azul, centrado) a una fila específica."""
    fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    for cell in ws[row]:
        if cell.value:
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border


def _escribir_dataframe_con_formato(ws, df):
    """
    Escribe un DataFrame en una hoja de cálculo y aplica formato a los encabezados.
    """
    # Encabezados
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name)

    # Datos
    for row_idx, row_data in enumerate(df.values, 2):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    _aplicar_formato_encabezado(ws, row=1)

    # Ajustar ancho de columnas (máx 50)
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
        adjusted_width = min(max_len + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted_width