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
# ------------------------------------------------------------
# 0. NORMALIZACIÓN DE TIPO DE DOCUMENTO (COLUMNA G)
# ------------------------------------------------------------

def normalizar_tipo_documento(valor):
    """
    Normaliza el tipo de documento de la columna G (índice 6).
    - Convierte a string
    - Elimina decimales (42.0 -> 42)
    - Rellena con cero a la izquierda para tener 2 dígitos
    - Maneja casos especiales (NaN, vacío, etc.)
    
    Ejemplos:
        1   -> '01'
        03  -> '03'
        42.0 -> '42'
        0.0 -> '00'
        NaN -> '00'
        ''  -> '00'
    """
    # Si es NaN o vacío, retornar '00'
    if pd.isna(valor) or str(valor).strip() == '':
        return '00'
    
    # Convertir a string y limpiar
    valor_str = str(valor).strip()
    
    # Si es un número con decimales (ej: 42.0), convertir a entero
    if '.' in valor_str:
        try:
            valor_str = str(int(float(valor_str)))
        except:
            pass
    
    # Intentar extraer solo dígitos (para casos como '03' o '3')
    import re
    digitos = re.search(r'(\d+)', valor_str)
    if digitos:
        valor_str = digitos.group(1)
    else:
        return '00'
    
    # Rellenar con cero a la izquierda hasta 2 dígitos
    return valor_str.zfill(2)

def leer_archivo_conciliacion(ruta_archivo, hoja=None, fila_inicio=1):
    """
    Lee un archivo Excel y extrae las columnas H (Serie) y J (Número) por posición fija.
    También extrae la columna G (Tipo de Documento) y la normaliza.
    """
    skip_rows = fila_inicio - 1

    if hoja is None:
        df = pd.read_excel(ruta_archivo, header=None, skiprows=skip_rows)
    else:
        df = pd.read_excel(ruta_archivo, sheet_name=hoja, header=None, skiprows=skip_rows)

    if df.empty:
        raise ValueError("No se encontraron datos en el archivo u hoja especificada.")

    if df.shape[1] < 10:
        raise ValueError("El archivo no tiene suficientes columnas (se necesitan G, H y J).")

    df_proc = df.copy()
    
    # --- Extraer y normalizar Tipo de Documento (columna G, índice 6) ---
    tipo_raw = df_proc.iloc[:, 6] if df_proc.shape[1] > 6 else pd.Series(['00'] * len(df_proc))
    df_proc['TipoDoc_Norm'] = tipo_raw.apply(normalizar_tipo_documento)
    
    # --- Extraer Serie (columna H, índice 7) ---
    serie_raw = df_proc.iloc[:, 7].astype(str).str.strip()
    
    # --- Extraer y normalizar Número (columna J, índice 9) ---
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
    Cuenta los registros por tipo de documento usando la columna normalizada 'TipoDoc_Norm'.
    """
    # Mapeo de códigos a descripciones (todos los tipos posibles)
    TIPOS_DOC = {
        '00': '00',
        '01': '01',
        '02': '02',
        '03': '03',
        '04': '04',
        '05': '05',
        '06': '06',
        '07': '07',
        '08': '08',
        '09': '09',
        '12': '12',
        '14': '14',
        '16': '16',
        '30': '30',
        '42': '42',
        '46': '46',
        '53': '53'
    }
    
    # Usar la columna normalizada 'TipoDoc_Norm'
    if 'TipoDoc_Norm' in df1.columns:
        tipos1 = df1['TipoDoc_Norm']
    else:
        # Fallback: intentar extraer de columna G
        tipos1 = df1.iloc[:, 6].apply(normalizar_tipo_documento) if df1.shape[1] > 6 else pd.Series(['00'] * len(df1))
    
    if 'TipoDoc_Norm' in df2.columns:
        tipos2 = df2['TipoDoc_Norm']
    else:
        tipos2 = df2.iloc[:, 6].apply(normalizar_tipo_documento) if df2.shape[1] > 6 else pd.Series(['00'] * len(df2))
    
    # Contar frecuencias
    conteo1 = tipos1.value_counts().to_dict()
    conteo2 = tipos2.value_counts().to_dict()
    
    # Obtener todos los tipos únicos
    todos_tipos = set(conteo1.keys()) | set(conteo2.keys())
    
    # Ordenar tipos numéricamente
    def orden_tipo(t):
        try:
            return int(t)
        except:
            return 999
    
    resultados = []
    for tipo in sorted(todos_tipos, key=orden_tipo):
        total1 = conteo1.get(tipo, 0)
        total2 = conteo2.get(tipo, 0)
        diferencia = abs(total1 - total2)
        
        if total1 > total2:
            libro_electronico = nombre1
        elif total2 > total1:
            libro_electronico = nombre2
        else:
            libro_electronico = 'IGUALES'
        
        descripcion = TIPOS_DOC.get(tipo, f' {tipo}')
        
        resultados.append({
            'Tipo de Documento': f'{tipo} ',
            f'Total {nombre1}': total1,
            f'Total {nombre2}': total2,
            'Diferencia': diferencia,
            'Libro Electrónico': libro_electronico
        })
    
    # Agregar fila de TOTAL
    total1 = sum(r[f'Total {nombre1}'] for r in resultados)
    total2 = sum(r[f'Total {nombre2}'] for r in resultados)
    total_fila = {
        'Tipo de Documento': 'TOTAL GENERAL',
        f'Total {nombre1}': total1,
        f'Total {nombre2}': total2,
        'Diferencia': abs(total1 - total2),
        'Libro Electrónico': nombre1 if total1 > total2 else (nombre2 if total2 > total1 else 'IGUALES')
    }
    resultados.append(total_fila)
    
    return pd.DataFrame(resultados)
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

# Se modifico para que acepte el parametro start_row  así insertar tablas en cualquier fila 
def _escribir_dataframe_con_formato(ws, df, start_row=1):
    """
    Escribe un DataFrame en una hoja de cálculo a partir de la fila start_row,
    aplicando formato a los encabezados (que estarán en start_row).
    """
    # Encabezados
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=start_row, column=col_idx, value=col_name)

    # Datos
    for row_idx, row_data in enumerate(df.values, start_row + 1):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    # Aplicar formato a la fila de encabezados
    _aplicar_formato_encabezado(ws, row=start_row)

    # Ajustar ancho de columnas (opcional, pero se puede hacer una vez por hoja)
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

# ------------------------------------------------------------
# 5. REPORTE DE PRESENTES VS NO PRESENTES EN SIRE
# ------------------------------------------------------------

# ------------------------------------------------------------
# 5. REPORTE DE PRESENTES VS NO PRESENTES EN SIRE (CORREGIDO)
# ------------------------------------------------------------

# ------------------------------------------------------------
# 5. REPORTE DE PRESENTES VS NO PRESENTES EN SIRE (CORREGIDO)
# ------------------------------------------------------------

def generar_reporte_presentes_no_presentes(df_sire, df_ple, nombre_sire, nombre_ple, ruta_salida):
    """
    Genera un reporte Excel que muestra, por tipo de documento:
    - Cuántos registros del PLE (archivo 2) están presentes en SIRE (archivo 1)
    - Cuántos NO están presentes
    - Detalle de TODAS las filas para cada tipo (incluyendo duplicados)
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    # Verificar que tengan las columnas necesarias
    required_cols = ['ID_CONCILIACION', 'TipoDoc_Norm']
    for col in required_cols:
        if col not in df_sire.columns or col not in df_ple.columns:
            raise ValueError(f"Ambos DataFrames deben tener la columna '{col}'")

    wb = Workbook()
    wb.remove(wb.active)

    # ---------- HOJA 1: Resumen ----------
    ws_resumen = wb.create_sheet("Resumen")

    # Título
    ws_resumen['A1'] = 'RESUMEN: REGISTROS DEL PLE PRESENTES / NO PRESENTES EN SIRE'
    ws_resumen.merge_cells('A1:D1')
    ws_resumen['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws_resumen['A1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws_resumen['A1'].alignment = Alignment(horizontal='center', vertical='center')

    # Encabezados de tabla (ahora usamos "Registros" en lugar de "IDs")
    headers = ['Tipo de Documento', f'Presentes en {nombre_sire}', f'No presentes en {nombre_sire}', 'Total en PLE']
    ws_resumen.append(headers)

    # Obtener todos los tipos de documento del PLE (archivo 2)
    # Reemplazar valores nulos o vacíos con 'DESCONOCIDO'
    df_ple['TipoDoc_Norm'] = df_ple['TipoDoc_Norm'].fillna('DESCONOCIDO')
    df_ple['TipoDoc_Norm'] = df_ple['TipoDoc_Norm'].replace('', 'DESCONOCIDO')
    
    tipos_ple = df_ple['TipoDoc_Norm'].unique()
    # Ordenar: primero los numéricos, luego 'DESCONOCIDO' al final
    tipos_numericos = sorted([t for t in tipos_ple if t != 'DESCONOCIDO' and t.isdigit()], key=lambda x: int(x))
    tipos_ordenados = tipos_numericos + (['DESCONOCIDO'] if 'DESCONOCIDO' in tipos_ple else [])

    # Conjuntos de IDs para comparación (esto sigue siendo necesario para saber qué está presente)
    ids_sire = set(df_sire['ID_CONCILIACION'].dropna())

    fila_actual = 3
    total_presentes = 0
    total_no_presentes = 0
    total_registros_ple = 0

    for tipo in tipos_ordenados:
        # Filtrar registros del PLE para este tipo (TODOS los registros, incluyendo duplicados)
        df_ple_tipo = df_ple[df_ple['TipoDoc_Norm'] == tipo]
        total_registros = len(df_ple_tipo)
        
        # Dividir en presentes y no presentes (basado en ID_CONCILIACION)
        mask_presente = df_ple_tipo['ID_CONCILIACION'].isin(ids_sire)
        presentes_count = mask_presente.sum()
        no_presentes_count = total_registros - presentes_count

        total_presentes += presentes_count
        total_no_presentes += no_presentes_count
        total_registros_ple += total_registros

        # Escribir fila
        ws_resumen.cell(row=fila_actual, column=1, value=tipo)
        ws_resumen.cell(row=fila_actual, column=2, value=presentes_count)
        ws_resumen.cell(row=fila_actual, column=3, value=no_presentes_count)
        ws_resumen.cell(row=fila_actual, column=4, value=total_registros)
        fila_actual += 1

    # Fila de TOTAL
    ws_resumen.cell(row=fila_actual, column=1, value='TOTAL')
    ws_resumen.cell(row=fila_actual, column=2, value=total_presentes)
    ws_resumen.cell(row=fila_actual, column=3, value=total_no_presentes)
    ws_resumen.cell(row=fila_actual, column=4, value=total_registros_ple)

    # Aplicar formato a los encabezados (fila 2)
    _aplicar_formato_encabezado(ws_resumen, row=2)
    
    # Negrita para la fila de total - CORREGIDO: usar índices numéricos
    for col_idx in range(1, 5):  # columnas 1 a 4 (A, B, C, D)
        ws_resumen.cell(row=fila_actual, column=col_idx).font = Font(bold=True)

    # Ajustar ancho de columnas
    for col in ['A', 'B', 'C', 'D']:
        ws_resumen.column_dimensions[col].width = 25

    # ---------- HOJAS POR TIPO DE DOCUMENTO ----------
    for tipo in tipos_ordenados:
        # Crear hoja con nombre "Tipo XX"
        sheet_name = f"Tipo {tipo}"
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]
        ws_tipo = wb.create_sheet(sheet_name)

        # Obtener TODOS los registros del PLE para este tipo
        df_ple_tipo = df_ple[df_ple['TipoDoc_Norm'] == tipo]
        mask_presente = df_ple_tipo['ID_CONCILIACION'].isin(ids_sire)

        # --- Sección: Presentes en SIRE ---
        fila_inicio = 1
        ws_tipo.merge_cells(f'A{fila_inicio}:E{fila_inicio}')
        celda_titulo = ws_tipo.cell(row=fila_inicio, column=1, value=f'✅ PRESENTES EN {nombre_sire} (Total: {mask_presente.sum()})')
        celda_titulo.font = Font(bold=True, size=12, color="FFFFFF")
        celda_titulo.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        celda_titulo.alignment = Alignment(horizontal='center', vertical='center')
        fila_inicio += 1

        df_presentes = df_ple_tipo[mask_presente].copy()
        if not df_presentes.empty:
            cols_orden = ['ID_CONCILIACION', 'Serie', 'Numero', 'Hoja_Origen'] + \
                         [c for c in df_presentes.columns if c not in ['ID_CONCILIACION', 'Serie', 'Numero', 'Hoja_Origen']]
            df_presentes = df_presentes[cols_orden]
            _escribir_dataframe_con_formato(ws_tipo, df_presentes, start_row=fila_inicio)
            fila_inicio += len(df_presentes) + 2
        else:
            ws_tipo.cell(row=fila_inicio, column=1, value='No hay registros presentes en SIRE para este tipo.')
            fila_inicio += 2

        # --- Sección: No presentes en SIRE ---
        ws_tipo.merge_cells(f'A{fila_inicio}:E{fila_inicio}')
        celda_titulo = ws_tipo.cell(row=fila_inicio, column=1, value=f'❌ NO PRESENTES EN {nombre_sire} (Total: {(~mask_presente).sum()})')
        celda_titulo.font = Font(bold=True, size=12, color="FFFFFF")
        celda_titulo.fill = PatternFill(start_color="EF4444", end_color="EF4444", fill_type="solid")
        celda_titulo.alignment = Alignment(horizontal='center', vertical='center')
        fila_inicio += 1

        df_no_presentes = df_ple_tipo[~mask_presente].copy()
        if not df_no_presentes.empty:
            cols_orden = ['ID_CONCILIACION', 'Serie', 'Numero', 'Hoja_Origen'] + \
                         [c for c in df_no_presentes.columns if c not in ['ID_CONCILIACION', 'Serie', 'Numero', 'Hoja_Origen']]
            df_no_presentes = df_no_presentes[cols_orden]
            _escribir_dataframe_con_formato(ws_tipo, df_no_presentes, start_row=fila_inicio)
        else:
            ws_tipo.cell(row=fila_inicio, column=1, value='Todos los registros de este tipo están presentes en SIRE.')

    # Guardar el archivo
    wb.save(ruta_salida)
    return ruta_salida
