# excel_reader.py - VERSIÓN CORREGIDA (LEE TODAS LAS HOJAS)
import pandas as pd
from config import COLUMNAS_CLAVE, MAPEO_COLUMNAS, HOJA_PRINCIPAL, HOJA_SOCIALES
from config import FILA_INICIO_PRINCIPAL, FILA_INICIO_SOCIALES, FILA_INICIO_PREDETERMINADA

def normalizar_valor(valor):
    """Limpia espacios, convierte a string, mayúsculas para texto."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, (int, float)):
        return str(valor).strip()
    if isinstance(valor, pd.Timestamp):
        return valor.strftime("%Y-%m-%d")
    return str(valor).strip().upper()

def leer_hoja_unificada(archivo_excel, mes_archivo):
    """
    Lee TODAS las hojas del archivo Excel que contengan datos válidos.
    Ahora soporta 8.1, 8.2, 8.3, Programas Sociales, etc.
    """
    dfs = []
    
    def letra_a_index(letra):
        letra = letra.upper()
        idx = 0
        for c in letra:
            idx = idx * 26 + (ord(c) - ord('A') + 1)
        return idx - 1
    
    columnas_interes = {col_letra: letra_a_index(col_letra) for col_letra in MAPEO_COLUMNAS.keys()}
    nombres_columnas = list(MAPEO_COLUMNAS.values())
    
    # Usar context manager para abrir el archivo UNA SOLA VEZ
    with pd.ExcelFile(archivo_excel) as xlsx:
        hojas_disponibles = xlsx.sheet_names
        
        for hoja in hojas_disponibles:
            try:
                # Determinar fila de inicio según el nombre de la hoja
                if hoja == HOJA_PRINCIPAL:
                    fila_inicio = FILA_INICIO_PRINCIPAL - 1  # 7
                elif hoja == HOJA_SOCIALES:
                    fila_inicio = FILA_INICIO_SOCIALES - 1  # 1
                else:
                    # Para hojas como 8.2, 8.3, etc.
                    fila_inicio = FILA_INICIO_PREDETERMINADA - 1  # 0

                df_hoja = pd.read_excel(xlsx, sheet_name=hoja, header=None)
                df_hoja = df_hoja.iloc[fila_inicio:, :].copy()

                # Verificar que la hoja tenga datos
                if df_hoja.empty:
                    continue

                # Verificar que tenga suficientes columnas (al menos 11)
                if df_hoja.shape[1] < 11:
                    print(f"Hoja '{hoja}' tiene {df_hoja.shape[1]} columnas, se necesitan al menos 11. Saltando...")
                    continue

                # Verificar que haya datos en las columnas clave (no estén todas vacías)
                indices_interes = [
                    columnas_interes[letra]
                    for letra in MAPEO_COLUMNAS.keys()
                    if columnas_interes[letra] < df_hoja.shape[1]
                ]
                
                if not indices_interes:
                    continue

                # Verificar si hay al menos una fila con datos
                primeras_filas = df_hoja.head(10)
                datos_validos = False
                for idx in indices_interes:
                    if not primeras_filas.iloc[:, idx].isna().all():
                        datos_validos = True
                        break

                if not datos_validos:
                    print(f"Hoja '{hoja}' no tiene datos válidos. Saltando...")
                    continue

                # Extraer solo las columnas que nos interesan
                indices = [
                    columnas_interes[letra]
                    for letra in MAPEO_COLUMNAS.keys()
                    if columnas_interes[letra] < df_hoja.shape[1]
                ]

                df_hoja_sub = df_hoja.iloc[:, indices].copy()
                df_hoja_sub.columns = nombres_columnas[:len(indices)]
                df_hoja_sub["hoja_origen"] = hoja

                dfs.append(df_hoja_sub)
                print(f"✅ Hoja '{hoja}' procesada correctamente con {len(df_hoja_sub)} filas")

            except Exception as e:
                print(f"❌ Error procesando hoja '{hoja}': {e}")
                continue

    if not dfs:
        print("No se encontraron hojas válidas para procesar.")
        return pd.DataFrame()

    df_final = pd.concat(dfs, ignore_index=True)

    # Limpiar valores nulos en columnas clave
    columnas_clave_existentes = [col for col in COLUMNAS_CLAVE if col in df_final.columns]
    if columnas_clave_existentes:
        df_final = df_final.dropna(how='all', subset=columnas_clave_existentes)

    # Normalizar columnas clave
    for col in COLUMNAS_CLAVE:
        if col in df_final.columns:
            df_final[col] = df_final[col].apply(normalizar_valor)

    df_final["mes_archivo"] = mes_archivo

    print(f"✅ Total de registros procesados: {len(df_final)}")
    return df_final