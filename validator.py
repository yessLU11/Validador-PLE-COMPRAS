 # Lógica de comparación y detección de duplicados
# validator.py
# ¿qué hace este apartado?  
# Este módulo se encarga de comparar el DataFrame del mes nuevo con el histórico para detectar duplicados,
# es decir, filas del nuevo que ya existían en meses anteriores. Retorna un DataFrame con los duplicados encontrados, incluyendo información de dónde se encontró el match (mes y hoja).
#¿Cómo lo hace?
# Utiliza un merge (inner join) entre el DataFrame nuevo y el histórico, 
# usando las columnas clave definidas en config.py. Luego renombra algunas columnas para que el reporte sea más claro.   
import pandas as pd
from config import COLUMNAS_CLAVE

# Columnas adicionales que se mostrarán en el reporte pero no se usan para la comparación
COLUMNAS_ADICIONALES = [
    "Periodo (AAAAMM00)",
    "Código Único de la Operación (CUO)"
]

def detectar_duplicados(df_nuevo, df_historico):
    """
    df_nuevo: DataFrame del mes actual (ya normalizado)
    df_historico: DataFrame con todos los registros de meses anteriores (sin incluir el nuevo)
    Retorna: DataFrame con los duplicados encontrados (cada fila del nuevo que tiene match)
            con columnas adicionales para el reporte final
    """
    # Hacemos un merge (inner join) usando las columnas clave
    # Para evitar comparar consigo mismo, aseguramos que df_historico no contenga el mismo mes
    duplicados = pd.merge(
        df_nuevo[COLUMNAS_CLAVE + COLUMNAS_ADICIONALES + ["hoja_origen", "mes_archivo"]],
        df_historico[COLUMNAS_CLAVE + ["mes_archivo", "hoja_origen"]],
        on=COLUMNAS_CLAVE,
        how="inner",
        suffixes=("_nuevo", "_existente")
    )
    if duplicados.empty:
        return pd.DataFrame()
    # Renombrar para reporte claro
    duplicados.rename(columns={
        "mes_archivo_existente": "Mes(es)_donde_hay_duplicados",
        "mes_archivo_nuevo": "mes_archivo_nuevo",
        "hoja_origen_nuevo": "hoja_en_mes_nuevo",
        "hoja_origen_existente": "hoja_en_mes(es)_anteriores"
    }, inplace=True)
    
    # Reordenar columnas para que Periodo y CUO estén al inicio
    columnas_reporte = (
        COLUMNAS_ADICIONALES +
        COLUMNAS_CLAVE +
        ["hoja_en_mes_nuevo", "mes_archivo_nuevo", "Mes(es)_donde_hay_duplicados", "hoja_en_mes(es)_anteriores"]
    )
    
    # Asegurar que todas las columnas esperadas existan
    columnas_existentes = [col for col in columnas_reporte if col in duplicados.columns]
    return duplicados[columnas_existentes]