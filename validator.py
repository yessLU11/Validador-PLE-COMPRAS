# validator.py - VERSIÓN REFACTORIZADA
import pandas as pd
from config import COLUMNAS_CLAVE
from database import marcar_duplicados_reportados

COLUMNAS_ADICIONALES = [
    "Periodo (AAAAMM00)",
    "Código Único de la Operación (CUO)"
]

def detectar_duplicados(df_nuevo, df_historico, mes_archivo=None):
    """
    Detecta duplicados y marca cuáles ya fueron reportados.
    """
    if df_nuevo.empty or df_historico.empty:
        return pd.DataFrame()
    
    duplicados = pd.merge(
        df_nuevo[COLUMNAS_CLAVE + COLUMNAS_ADICIONALES + ["hoja_origen", "mes_archivo"]],
        df_historico[COLUMNAS_CLAVE + ["mes_archivo", "hoja_origen"]],
        on=COLUMNAS_CLAVE,
        how="inner",
        suffixes=("_nuevo", "_existente")
    )
    
    if duplicados.empty:
        return pd.DataFrame()
    
    duplicados.rename(columns={
        "mes_archivo_existente": "Mes(es)_donde_hay_duplicados",
        "mes_archivo_nuevo": "mes_archivo_nuevo",
        "hoja_origen_nuevo": "hoja_en_mes_nuevo",
        "hoja_origen_existente": "hoja_en_mes(es)_anteriores"
    }, inplace=True)

    # Marcar duplicados ya reportados
    if mes_archivo:
        duplicados = marcar_duplicados_reportados(duplicados, mes_archivo)
    else:
        duplicados['ya_reportado'] = False
        duplicados['meses_previos'] = ''
        
    columnas_reporte = (
        COLUMNAS_ADICIONALES +
        COLUMNAS_CLAVE +
        ["hoja_en_mes_nuevo", "mes_archivo_nuevo", "Mes(es)_donde_hay_duplicados", 
         "hoja_en_mes(es)_anteriores", "ya_reportado", "meses_previos"]
    )
    
    columnas_existentes = [col for col in columnas_reporte if col in duplicados.columns]
    return duplicados[columnas_existentes]