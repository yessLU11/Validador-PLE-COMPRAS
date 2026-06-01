 # Generación de reporte Excel (con openpyxl)
# report_generator.py
# este módulo se encarga de generar un reporte Excel con los datos de los duplicados encontrados
# 
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill

def generar_reporte_excel(duplicados_df, nombre_salida="reporte_duplicados.xlsx"):
    """
    Crea un archivo Excel con dos hojas:
    - Duplicados_Detalle: tabla completa
    - Resumen_por_Mes: conteo de duplicados agrupado por mes y tipo comprobante
    """
    with pd.ExcelWriter(nombre_salida, engine="openpyxl") as writer:
        # Hoja detalle
        duplicados_df.to_excel(writer, sheet_name="Duplicados_Detalle", index=False)
        
        # Hoja resumen
        if not duplicados_df.empty: # Si hay duplicados, hacemos el resumen, empty es un DataFrame vacío
            resumen_col = "Mes(es)_donde_hay_duplicados" if "Mes(es)_donde_hay_duplicados" in duplicados_df.columns else "Mes_donde_ya_existia"
            if resumen_col in duplicados_df.columns and "tipo_comprobante" in duplicados_df.columns:
                resumen = duplicados_df.groupby([resumen_col, "tipo_comprobante"]).size().reset_index(name="cantidad_duplicados")
                resumen.to_excel(writer, sheet_name="Resumen_por_Mes", index=False)
            else:
                pd.DataFrame({"Mensaje": ["No hay columnas necesarias para generar el resumen"]}).to_excel(writer, sheet_name="Resumen_por_Mes", index=False)
        else:
            pd.DataFrame({"Mensaje": ["No se encontraron duplicados"]}).to_excel(writer, sheet_name="Resumen_por_Mes", index=False)
    
    # Opcional: resaltar en rojo celdas (ejemplo simple)
    # Se puede ampliar después
    return nombre_salida