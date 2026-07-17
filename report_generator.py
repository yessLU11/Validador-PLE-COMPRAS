# report_generator.py - VERSIÓN REFACTORIZADA
import pandas as pd

def generar_reporte_excel(duplicados_df, nombre_salida="reporte_duplicados.xlsx"):
    """
    Crea un archivo Excel con:
    - Duplicados_Detalle: tabla completa con estado de reporte
    - Resumen_por_Mes: conteo de duplicados agrupado por mes y tipo
    - Estado_Reportes: resumen de nuevos vs ya reportados
    """
    with pd.ExcelWriter(nombre_salida, engine="openpyxl") as writer:
        # Hoja detalle
        duplicados_df.to_excel(writer, sheet_name="Duplicados_Detalle", index=False)
        
        if not duplicados_df.empty:
            # Hoja resumen
            resumen_col = "Mes(es)_donde_hay_duplicados" if "Mes(es)_donde_hay_duplicados" in duplicados_df.columns else "Mes_donde_ya_existia"
            if resumen_col in duplicados_df.columns and "tipo_comprobante" in duplicados_df.columns:
                resumen = duplicados_df.groupby([resumen_col, "tipo_comprobante"]).size().reset_index(name="cantidad_duplicados")
                resumen.to_excel(writer, sheet_name="Resumen_por_Mes", index=False)
            
            # Hoja: Estado de Reportes
            if 'ya_reportado' in duplicados_df.columns:
                estado_resumen = duplicados_df.groupby('ya_reportado').size().reset_index(name='cantidad')
                estado_resumen['Estado'] = estado_resumen['ya_reportado'].map({True: 'Ya reportados', False: 'Nuevos'})
                estado_resumen[['Estado', 'cantidad']].to_excel(writer, sheet_name="Estado_Reportes", index=False)
        else:
            pd.DataFrame({"Mensaje": ["No se encontraron duplicados"]}).to_excel(writer, sheet_name="Resumen_por_Mes", index=False)
    
    return nombre_salida