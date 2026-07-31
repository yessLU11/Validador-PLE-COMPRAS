# src/sire_utils.py
"""
Utilidades para el Convertidor SIRE/PLE
"""

import os
import pandas as pd
from datetime import datetime


def obtener_tamano_archivo(ruta):
    """Obtiene el tamaño de un archivo en MB"""
    try:
        tamaño_bytes = os.path.getsize(ruta)
        return tamaño_bytes / (1024 * 1024)
    except:
        return 0


def obtener_info_archivo(ruta):
    """Obtiene información detallada de un archivo"""
    try:
        stats = os.stat(ruta)
        return {
            'nombre': os.path.basename(ruta),
            'tamaño_mb': stats.st_size / (1024 * 1024),
            'tamaño_bytes': stats.st_size,
            'fecha_modificacion': datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
        }
    except:
        return None


def limpiar_valor_numerico(valor):
    """Limpia un valor numérico para conversión"""
    if pd.isna(valor):
        return 0
    if isinstance(valor, (int, float)):
        return valor
    if isinstance(valor, str):
        # Remover caracteres no numéricos excepto punto y coma
        valor = valor.replace(',', '').strip()
        try:
            return float(valor)
        except:
            return 0
    return 0


def formatear_moneda(valor):
    """Formatea un valor como moneda"""
    try:
        return f"{float(valor):,.2f}"
    except:
        return "0.00"