# config.py - Configuración de parámetros clave para el validador PLE Compras

COLUMNAS_CLAVE = [
    "fecha_emision",      # columna E
    "tipo_comprobante",   # columna G
    "serie_comprobante",  # columna H
    "numero_comprobante", # columna J
    "ruc_proveedor",      # columna M
    "razon_social",       # columna N
    "base_imponible",     # columna Q
    "igv",                # columna R
    "importe_total"       # columna Y
]

MAPEO_COLUMNAS = {
    "B": "Periodo (AAAAMM00)",
    "C": "Código Único de la Operación (CUO)",
    "E": "fecha_emision",
    "G": "tipo_comprobante",
    "H": "serie_comprobante",
    "J": "numero_comprobante",
    "M": "ruc_proveedor",
    "N": "razon_social",
    "Q": "base_imponible",
    "R": "igv",
    "Y": "importe_total"
}

HOJA_PRINCIPAL = "8.1"
HOJA_SOCIALES = "Programas Sociales"
FILA_INICIO_PRINCIPAL = 8   # datos empiezan fila 8
FILA_INICIO_SOCIALES = 2    # datos empiezan fila 2 

# ===== NUEVAS CONFIGURACIONES =====
# Fila de inicio predeterminada para hojas no especificadas (8.2, 8.3, etc.)
FILA_INICIO_PREDETERMINADA = 1

# Opcional: puedes definir filas de inicio específicas para otras hojas
HOJAS_ADICIONALES = {
    "8.2": 1,
    "8.3": 1,
    # Agrega más si es necesario
}

MESES_A_MANTENER = 12