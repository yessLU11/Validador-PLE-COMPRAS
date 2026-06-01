# Parámetros: columnas clave, nombres de hojas, meses a mantener

# config.py
# Este módulo define los parámetros clave para la comparación de datos, como las columnas que se 
# usarán para detectar duplicados, los nombres de las hojas donde se encuentran los datos, 
# y cuántos meses de histórico mantener para la comparación.
# Este modulo se conecta con validator.py para que este sepa qué columnas usar para comparar, y 
# con el main.py para que este sepa qué hojas leer y cuántos meses mantener.
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

# Mapeo de nombres de columna originales (Excel) a nuestros nombres internos (minúsculas)
# Esto es útil para mantener el código más limpio y evitar errores de espacios o caracteres especiales.
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
HOJA_SOCIALES = "PROGRAMAS SOCIALES"
FILA_INICIO_PRINCIPAL = 8   # datos empiezan fila 8
FILA_INICIO_SOCIALES = 2 # datos empiezan fila 2

MESES_A_MANTENER = 12 