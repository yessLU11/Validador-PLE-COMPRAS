"""
Script para eliminar la BD antigua y reinicializar con el nuevo esquema
"""
import os
import sys
import shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Importar después de agregar al path
from database import init_db, DB_PATH

print("=" * 60)
print("REINICIALIZANDO BASE DE DATOS")
print("=" * 60)

try:
    print(f"\nRuta de BD: {DB_PATH}")
    
    # Eliminar BD si existe
    if os.path.exists(DB_PATH):
        print(f"❌ Eliminando BD antigua: {DB_PATH}")
        try:
            os.remove(DB_PATH)
            print("✓ BD eliminada correctamente")
        except Exception as e:
            print(f"⚠️ Error al eliminar: {e}")
            print("Intentando con shutil...")
            shutil.rmtree(os.path.dirname(DB_PATH), ignore_errors=True)
    else:
        print("ℹ️ No hay BD anterior, continuando...")
    
    # Reinicializar con nuevo esquema
    print("\n📚 Inicializando BD con nuevo esquema...")
    init_db()
    print("✓ BD inicializada correctamente")
    
    print("\n" + "=" * 60)
    print("✅ ¡ÉXITO! El sistema está listo para cargar archivos nuevos.")
    print("=" * 60)
    
except Exception as e:
    print(f"\n❌ ERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
