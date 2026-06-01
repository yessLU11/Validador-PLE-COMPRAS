@echo off
chcp 65001 >nul
echo ============================================================
echo REINICIALIZANDO BASE DE DATOS - PLE COMPRAS
echo ============================================================
echo.

cd /d "%~dp0"
python fix_db.py

if errorlevel 1 (
    echo.
    echo ============================================================
    echo ERROR: No se pudo reinicializar la base de datos
    echo ============================================================
    pause
    exit /b 1
) else (
    echo.
    echo ============================================================
    echo LISTO: Presiona cualquier tecla para continuar...
    echo ============================================================
    pause
)
