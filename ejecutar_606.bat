@echo off
chcp 65001 > nul
title Generador DGII 606 y Compras sin Comprobante
color 1F

echo =======================================================================
echo          GENERADOR DEL FORMATO 606 (DGII) Y COMPRAS
echo =======================================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] No se detecto Python instalado en el sistema.
    echo Por favor instale Python 3.10 o superior y marque la casilla "Add to PATH".
    echo.
    pause
    exit /b 1
)

echo Seleccione como desea ejecutar el programa:
echo.
echo   [1] Iniciar Interfaz Grafica (Recomendado - Ventana con botones)
echo   [2] Procesar directamente con muestra.xlsx y Proveedores.xlsx
echo   [3] Ejecutar pruebas unitarias automatizadas (pytest)
echo   [4] Salir
echo.
set /p opcion="Opcion [1]: "

if "%opcion%"=="" set opcion=1
if "%opcion%"=="1" goto gui
if "%opcion%"=="2" goto consola
if "%opcion%"=="3" goto pruebas
if "%opcion%"=="4" goto salir

:gui
echo.
echo Iniciando interfaz grafica...
python app_gui.py
goto fin

:consola
echo.
echo Ejecutando procesamiento en consola...
python generar_606.py --transacciones muestra.xlsx --proveedores Proveedores.xlsx --salida ./salida
echo.
pause
goto fin

:pruebas
echo.
echo Ejecutando pruebas con pytest...
pytest -v
echo.
pause
goto fin

:salir
exit /b 0

:fin
