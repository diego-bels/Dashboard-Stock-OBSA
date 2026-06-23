@echo off
title Actualizando Dashboard - Stock U.Unidad...

echo.
echo  ================================================
echo   STOCK U.UNIDAD - Actualizacion de datos
echo  ================================================
echo.

:: Verificar que existe al menos un archivo Stock Total*.xlsx
set FOUND=0
for %%f in ("%~dp0Stock Total*.xlsx") do set FOUND=1

if %FOUND%==0 (
    echo  ERROR: No se encontro el archivo de stock en esta carpeta.
    echo.
    echo  Copia el archivo exportado del sistema a esta carpeta.
    echo  El nombre debe empezar con "Stock Total" (ej: Stock Total 2206.xlsx)
    echo.
    pause
    exit /b 1
)

echo  Paso 1/2: Extrayendo datos del Excel...
python "%~dp01_extraer_datos.py"
if errorlevel 1 (
    echo.
    echo  ERROR en la extraccion de datos.
    pause
    exit /b 1
)

echo.
echo  Paso 2/2: Generando Dashboard HTML...
python "%~dp02_generar_dashboard.py"
if errorlevel 1 (
    echo.
    echo  ERROR al generar el Dashboard.
    pause
    exit /b 1
)

echo.
echo  ================================================
echo   Dashboard actualizado correctamente!
echo  ================================================
echo.
echo  Abriendo Dashboard en el navegador...
start "" "%~dp0Dashboard.html"

timeout /t 2 >nul
