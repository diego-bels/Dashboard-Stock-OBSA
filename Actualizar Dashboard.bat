@echo off
chcp 65001 >nul
title Actualizando Dashboard Stock LED...

echo.
echo  ================================================
echo   DASHBOARD STOCK LED — Actualización de datos
echo  ================================================
echo.

:: Verificar que existe el archivo Excel
if not exist "%~dp0Stock LED.xlsx" (
    echo  ERROR: No se encontró "Stock LED.xlsx" en esta carpeta.
    echo.
    echo  Copiá el archivo actualizado de stock aquí antes de continuar.
    echo.
    pause
    exit /b 1
)

echo  Paso 1/2: Extrayendo datos del Excel...
python "%~dp01_extraer_datos.py"
if errorlevel 1 (
    echo.
    echo  ERROR en la extracción de datos.
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
echo  Abriendo Dashboard.html en el navegador...
start "" "%~dp0Dashboard.html"

timeout /t 2 >nul
