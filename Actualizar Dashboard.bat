@echo off
title Actualizando Dashboard - Stock U.Unidad...

echo.
echo  ================================================
echo   STOCK U.UNIDAD - Actualizacion de datos
echo  ================================================
echo.

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

:: Publicar en GitHub Pages
echo  Publicando en GitHub Pages...
if not exist "%~dp0docs" mkdir "%~dp0docs"
copy /Y "%~dp0Dashboard.html" "%~dp0docs\index.html" >nul

for /f "tokens=1-3 delims=/ " %%a in ("%DATE%") do set HOY=%%c-%%b-%%a
for /f "tokens=1-2 delims=:." %%a in ("%TIME: =0%") do set HORA=%%a:%%b

cd /d "%~dp0"
git add docs\index.html
git commit -m "Actualizar dashboard %HOY% %HORA%"
git push origin master

if errorlevel 1 (
    echo.
    echo  AVISO: No se pudo publicar en GitHub. Revisa la conexion.
    echo  El dashboard local funciona igual.
    echo.
) else (
    echo.
    echo  Publicado en: https://diego-bels.github.io/Dashboard-Stock-OBSA/
    echo.
)

echo  Abriendo Dashboard en el navegador...
start "" "%~dp0Dashboard.html"

timeout /t 2 >nul
