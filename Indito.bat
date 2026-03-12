@echo off
chcp 65001 >nul
title UPS PoD Letolto

:: Proxy beallitas
set PROXY=http://cloudproxy.dhl.com:10123

:: Csomagok ellenorzese
py -c "import pandas, openpyxl, selenium" >nul 2>&1
if %errorlevel% neq 0 (
    echo Hianyoznak a Python csomagok, telepites folyamatban...
    py -m pip install --user --proxy %PROXY% pandas openpyxl selenium
    
    :: Ujraellenorzes
    py -c "import pandas, openpyxl, selenium" >nul 2>&1
    if %errorlevel% neq 0 (
        echo [HIBA] Telepites sikertelen! Kerjen segitseget az adminisztatortol.
        pause
        exit /b 1
    )
    echo Csomagok sikeresen telepitve!
)

powershell -ExecutionPolicy Bypass -File "%~dp0UPS_PoD_Letolto.ps1"
pause
