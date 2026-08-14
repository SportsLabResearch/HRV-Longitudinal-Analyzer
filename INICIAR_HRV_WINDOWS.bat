@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo HRV-LONGITUDINAL-ANALYZER - INICIO LOCAL
echo ============================================================
py launcher_hrv.py
if errorlevel 1 (
    echo.
    echo Intentando con python...
    python launcher_hrv.py
)
pause
