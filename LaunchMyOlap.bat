@echo off
title MyOlap Launcher
echo ============================================
echo   MyOlap v2.2 Add-in Launcher
echo ============================================
echo.

echo Step 1: Closing any running Excel...
taskkill /f /im EXCEL.EXE >nul 2>&1
timeout /t 3 /nobreak >nul

echo Step 2: Clearing Excel cache and resiliency data...
reg delete "HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency" /f >nul 2>&1

echo Step 3: Registering add-in for auto-load...
set XLL=%LOCALAPPDATA%\MyOlap\MyOlap-AddIn64.xll
if not exist "%XLL%" (
    echo ERROR: Add-in not found at %XLL%
    pause
    exit /b 1
)
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Options" /v OPEN /t REG_SZ /d "/R \"%XLL%\"" /f >nul

echo Step 4: Opening Excel...
echo.
echo LOOK FOR: "MyOlap v2.2" tab in the ribbon
echo.
start "" "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
echo Excel launched. This window will close in 5 seconds.
timeout /t 5
