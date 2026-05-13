@echo off
setlocal
title MyOlap Launcher

echo ============================================
echo   MyOlap Add-in Launcher
echo ============================================
echo.

set "SCRIPT_DIR=%~dp0"
set "PROJECT=%SCRIPT_DIR%MyOlap\MyOlap.csproj"
set "BUILD_OUT=%SCRIPT_DIR%MyOlap\bin\Release\net8.0-windows"
set "DEPLOY=%LOCALAPPDATA%\MyOlap"
set "XLL=%DEPLOY%\MyOlap-AddIn64.xll"
set "EXCEL=C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"

echo [1/5] Closing any running Excel...
taskkill /f /im EXCEL.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

echo [2/5] Clearing Excel resiliency data...
reg delete "HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency" /f >nul 2>&1

echo [3/5] Building add-in (Release)...
where dotnet >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: .NET SDK not found on PATH.
    echo Install .NET 8 SDK ^(x64^): https://dotnet.microsoft.com/en-us/download/dotnet/8.0
    echo Then re-run this script.
    pause
    exit /b 1
)
dotnet build "%PROJECT%" -c Release --nologo -v minimal
if errorlevel 1 (
    echo.
    echo ERROR: Build failed. See messages above.
    pause
    exit /b 1
)

echo [4/5] Deploying to %DEPLOY%...
if not exist "%DEPLOY%" mkdir "%DEPLOY%" >nul 2>&1
xcopy /s /y /i /q "%BUILD_OUT%\*" "%DEPLOY%\" >nul
if not exist "%XLL%" (
    echo ERROR: Add-in not found at %XLL% after deploy.
    pause
    exit /b 1
)

echo [5/5] Registering add-in for auto-load and launching Excel...
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Options" /v OPEN /t REG_SZ /d "/R \"%XLL%\"" /f >nul

echo.
echo Look for the "MyOlap" tab in the Excel ribbon.
echo.
start "" "%EXCEL%"

timeout /t 3 >nul
endlocal
