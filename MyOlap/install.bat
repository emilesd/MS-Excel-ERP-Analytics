@echo off
echo ============================================
echo   MyOlap v1.0 - Excel Add-in Installer
echo   GoLive Systems Ltd
echo ============================================
echo.

REM Detect Excel bitness from registry
set XLL_FILE=MyOlap-AddIn64-packed.xll
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Excel" >nul 2>&1
if %errorlevel%==0 (
    echo Detected 64-bit Office installation.
    set XLL_FILE=MyOlap-AddIn64-packed.xll
) else (
    echo Assuming 32-bit Office installation.
    set XLL_FILE=MyOlap-AddIn-packed.xll
)

REM Create install directory
set INSTALL_DIR=%LOCALAPPDATA%\MyOlap\AddIn
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM Copy the packed XLL
echo Copying add-in files to %INSTALL_DIR%...
copy /Y "publish\%XLL_FILE%" "%INSTALL_DIR%\MyOlap-AddIn.xll" >nul

echo.
echo ============================================
echo   Installation Complete!
echo ============================================
echo.
echo To activate the add-in in Excel:
echo   1. Open Excel
echo   2. Go to File ^> Options ^> Add-ins
echo   3. At the bottom, select "Excel Add-ins" and click "Go..."
echo   4. Click "Browse..." and navigate to:
echo      %INSTALL_DIR%\MyOlap-AddIn.xll
echo   5. Click OK - the "MyOlap" tab will appear on the ribbon.
echo.
echo Alternatively, just double-click the .xll file to load it.
echo.
pause
