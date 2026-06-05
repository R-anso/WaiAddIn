@echo off
chcp 437 >nul
setlocal enabledelayedexpansion

title WAI Add-in Installer

echo ============================================
echo   WAI - Word AI Sidebar v0.1.0
echo   One-click Installer
echo ============================================
echo.

:: Run as admin check
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Please right-click Install.bat and select "Run as administrator".
    echo.
    pause
    exit /b 1
)

:: Get script directory
set SCRIPT_DIR=%~dp0

echo [1/3] Installing VSTO add-in...

if not exist "%SCRIPT_DIR%WaiAddIn.vsto" (
    echo [ERROR] WaiAddIn.vsto not found.
    echo         Please make sure all files are in the same folder as Install.bat.
    echo.
    pause
    exit /b 1
)

:: Launch VSTO installer (clickonce)
start /wait "" "%SCRIPT_DIR%WaiAddIn.vsto"

echo [2/3] Checking installation...

reg add "HKCU\Software\Microsoft\Office\Word\Addins\WaiAddIn" /f >nul 2>&1

echo [3/3] Done!
echo.
echo ============================================
echo   SUCCESS! WAI has been installed.
echo
echo   Please restart Microsoft Word.
echo   You will see the WAI panel on the right side.
echo
echo   First time: Click "Settings" button
echo   at the bottom of the panel to enter your API Key.
echo ============================================
echo.
pause
