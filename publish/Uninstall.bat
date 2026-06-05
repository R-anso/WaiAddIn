@echo off
chcp 437 >nul
setlocal enabledelayedexpansion

title WAI Add-in Uninstaller

echo ============================================
echo   WAI - Word AI Sidebar v0.1.0
echo   Uninstaller
echo ============================================
echo.

:: Run as admin check
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Please right-click Uninstall.bat and select "Run as administrator".
    echo.
    pause
    exit /b 1
)

echo [1/3] Removing VSTO add-in registration...

:: Remove COM add-in registry entries
reg delete "HKCU\Software\Microsoft\Office\Word\Addins\WaiAddIn" /f >nul 2>&1
reg delete "HKLM\Software\Microsoft\Office\Word\Addins\WaiAddIn" /f >nul 2>&1

echo [2/3] Clearing local settings and cache...

:: Remove settings directory
if exist "%localappdata%\WAI" (
    rd /s /q "%localappdata%\WAI"
)

:: Remove VSTO ClickOnce cache
if exist "%localappdata%\Apps\2.0" (
    for /f "tokens=*" %%a in ('dir "%localappdata%\Apps\2.0\*WaiAddIn*" /s /b /ad 2^>nul') do (
        rd /s /q "%%a" 2>nul
    )
)

echo [3/3] Done!
echo.
echo ============================================
echo   SUCCESS! WAI has been uninstalled.
echo
echo   Please restart Microsoft Word.
echo ============================================
echo.
pause
