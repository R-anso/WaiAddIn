@echo off
chcp 65001 >nul
title WAI 插件卸载程序

echo ============================================
echo   WAI - Word AI 侧边栏插件 卸载
echo ============================================
echo.

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [!] 请右键选择"以管理员身份运行"本程序。
    pause
    exit /b 1
)

echo [1/2] 从系统中移除 WAI 插件...

:: 通过注册表查找并移除 ClickOnce 安装的 WAI 插件
for /f "tokens=*" %%a in ('dir "%userprofile%\AppData\Local\Apps\2.0\*WaiAddIn*" /s /b /ad 2^>nul') do (
    echo     找到插件目录: %%a
    rd /s /q "%%a" 2>nul
)

:: 清除 VSTO 缓存
if exist "%localappdata%\Microsoft\VSTO\Cache" (
    echo     清除 VSTO 缓存...
    del /f /q "%localappdata%\Microsoft\VSTO\Cache\*WaiAddIn*" 2>nul
)

:: 清除本地设置数据
if exist "%localappdata%\WAI" (
    echo     清除本地设置...
    rd /s /q "%localappdata%\WAI" 2>nul
)

echo [2/2] 正在清除 Word COM 加载项注册...
:: 移除 Word 加载项注册表项
reg delete "HKCU\Software\Microsoft\Office\Word\Addins\WaiAddIn" /f 2>nul
reg delete "HKLM\Software\Microsoft\Office\Word\Addins\WaiAddIn" /f 2>nul

echo.
echo ============================================
echo   ✅ WAI 已从系统中卸载。
echo   请重启 Word 确认插件已移除。
echo ============================================
echo.
pause
