@echo off
chcp 65001 >nul
title WAI 插件安装程序

echo ============================================
echo   WAI - Word AI 侧边栏插件 一键安装
echo ============================================
echo.

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [!] 请右键选择"以管理员身份运行"本程序。
    pause
    exit /b 1
)

echo [1/3] 安装 VSTO 运行时依赖...
:: VSTO Runtime 通常已随 Office 安装，无需额外操作

echo [2/3] 安装插件...
:: 搜索当前目录下的 .vsto 文件
for %%f in (*.vsto) do (
    echo     找到安装包: %%f
    echo     正在安装，请等待...
    start /wait "" "%%f"
    goto :installed
)

echo [!] 未找到 .vsto 安装文件。
echo     请确保 Install.bat 和 WaiAddIn.vsto 在同一目录。
pause
exit /b 1

:installed
echo [3/3] 安装完成！
echo.
echo ============================================
echo   ✅ WAI 已安装！
echo   请重新启动 Word，在右侧即可看到 WAI 面板。
echo   首次使用请在插件设置中配置 API Key。
echo ============================================
echo.
pause
