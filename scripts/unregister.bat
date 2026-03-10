@echo off
echo ============================================
echo   COM Plugin Demo - 取消注册 COM 组件
echo   需要以管理员身份运行
echo ============================================
echo.

:: 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [错误] 请以管理员身份运行此脚本！
    echo 右键点击此文件 -^> "以管理员身份运行"
    pause
    exit /b 1
)

:: 设置路径
set SCRIPT_DIR=%~dp0
set SLN_DIR=%SCRIPT_DIR%..
set CS_PLUGIN_DIR=%SLN_DIR%\src\CSharpComPlugin\bin\Release\net8.0
set VB6_PLUGIN_DIR=%SLN_DIR%\src\VB6ComPlugin

:: 取消注册 C# COM 插件
echo [1/2] 取消注册 C# COM 插件...
if exist "%CS_PLUGIN_DIR%\CSharpComPlugin.comhost.dll" (
    regsvr32 /u /s "%CS_PLUGIN_DIR%\CSharpComPlugin.comhost.dll"
    echo       完成
) else (
    set CS_PLUGIN_DIR=%SLN_DIR%\src\CSharpComPlugin\bin\Debug\net8.0
    if exist "%SLN_DIR%\src\CSharpComPlugin\bin\Debug\net8.0\CSharpComPlugin.comhost.dll" (
        regsvr32 /u /s "%SLN_DIR%\src\CSharpComPlugin\bin\Debug\net8.0\CSharpComPlugin.comhost.dll"
        echo       完成 (Debug)
    ) else (
        echo       未找到文件，跳过
    )
)

:: 取消注册 VB6 COM 插件
echo.
echo [2/2] 取消注册 VB6 COM 插件...
if exist "%VB6_PLUGIN_DIR%\VB6ComPlugin.dll" (
    regsvr32 /u /s "%VB6_PLUGIN_DIR%\VB6ComPlugin.dll"
    echo       完成
) else (
    echo       未找到文件，跳过
)

echo.
echo ============================================
echo   取消注册完成
echo ============================================
pause
