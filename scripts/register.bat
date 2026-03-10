@echo off
echo ============================================
echo   COM 组件注册 (仅在不使用免注册方式时需要)
echo   推荐使用免注册方式 (SxS Manifest)
echo   需要管理员权限
echo ============================================
echo.

net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [错误] 请以管理员身份运行！
    pause
    exit /b 1
)

set SCRIPT_DIR=%~dp0
set SLN_DIR=%SCRIPT_DIR%..
set CS_DIR_REL=%SLN_DIR%\src\CSharpComPlugin\bin\Release\net8.0
set CS_DIR_DBG=%SLN_DIR%\src\CSharpComPlugin\bin\Debug\net8.0
set VB6_DIR=%SLN_DIR%\src\VB6ComPlugin

echo [1/2] 注册 C# COM 插件...
if exist "%CS_DIR_REL%\CSharpComPlugin.comhost.dll" (
    regsvr32 /s "%CS_DIR_REL%\CSharpComPlugin.comhost.dll"
    echo       成功 (Release)
) else if exist "%CS_DIR_DBG%\CSharpComPlugin.comhost.dll" (
    regsvr32 /s "%CS_DIR_DBG%\CSharpComPlugin.comhost.dll"
    echo       成功 (Debug)
) else (
    echo       未找到，请先编译
)

echo.
echo [2/2] 注册 VB6 COM 插件...
if exist "%VB6_DIR%\VB6ComPlugin.dll" (
    regsvr32 /s "%VB6_DIR%\VB6ComPlugin.dll"
    echo       成功
) else (
    echo       未找到 VB6ComPlugin.dll
)

echo.
pause
