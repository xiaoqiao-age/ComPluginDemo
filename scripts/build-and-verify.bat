@echo off
echo ============================================
echo   COM Plugin Demo - 免注册 COM (Reg-Free)
echo   无需 regsvr32, 通过 SxS Manifest 加载
echo ============================================
echo.

:: 设置路径
set SCRIPT_DIR=%~dp0
set SLN_DIR=%SCRIPT_DIR%..

echo [1] 编译解决方案...
dotnet build "%SLN_DIR%\ComPluginDemo.sln" -c Debug
if %errorLevel% neq 0 (
    echo [错误] 编译失败！
    pause
    exit /b 1
)

echo.
echo [2] 验证输出文件...
set HOST_DIR=%SLN_DIR%\src\ComPluginHost\bin\Debug\net8.0-windows

echo     检查 ComPluginHost.exe...
if exist "%HOST_DIR%\ComPluginHost.exe" (
    echo       [OK]
) else (
    echo       [MISSING]
)

echo     检查 CSharpComPlugin\CSharpComPlugin.comhost.dll...
if exist "%HOST_DIR%\CSharpComPlugin\CSharpComPlugin.comhost.dll" (
    echo       [OK]
) else (
    echo       [MISSING]
)

echo     检查 CSharpComPlugin\CSharpComPlugin.manifest...
if exist "%HOST_DIR%\CSharpComPlugin\CSharpComPlugin.manifest" (
    echo       [OK]
) else (
    echo       [MISSING]
)

echo     检查 VB6ComPlugin\VB6ComPlugin.manifest...
if exist "%HOST_DIR%\VB6ComPlugin\VB6ComPlugin.manifest" (
    echo       [OK]
) else (
    echo       [MISSING]
)

echo     检查 VB6ComPlugin\VB6ComPlugin.dll...
if exist "%HOST_DIR%\VB6ComPlugin\VB6ComPlugin.dll" (
    echo       [OK]
) else (
    echo       [需手动放置] VB6 编译后将 DLL 复制到:
    echo       %HOST_DIR%\VB6ComPlugin\
)

echo.
echo ============================================
echo   运行方式:
echo     cd %HOST_DIR%
echo     ComPluginHost.exe
echo.
echo   免注册 COM 说明:
echo     - 不需要 regsvr32 注册
echo     - 不需要管理员权限
echo     - 通过 exe.manifest 声明 SxS 依赖
echo     - 每个插件有独立的组件清单
echo ============================================
pause
