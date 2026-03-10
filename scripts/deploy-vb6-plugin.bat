@echo off
chcp 936 >nul
echo ============================================
echo   VB6 插件部署工具
echo   自动读取 CLSID + 更新 manifest + 复制 DLL
echo ============================================
echo.

set SCRIPT_DIR=%~dp0
set SLN_DIR=%SCRIPT_DIR%..
set VB6_SRC=%SLN_DIR%\src\VB6ComPlugin
set HOST_OUT=%SLN_DIR%\src\ComPluginHost\bin\Debug\net8.0-windows\VB6ComPlugin

:: 1. 从注册表读取 VB6 实际 CLSID
echo [1/3] 读取 VB6ComPlugin.StringProcessor 的 CLSID...
for /f "tokens=2*" %%a in ('reg query "HKCR\VB6ComPlugin.StringProcessor\CLSID" /ve 2^>nul') do set CLSID=%%b

if "%CLSID%"=="" (
    echo   [错误] 注册表中找不到 VB6ComPlugin.StringProcessor
    echo   请先在 VB6 IDE 中编译项目 (File -^> Make VB6ComPlugin.dll)
    pause
    exit /b 1
)
echo   CLSID = %CLSID%

:: 2. 查找 VB6 编译的 DLL (InprocServer32 路径)
echo.
echo [2/3] 查找 VB6ComPlugin.dll 路径...
for /f "tokens=2*" %%a in ('reg query "HKCR\CLSID\%CLSID%\InprocServer32" /ve 2^>nul') do set DLL_PATH=%%b

if "%DLL_PATH%"=="" (
    echo   [错误] 找不到 DLL 路径
    pause
    exit /b 1
)
echo   DLL = %DLL_PATH%

if not exist "%DLL_PATH%" (
    echo   [错误] DLL 文件不存在: %DLL_PATH%
    pause
    exit /b 1
)

:: 3. 生成新的 manifest 并复制 DLL 到输出目录
echo.
echo [3/3] 更新 manifest 并复制文件...

if not exist "%HOST_OUT%" mkdir "%HOST_OUT%"

:: 写入 manifest (CLSID 使用实际值)
> "%HOST_OUT%\VB6ComPlugin.manifest" (
    echo ^<?xml version="1.0" encoding="UTF-8" standalone="yes"?^>
    echo ^<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"^>
    echo   ^<assemblyIdentity type="win32" name="VB6ComPlugin" version="1.0.0.0" /^>
    echo   ^<file name="VB6ComPlugin.dll"^>
    echo     ^<comClass
    echo       clsid="%CLSID%"
    echo       threadingModel="Apartment"
    echo       progid="VB6ComPlugin.StringProcessor"
    echo       description="StringProcessor Plugin - VB6" /^>
    echo   ^</file^>
    echo ^</assembly^>
)

:: 复制 DLL
copy /Y "%DLL_PATH%" "%HOST_OUT%\VB6ComPlugin.dll" >nul

:: 同时更新源目录的 manifest
> "%VB6_SRC%\VB6ComPlugin.manifest" (
    echo ^<?xml version="1.0" encoding="UTF-8" standalone="yes"?^>
    echo ^<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"^>
    echo   ^<assemblyIdentity type="win32" name="VB6ComPlugin" version="1.0.0.0" /^>
    echo   ^<file name="VB6ComPlugin.dll"^>
    echo     ^<comClass
    echo       clsid="%CLSID%"
    echo       threadingModel="Apartment"
    echo       progid="VB6ComPlugin.StringProcessor"
    echo       description="StringProcessor Plugin - VB6" /^>
    echo   ^</file^>
    echo ^</assembly^>
)

echo.
echo ============================================
echo   完成!
echo   CLSID:     %CLSID%
echo   Manifest:  %HOST_OUT%\VB6ComPlugin.manifest
echo   DLL:       %HOST_OUT%\VB6ComPlugin.dll
echo.
echo   现在可以运行 ComPluginHost.exe
echo ============================================
pause
