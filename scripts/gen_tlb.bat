@echo off
set MIDL="C:\Program Files (x86)\Windows Kits\10\bin\x86\midl.exe"
set SDK_INC=C:\Program Files (x86)\Windows Kits\10\Include\10.0.26100.0
set IDL_DIR=%~dp0..\src\CSharpComPlugin
set OUT_DIR=%~dp0..\src\VB6ComPlugin

%MIDL% /tlb "%OUT_DIR%\CSharpComPlugin.tlb" /I "%SDK_INC%\um" /I "%SDK_INC%\shared" /no_cpp "%IDL_DIR%\CSharpComPlugin.idl"

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] MIDL failed with code %ERRORLEVEL%
    exit /b 1
)

echo [OK] Generated: %OUT_DIR%\CSharpComPlugin.tlb
dir "%OUT_DIR%\CSharpComPlugin.tlb"
