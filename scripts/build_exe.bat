@echo off
chcp 65001 > nul
echo Building MergeBooster.exe using csc.exe...

set CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe

if not exist "%CSC%" (
    echo Error: csc.exe not found at %CSC%
    exit /b 1
)

"%CSC%" /target:winexe /win32icon:app_icon.ico /out:MergeBooster.exe /r:System.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Core.dll src\cs\*.cs

if %ERRORLEVEL% equ 0 (
    echo Build successful! Created MergeBooster.exe
) else (
    echo Build failed with error code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
