@echo off
chcp 65001 > nul
echo ========================================================
echo   Excel Diff Aggregator - PDF Manual Generation Script
echo ========================================================
echo.

set EDGE_PATH="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

if not exist %EDGE_PATH% (
    set EDGE_PATH="C:\Program Files\Microsoft\Edge\Application\msedge.exe"
)

echo [1/2] Generating Quick Start Guide PDF (1 Page)...
%EDGE_PATH% --headless --print-to-pdf="%~dp0..\ExcelDiffAggregator_QuickGuide.pdf" --no-pdf-header-footer --print-to-pdf-no-header "%~dp0..\docs\manuals\quick_guide.html"
if %errorlevel% equ 0 (
    echo     - SUCCESS: ExcelDiffAggregator_QuickGuide.pdf created!
) else (
    echo     - ERROR: Failed to generate Quick Start Guide PDF!
)

echo.
echo [2/2] Generating Standard User Manual PDF (Multi-Page)...
%EDGE_PATH% --headless --print-to-pdf="%~dp0..\ExcelDiffAggregator_UserManual.pdf" --no-pdf-header-footer --print-to-pdf-no-header "%~dp0..\docs\manuals\user_manual.html"
if %errorlevel% equ 0 (
    echo     - SUCCESS: ExcelDiffAggregator_UserManual.pdf created!
) else (
    echo     - ERROR: Failed to generate Standard User Manual PDF!
)

echo.
echo ========================================================
echo   PDF Generation Completed!
echo ========================================================
