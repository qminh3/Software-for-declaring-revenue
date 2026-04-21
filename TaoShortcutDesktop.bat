@echo off
chcp 65001 > nul
title Tao Shortcut Desktop

set APP_DIR=%~dp0
set APP_DIR=%APP_DIR:~0,-1%
set LNK_NAME=HKD Nhom 3

:: Tim pythonw.exe (chay khong hien console)
for /f "delims=" %%i in ('python -c "import sys,os; print(sys.executable.replace(chr(39),''))"') do set PY_EXE=%%i
set PYW_EXE=%PY_EXE:python.exe=pythonw.exe%
if not exist "%PYW_EXE%" set PYW_EXE=%PY_EXE%

:: Desktop path
for /f "delims=" %%i in ('powershell -command "[Environment]::GetFolderPath(\"Desktop\")"') do set DESKTOP=%%i

set LNK_PATH=%DESKTOP%\%LNK_NAME%.lnk

echo App dir  : %APP_DIR%
echo Python   : %PYW_EXE%
echo Shortcut : %LNK_PATH%
echo.

:: Ghi PowerShell script ra file ASCII (tranh loi encoding)
set PS_TMP=%TEMP%\make_hkd_shortcut.ps1
(
echo $ws = New-Object -ComObject WScript.Shell
echo $s = $ws.CreateShortcut('%LNK_PATH%'^)
echo $s.TargetPath = '%PYW_EXE%'
echo $s.Arguments = '"%APP_DIR%\app.py"'
echo $s.WorkingDirectory = '%APP_DIR%'
echo $s.IconLocation = '%APP_DIR%\icon.ico'
echo $s.Description = 'Quan Ly Ho Kinh Doanh'
echo $s.WindowStyle = 1
echo $s.Save(^)
echo Write-Host 'DONE'
) > "%PS_TMP%"

powershell -ExecutionPolicy Bypass -File "%PS_TMP%"
del "%PS_TMP%" 2>nul

if exist "%LNK_PATH%" (
    echo.
    echo [THANH CONG] Shortcut da duoc tao:
    echo   %LNK_PATH%
    echo.
    echo Double-click vao icon "Quan Ly Ho Kinh Doanh" ngoai Desktop de mo phan mem!
) else (
    echo.
    echo [LOI] Khong tao duoc shortcut. Thu chay TaoShortcut.py bang Python.
)
echo.
pause
