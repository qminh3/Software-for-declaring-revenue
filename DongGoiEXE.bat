@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
title Dong Goi Phan Mem - HKD Nhom 3 - Tao File EXE

echo ============================================================
echo   DONG GOI PHAN MEM - TAO FILE .EXE CHAY KHONG CAN PYTHON
echo   HKD Nhom 3 - Quan Ly Ho Kinh Doanh
echo ============================================================
echo.

:: --- Kiem tra Python ---
python --version > nul 2>&1
if errorlevel 1 (
    echo [LOI] Python chua duoc cai dat!
    echo Tai Python tai: https://www.python.org/downloads/
    pause & exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PY_VER=%%v
echo [OK] !PY_VER!

:: --- Cai PyInstaller neu chua co ---
python -c "import PyInstaller" > nul 2>&1
if errorlevel 1 (
    echo [..] Dang cai PyInstaller...
    pip install pyinstaller pillow openpyxl --quiet
    if errorlevel 1 ( echo [LOI] Khong cai duoc! Kiem tra mang. & pause & exit /b 1 )
    echo [OK] Da cai PyInstaller
) else (
    echo [OK] PyInstaller san co
)
pip install pillow openpyxl --quiet --exists-action i
echo [OK] Thu vien OK

:: --- Lay duong dan tuyet doi (xu ly khoang trang trong ten thu muc) ---
set "APP_DIR=%~dp0"
if "!APP_DIR:~-1!"=="\" set "APP_DIR=!APP_DIR:~0,-1!"
echo [..] Thu muc: !APP_DIR!

:: --- Don dep build cu ---
if exist "!APP_DIR!\build_temp" rmdir /s /q "!APP_DIR!\build_temp"
if exist "!APP_DIR!\dist"       rmdir /s /q "!APP_DIR!\dist"

:: --- Tao file spec de tranh van de khoang trang trong duong dan ---
set "SPEC_FILE=!APP_DIR!\build_spec.spec"
(
echo # -*- mode: python ; coding: utf-8 -*-
echo block_cipher = None
echo a = Analysis^(
echo     [r'!APP_DIR!\app.py'],
echo     pathex=[r'!APP_DIR!'],
echo     binaries=[],
echo     datas=[
echo         ^(r'!APP_DIR!\icon.png', '.'^),
echo         ^(r'!APP_DIR!\icon.ico', '.'^),
echo         ^(r'!APP_DIR!\tabs', 'tabs'^),
echo     ],
echo     hiddenimports=['tkinter','tkinter.ttk','tkinter.messagebox',
echo         'tkinter.filedialog','openpyxl','openpyxl.styles',
echo         'openpyxl.utils','sqlite3','excel_export_nhom3'],
echo     hookspath=[],
echo     runtime_hooks=[],
echo     excludes=['matplotlib','numpy','pandas','scipy'],
echo     cipher=block_cipher,
echo     noarchive=False,
echo ^)
echo pyz = PYZ^(a.pure, a.zipped_data, cipher=block_cipher^)
echo exe = EXE^(
echo     pyz, a.scripts, a.binaries, a.zipfiles, a.datas, [],
echo     name='HKD_Nhom3',
echo     debug=False, strip=False, upx=True,
echo     upx_exclude=[], runtime_tmpdir=None,
echo     console=False,
echo     icon=r'!APP_DIR!\icon.ico',
echo ^)
) > "!SPEC_FILE!"

echo [..] Da tao spec file OK
echo.
echo [..] Bat dau dong goi... (co the mat 3-6 phut)
echo.

:: --- Chay PyInstaller voi spec file ---
python -m PyInstaller ^
    --distpath "!APP_DIR!\dist" ^
    --workpath "!APP_DIR!\build_temp" ^
    "!SPEC_FILE!"

if errorlevel 1 (
    echo.
    echo [LOI] Dong goi that bai!
    echo Thu click phai file nay -^> Run as administrator
    pause & exit /b 1
)

echo.
echo ============================================================
if exist "!APP_DIR!\dist\HKD_Nhom3.exe" (
    echo   THANH CONG^^!
    echo ============================================================
    echo.
    echo [FILE] !APP_DIR!\dist\HKD_Nhom3.exe
    for %%f in ("!APP_DIR!\dist\HKD_Nhom3.exe") do (
        set /a SIZE=%%~zf/1024/1024
        echo [KICH THUOC] %%~zf bytes ^(khoang !SIZE! MB^)
    )
    echo.
    echo San sang copy sang may khac chay khong can Python^^!
    echo.
    set /p ANS="Mo thu muc dist ngay bay gio? (Y/N): "
    if /i "!ANS!"=="Y" explorer "!APP_DIR!\dist"
) else (
    echo [LOI] Khong tim thay HKD_Nhom3.exe
)

pause
