@echo off
chcp 65001 > nul
title Quan Ly Ho Kinh Doanh - Nhom 3

python --version > nul 2>&1
if errorlevel 1 (
    echo [LOI] Python chua duoc cai dat!
    pause
    exit /b 1
)

cd /d "%~dp0"
python app.py

if errorlevel 1 (
    echo.
    echo [LOI] Ung dung bi loi.
    pause
)
