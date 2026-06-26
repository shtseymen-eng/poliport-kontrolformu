@echo off
chcp 65001 >nul
title Poliport Arsivi - Kurulum

echo ============================================
echo  Poliport Kontrol Formu Arsivi - Kurulum
echo ============================================
echo.

:: Python kontrolu
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [HATA] Python bulunamadi!
    echo Lutfen https://www.python.org adresinden Python 3.10+ indirin.
    echo Kurulum sirasinda "Add Python to PATH" secenegini isaretleyin.
    pause
    exit /b 1
)

echo [1/3] Python bulundu.

:: Paketleri kur
echo [2/3] Gerekli paketler kuruluyor...
python -m pip install --upgrade pip -q
python -m pip install openpyxl requests requests-ntlm requests-negotiate-sspi -q

if %errorlevel% neq 0 (
    echo [UYARI] Bazi paketler kurulmadi. Program yine de calisabilir.
)

echo [3/3] Paketler kuruldu.
echo.
echo ============================================
echo  Uygulama baslatiliyor...
echo ============================================
echo.

python "%~dp0poliport_arsiv.py"

if %errorlevel% neq 0 (
    echo.
    echo [HATA] Uygulama baslatılamadi. Hata kodu: %errorlevel%
    pause
)
