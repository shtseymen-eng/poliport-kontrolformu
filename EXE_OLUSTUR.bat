@echo off
chcp 65001 >nul
title Poliport Arsivi - EXE Olustur

echo ============================================
echo  Poliport Arsivi - Tek Dosya EXE Olustur
echo ============================================
echo.

python -m pip install pyinstaller pillow -q

echo PyInstaller ile derleniyor...
echo (Bu islem 1-2 dakika surebilir)
echo.

:: icon.png varsa ICO'ya cevir
set ICON_ARG=--icon NONE
if exist "%~dp0icon.png" (
    python -c "from PIL import Image; img=Image.open(r'%~dp0icon.png'); img.save(r'%~dp0icon.ico', sizes=[(256,256),(64,64),(32,32),(16,16)])" 2>nul
)
if exist "%~dp0icon.ico" set ICON_ARG=--icon "%~dp0icon.ico"

pyinstaller ^
    --onefile ^
    --windowed ^
    --name "PoliportArsivi" ^
    %ICON_ARG% ^
    --add-data "%~dp0icon.png;." ^
    --hidden-import requests_ntlm ^
    --hidden-import requests_negotiate_sspi ^
    --hidden-import openpyxl ^
    --hidden-import PIL ^
    --hidden-import PIL.Image ^
    "%~dp0poliport_arsiv.py"

if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo  BASARILI!
    echo  EXE dosyasi: dist\PoliportArsivi.exe
    echo ============================================
    echo.
    explorer "%~dp0dist"
) else (
    echo.
    echo [HATA] Derleme basarisiz oldu.
    pause
)
