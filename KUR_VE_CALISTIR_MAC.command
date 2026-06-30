#!/bin/bash
# Poliport Kontrol Formu Arsivi - Mac Baslatici
# Bu dosyaya cift tikla - uygulama kurulur ve acilir

cd "$(dirname "$0")"

echo "================================"
echo " Poliport Kontrol Formu Arsivi"
echo " Mac Baslatici"
echo "================================"
echo ""

# Python3 var mi?
if ! command -v python3 &>/dev/null; then
    osascript -e 'display alert "Python3 Gerekli" message "Python3 bulunamadi. https://python.org adresinden indirin." as critical'
    exit 1
fi

echo "Bagimliliklar yukleniyor..."
python3 -m pip install --upgrade pip -q
python3 -m pip install openpyxl requests requests-ntlm pillow -q 2>/dev/null || \
python3 -m pip install openpyxl requests requests-ntlm pillow --break-system-packages -q

echo "Program baslatiliyor..."
python3 poliport_arsiv.py
