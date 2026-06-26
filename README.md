# Poliport Kontrol Formu Arşivi

Poliport Terminal araç kontrol formlarını otomatik indirip arşivleyen Windows masaüstü uygulaması.

---

## ⬇️ İndir & Kur

**[Releases sayfasına git →](../../releases/latest)**

Son sürümden `PoliportArsiv.exe` dosyasını indir, çalıştır — kurulum gerekmez.

---

## Özellikler

- Aylık Excel raporundan (Arac_Kayit_Sorgulama.xlsx) Poliport Terminal araçlarını okur
- SSRS kontrol formlarını otomatik indirir
- Windows NTLM kimlik doğrulama (domain şifresiyle çalışır)
- Dosyaları `Kontrol Formları / YIL / AY / GÜN / PLAKA DORSE KONTEYNER AD SOYAD NAKLİYECİ.xlsx` yapısında kaydeder

---

## Geliştirici — Kaynak Koddan Çalıştırma

```bash
pip install -r requirements.txt
python poliport_arsiv.py
```

.exe derlemek için:
```
EXE_OLUSTUR.bat
```

---

*S. SEYMEN tarafından hazırlanmıştır*
