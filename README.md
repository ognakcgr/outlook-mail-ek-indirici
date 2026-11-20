# 📧 Outlook Mail Arama ve Ek İndirme

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://microsoft.com/windows)
[![Version](https://img.shields.io/badge/Version-1.2.0-orange.svg)](https://github.com)

Outlook masaüstü uygulamasında **seçilen herhangi bir klasörde** (Gelen Kutusu, Gönderilenler, Arşiv, özel klasörler, paylaşılan kutular) filtreleme yaparak maillerdeki ekleri (Excel, PDF veya tümü) toplu olarak indiren modern, koyu temalı bir araç.

---

## ✨ Özellikler

| Özellik | Açıklama |
|---------|----------|
| 🔍 **Gelişmiş Filtreleme** | Konu, alıcı (To), gönderen adı, tarih aralığı |
| 📁 **Esnek Klasör Seçimi** | Outlook'taki **tüm** klasörler (ağaç yapısında) |
| 📎 **Ek Türü Filtresi** | Tümü / Excel (.xls, .xlsx, .csv) / PDF |
| 📂 **Otomatik Klasörleme** | `Tarih_Konu` formatında alt klasörler |
| 📊 **Canlı İlerleme** | İşlenen mail, ilerleme çubuğu, süre sayacı |
| 🎨 **Modern Koyu Tema** | Göz yormayan, profesyonel arayüz |
| ⚙️ **Ayarları Hatırlar** | Son klasör, filtreler, tercihler kaydedilir |

---

## 🖥️ Ekran Görüntüleri

> *Ekran görüntüleri eklenecek*

---

## 📋 Gereksinimler

- **Windows** (7/10/11)
- **Outlook Masaüstü Uygulaması** (açık ve oturum açmış) — *Web/OWA desteklenmez*
- **Python 3.8+**
- `pywin32` kütüphanesi

---

## 🚀 Kurulum

### Yöntem 1: Batch Dosyası ile (En Kolayı, Windows)

```cmd
# 1. Repoyu klonlayın
git clone https://github.com/ognakcgr/outlook-mail-ek-indirici.git
cd outlook-mail-ek-indirici

# 2. Çalıştırın (pywin32 otomatik kurulur)
outlook_ek_indir.bat
```

### Yöntem 2: Manuel Python

```bash
# 1. Klonla
git clone https://github.com/ognakcgr/outlook-mail-ek-indirici.git
cd outlook-mail-ek-indirici

# 2. Bağımlılıkları kur
pip install -r requirements.txt
# veya: pip install pywin32 Pillow

# 3. Çalıştır
python outlook_mail_ekleri_gui.py
```

---

## 🎯 Kullanım Kılavuzu

### 1️⃣ Outlook Klasörü Seçin
- **"Klasör Seç..."** butonuna tıklayın
- Ağaç yapısından istediğiniz klasörü seçin:
  - 📥 **Gelen Kutusu** (Inbox)
  - 📤 **Gönderilenler** (Sent Items)
  - 📦 **Arşiv** (Archive)
  - 🗂️ **Özel klasörleriniz** / alt klasörler
  - 👥 **Paylaşılan posta kutuları** (izin varsa)
- **"Seç"** butonuyla onaylayın

> ⚠️ **İlk çalıştırmada** Outlook güvenlik uyarısı çıkabilir: **"İzin Ver"** → *"Tüm bu tür işlemler için izin ver"* işaretleyin.

### 2️⃣ Filtreleri Ayarlayın

| Alan | Açıklama | Örnek |
|------|----------|-------|
| **Konu içeriyor** | Kısmi eşleşme (büyük/küçük harf duyarsız) | `fatura`, `rapor` |
| **Alıcı (To) içeriyor** | Alıcı e-posta/adı içinde arar | `muhasebe@sirket.com` |
| **Gönderen adı içeriyor** | Gönderen **görünen adı** içinde arar | `Yemekhane`, `Ogün Akciğer` |
| **Tarih Aralığı** | Gönderim/Alım tarihi (saat dahil) | `01.01.2024 00:00` - `31.12.2024 23:59` |

> **İpucu:** Filtreleri boş bırakırsanız sadece **tarih aralığı** uygulanır.

### 3️⃣ Ek Türü Seçin
- 📎 **Tüm Ekler** — Her dosya türü
- 📊 **Excel Dosyaları** — `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.csv`
- 📄 **PDF Dosyaları** — `.pdf`

### 4️⃣ İleri Ayarlar
- **Maksimum mail sayısı** — İşlenecek üst limit (varsayılan: 200)
- **Alt klasör oluştur** — Her mail için `YYYY-MM-DD_Konu` klasörü
- **İndirme sonrası aç** — İşlem bitince klasörü açar
- **Debug modu** — Konsola detaylı log basar

### 5️⃣ **"Ara ve İndir"** Butonuna Basın
- Canlı loglardan süreci takip edin
- İlerleme çubuğu ve işlenen mail konusu gösterilir
- Bitince **Klasörü Aç** butonu aktif olur

---

## 📁 Çıktı Yapısı

```
İndirme_Klasörü/
├── 2024-01-15_Gunluk_Yemek_Maliyeti/
│   ├── gunluk_yemek_ocak.xls
│   └── detay.pdf
├── 2024-01-14_Rapor_Talebi/
│   └── veri.xlsx
└── ...
```

---

## ⚙️ Ayarlar

Ayarlar `outlook_ayarlar.json` dosyasında **şifrelenmemiş** saklanır:
- Son indirme klasörü
- Filtre değerleri (konu, alıcı, gönderen, tarih)
- Seçili Outlook klasörü (EntryID + isim)
- Tercihler (ek türü, alt klasör, debug vb.)

---

## 🔧 Sorun Giderme

| Hata | Çözüm |
|------|-------|
| `Outlook'a bağlanılamadı` | Outlook masaüstü açık mı? Oturum açmış mı? |
| `pywin32 bulunamadı` | `pip install pywin32` |
| Klasör listesi boş | 1-2 sn bekleyin, Outlook "İzin Ver" uyarısını onaylayın |
| Mail bulunamıyor | Filtreleri genişletin / klasör doğru mu? |
| "Klasör bulunamadı" | Klasör taşınmış olabilir, tekrar seçin |
| Zaman aşımı | `max_mails` azaltın veya kodda `timeout` artırın |
| İzin hatası | Yönetici olarak çalıştırın |

---

## 📦 Dağıtım / EXE Oluşturma (İsteğe Bağlı)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico outlook_mail_ekleri_gui.py
# Çıktı: dist/outlook_mail_ekleri_gui.exe
```

---

## 🤝 Katkıda Bulunma

1. Fork yapın
2. Feature branch oluşturun (`git checkout -b feature/amazing-feature`)
3. Commit yapın (`git commit -m 'Add amazing feature'`)
4. Push yapın (`git push origin feature/amazing-feature`)
5. Pull Request açın

---

## 📄 Lisans

MIT License — Detaylar için [LICENSE](LICENSE) dosyasına bakın.

---

## 🙏 Teşekkürler

- [pywin32](https://github.com/mhammond/pywin32) — Windows COM API erişimi
- [Pillow](https://python-pillow.org/) — İkon oluşturma
- [tkinter](https://docs.python.org/3/library/tkinter.html) — GUI framework

---

**⭐ Bu projeyi faydalı bulduysanız yıldızlamayı unutmayın!**
