# calcone-appads — Ebeveyn Ekran İzleme Aracı

Windows bilgisayarlarda, çocuğunuzun güvenliği için **belirli aralıklarla ekran
görüntüsü** alan, **şeffaf** (gizli olmayan) bir ebeveyn kontrol aracı.

Görüntüler yalnızca **o bilgisayardaki yerel diske** kaydedilir; internete veya
herhangi bir sunucuya **hiçbir veri gönderilmez**. İncelemeyi, oluşturulan basit
bir HTML galerisi üzerinden yaparsınız.

## Önemli: Etik ve yasal kullanım

Bu araç **kendi velayetiniz altındaki küçük çocuğunuzun** cihazını, onun
güvenliği için izlemek amacıyla tasarlanmıştır. Lütfen sorumlu kullanın:

- **Şeffaf olun.** Araç kendini gizlemez: açılışta bir bilgi penceresi gösterir
  ve sistem tepsisinde (saat yanında) bir simge bulundurur. Çocuğunuzu yaşına
  uygun bir şekilde bilgilendirmeniz hem güven hem de yasal açıdan doğru olan
  yaklaşımdır.
- **Yalnızca kendi cihazınızda / velisi olduğunuz çocuğun cihazında kullanın.**
  Rızası olmayan bir yetişkini (eş, arkadaş, çalışan vb.) gizlice izlemek birçok
  ülkede **suçtur**. Bu araç o amaç için tasarlanmamıştır.
- **Verileri koruyun.** Ekran görüntüleri özel bilgiler içerebilir. Kayıt
  klasörüne yalnızca sizin erişebildiğinizden emin olun ve `retention_days` ile
  eski kayıtların otomatik silinmesini sağlayın.

## Kurulum

Python 3.10+ gerekir.

```bash
pip install -r requirements.txt
```

`pystray` (tepsi simgesi) isteğe bağlıdır; kurulu değilse araç yine çalışır,
ancak şeffaflık için kurulması önerilir.

## Yapılandırma

Örnek dosyayı kopyalayıp düzenleyin:

```bash
copy config.example.json config.json   REM Windows
```

| Ayar | Açıklama | Varsayılan |
|------|----------|-----------|
| `interval_seconds` | Ekran görüntüleri arası süre (saniye) | `300` (5 dk) |
| `output_dir` | Görüntülerin kaydedileceği klasör | `screenshots` |
| `image_format` | `JPEG` (küçük boyut) veya `PNG` | `JPEG` |
| `jpeg_quality` | JPEG kalitesi (1–95) | `60` |
| `capture_all_monitors` | Tüm ekranları birlikte al | `true` |
| `retention_days` | Kaç günden eski kayıtlar silinsin (0 = silme) | `14` |
| `show_startup_notice` | Açılışta bilgi penceresi göster | `true` |
| `show_tray_icon` | Sistem tepsisi simgesi göster | `true` |
| `notice_text` | Bilgi penceresinde gösterilecek metin | — |

## Kullanım

```bash
# Sürekli izleme (Ctrl+C veya tepsi menüsünden durdurulur)
python -m screen_monitor run

# Tek seferlik tek bir ekran görüntüsü al
python -m screen_monitor once

# Kayıtları incelemek için HTML galeriyi yeniden oluştur
python -m screen_monitor gallery
```

Galeri `screenshots/gallery.html` dosyasında oluşturulur; tarayıcıda açarak
günlere göre gruplanmış görüntüleri inceleyebilirsiniz.

## Windows'ta oturum açılışında otomatik başlatma

Gizli bir servis **yerine**, Görev Zamanlayıcı ile normal bir görev
tanımlamanız önerilir (şeffaflığı korur):

1. **Görev Zamanlayıcı**'yı açın → **Temel Görev Oluştur**.
2. Tetikleyici: **Oturum açtığımda**.
3. Eylem: **Program başlat**
   - Program: `pythonw.exe` yolunuz (ör. `C:\Python312\pythonw.exe`)
   - Bağımsız değişkenler: `-m screen_monitor run`
   - Başlangıç konumu: bu projenin bulunduğu klasör.

## Proje yapısı

```
screen_monitor/
  config.py     # Ayarların okunması ve doğrulanması
  capture.py    # mss ile ekran görüntüsü alma
  storage.py    # Eski kayıtların otomatik temizlenmesi
  gallery.py    # İnceleme için HTML galeri üretimi
  notice.py     # Açılış bildirimi + sistem tepsisi simgesi (şeffaflık)
  monitor.py    # Ana döngü
  cli.py        # Komut satırı arayüzü
```
