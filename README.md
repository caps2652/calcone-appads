# Remote Support

VB6 ile yazılmış, **OCX/ActiveX kullanmayan**, Windows API/GDI tabanlı modern
bir **LAN uzaktan destek** uygulaması (TeamViewer/AnyDesk kullanım mantığı).
Öncelikli hedef: kendi bilgisayarlarınız ve **açıkça izin verdiğiniz**
bilgisayarlar arasında uzaktan destek.

> Bağlantılar onay tabanlıdır: host tarafı gelen bağlantıyı görür, kimlik
> doğrulama olmadan uzaktan kontrol başlamaz, şifreler düz metin saklanmaz
> veya loglanmaz.

## Durum

**PHASE 1 tamamlandı** — Modern (Windows 11 tarzı, dark/light) ana arayüz +
LAN bağlantı altyapısı (ham Winsock, asenkron, kendi paket protokolü,
HELLO/WELCOME el sıkışması).

Tüm faz yol haritası, mimari ve derleme talimatları için bkz.
[`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md).

## Hızlı başlangıç (Windows + VB6 IDE)

1. `src\RemoteSupport.vbp` dosyasını Visual Basic 6 ile açın.
2. **F5** ile çalıştırın veya **Make RemoteSupport.exe** ile derleyin.
3. İki makinede çalıştırıp birinde diğerinin IP'sini (`192.168.1.50:7070`)
   girerek **BAĞLAN**'a basın.

Ek runtime/DLL gerekmez (yalnızca Windows'ta hazır bulunan `msvbvm60.dll`).

## Teknik özellikler

- Saf VB6, 32-bit, OCX'siz; yalnızca yerleşik Windows API.
- Unicode için `W` (wide) API fonksiyonları.
- Tamamen GDI ile çizilen modern arayüz (owner-drawn butonlar, DPI ölçekleme,
  yuvarlatılmış köşeler, hover/pressed durumları, toast bildirimleri).
- Asenkron socket (UI donmaz), taşma korumalı paket çözümleyici.
- GDI/handle sızıntısı bırakmayan kaynak yönetimi.

## Dizin yapısı

```
src/     VB6 proje ve kaynak dosyaları (.vbp .frm .bas .cls)
docs/    Mimari ve yol haritası
```

> Bu depo Linux ortamında hazırlandığından burada VB6 ile derlenemez;
> derleme/çalıştırma Windows üzerinde yapılır.
