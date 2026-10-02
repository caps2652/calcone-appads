# VB6 Canlı Ekran İzleme (OCX'siz)

Ev bilgisayarının ekranını (çocuk PC'si = **Server**) iş yerindeki bilgisayardan
(**Yönetici Panel = Client**) canlı izlemek için basit, hafif bir VB6 uygulaması.

- **OCX YOK.** Hiçbir `.ocx` bileşeni gerekmez (Winsock, GDI+, BitBlt hepsi
  doğrudan Windows API ile). Arkadaşın kendi PC'sinde VB6 ile sorunsuz derler.
- **Diske yazma YOK.** Çocuğun PC'sinde ekran görüntüsü dosyası oluşmaz; her kare
  sadece RAM'de JPEG'e sıkıştırılıp ağdan geçirilir.
- **Hafif ve hızlı.** Gri tonlama + küçültme + düşük JPEG kalitesi ile küçük
  kareler, canlı akış.

---

## 📁 Dosya yapısı

```
vb6-remote-monitor/
├─ server/                 → EV / ÇOCUK PC'de çalışır (ekranı yayınlar)
│   ├─ Server.vbp
│   ├─ frmServer.frm
│   ├─ modNet.bas          (Winsock API - ws2_32)
│   └─ modGraphics.bas     (ekran yakalama + JPEG - gdi32/gdiplus)
└─ client/                 → İŞ YERİ PC'de çalışır (ekranı gösterir)
    ├─ Client.vbp
    ├─ frmClient.frm
    ├─ modNet.bas          (server ile aynı)
    └─ modGraphics.bas     (server ile aynı)
```

> `modNet.bas` ve `modGraphics.bas` iki projede **birebir aynıdır**; her klasöre
> kendi kopyası konmuştur ki projeler bağımsız açılabilsin.

---

## 🔨 Derleme (VB6 kurulu olan PC'de)

> ⚠️ Derleme için **Visual Basic 6.0 IDE** gereklidir. İş yerindeki makinede VB6
> kurulu değilse, arkadaşın kendi PC'sinde derleyip çıkan `.exe`'leri sana verir.

Her iki proje için ayrı ayrı:

1. VB6 IDE'yi aç.
2. **File → Open Project** ile `server/Server.vbp` (veya `client/Client.vbp`) dosyasını aç.
3. Formdaki nesneler zaten yerleştirilmiş gelir (koordinatlar `.frm` içinde yazılı).
4. İstersen **F5** ile çalıştırıp dene.
5. **File → Make Server.exe** (veya **Make Client.exe**) ile `.exe` üret.

Her iki `.exe` için de aynı adımlar. İki ayrı exe elde edersin:
- `Server.exe` → ev/çocuk PC'sine
- `Client.exe` → iş yeri PC'sine

### Çalışma zamanı (runtime) notu
VB6 ile derlenen **her** program `MSVBVM60.DLL` (VB6 runtime) ister. Bu bir OCX
**değildir**; Windows 7/8/10/11'de genelde zaten yüklüdür. Nadiren eksikse,
Microsoft "Visual Basic 6.0 Runtime" paketi tek seferlik kurulur. Kullandığımız
diğer DLL'ler (`ws2_32`, `gdi32`, `gdiplus`, `ole32`, `user32`, `kernel32`)
Windows'un kendi parçasıdır, ayrıca kurulum gerektirmez.

---

## 🚀 Kullanım

### 1) Ev / Çocuk PC (Server.exe)
1. `Server.exe`'yi çalıştır.
2. **Dinleme portu**: `5900` (varsayılan).
3. **FPS**: saniyedeki kare (1-10). Düşük = daha az ağ/CPU. Öneri: **3**.
4. **Ölçek (%)**: ekranı küçültme oranı. **50** = yarı boyut (4 kat küçük veri).
5. **Kalite (1-100)**: JPEG kalitesi. **35** = küçük boyut, "kötü ama okunur".
6. **Gri tonlama**: işaretli = renksiz, daha küçük/hızlı.
7. **Baslat (Dinlemeye geç)** butonuna bas. Panel bağlantısını bekler.

### 2) İş Yeri PC (Client.exe)
1. `Client.exe`'yi çalıştır.
2. **Sunucu**: evin **dış IP adresi** veya DynDNS adresi (aşağıya bak).
3. **Port**: `5900` (server'daki ile aynı).
4. **Baglan** butonuna bas. Canlı görüntü pencerede akmaya başlar.

---

## 🌐 Port Yönlendirme (ÇOK ÖNEMLİ)

Ev ve iş yeri ayrı internet hatlarında olduğu için, iş yerindeki panelin evdeki
server'a ulaşabilmesi gerekir. Bunun için ev modeminde **port yönlendirme** yapılır:

1. Ev modeminin arayüzüne gir (genelde `192.168.1.1`).
2. **Port Forwarding / Sanal Sunucu / NAT** bölümünü bul.
3. Yeni kural ekle:
   - **Dış port (WAN)**: `5900`
   - **İç port (LAN)**: `5900`
   - **Hedef IP**: çocuğun PC'sinin yerel IP'si (örn. `192.168.1.10`)
     - Yerel IP'yi öğrenmek: çocuğun PC'sinde `cmd` → `ipconfig` → "IPv4 Address".
   - **Protokol**: TCP
4. Kaydet.

### Evin dış IP'sini öğrenme
Çocuğun PC'sinde tarayıcıdan `https://api.ipify.org` veya "my ip" araması.
İş yerinde panele bu IP'yi girersin.

### IP sürekli değişiyorsa (DynDNS)
Ev internetinin dış IP'si genelde dinamiktir (değişir). Her seferinde yeni IP
öğrenmemek için ücretsiz bir **DynDNS** servisi (No-IP, DuckDNS vb.) kurup
modeme tanıtabilirsin. Sonra panele IP yerine sabit adını yazarsın
(örn. `evim.duckdns.org`). Uygulama host adını otomatik çözer.

### Windows Güvenlik Duvarı
İlk çalıştırmada `Server.exe` için güvenlik duvarı izni çıkarsa **İzin ver**.
Çıkmazsa, Windows Defender Firewall → Gelen Kuralları → TCP **5900** portuna
izin ver.

---

## ⚙️ Nasıl çalışıyor (teknik özet)

```
[ÇOCUK PC - Server]                         [İŞ YERİ - Client/Panel]
 Timer (FPS)                                  Timer (poll)
   │                                            │
   ├─ BitBlt ile ekranı yakala                  │
   ├─ %50 küçült + gri (GDI+ ColorMatrix)       │
   ├─ RAM'de JPEG'e sıkıştır (GDI+)             │
   ├─ [4 byte uzunluk + JPEG] gönder ──TCP──►   ├─ veriyi oku, tamponda biriktir
   │                                            ├─ tam kareyi ayıkla
   │                                            ├─ JPEG'i çöz (GDI+)
   │                                            └─ PictureBox'a çiz
```

- **Protokol**: Her kare `4 byte küçük-endian uzunluk` + `JPEG verisi`.
- **Soketler**: `ws2_32.dll` API (OCX değil). Panel tarafı bloklamayan (non-blocking)
  + Timer ile yoklama → arayüz donmaz.
- **Tek panel**: Server aynı anda tek panele yayın yapar (parola yok).

---

## ⚠️ Güvenlik ve yasal notlar

- **Parola yok.** Port yönlendirme açıkken, evin dış IP/portunu bilen herkes
  bağlanıp ekranı görebilir. Riski azaltmak istersen ileride basit bir parola
  eklenebilir.
- Bu yazılım **kendi çocuğunun** bilgisayarını, ebeveyn gözetimi amacıyla
  izlemek için tasarlanmıştır. Başkalarının bilgisayarını onların bilgisi/izni
  olmadan izlemek yasa aykırı olabilir. Kullanım sorumluluğu kullanıcıya aittir.
- Sağlıklı bir yaklaşım: çocuğun yaşına uygun şekilde izlendiğinden haberdar olması.

---

## 🔧 Ayarlarla oynama (performans)

| Daha hızlı/küçük istiyorsan | Daha net görüntü istiyorsan |
|---|---|
| FPS'i düşür (1-2)           | FPS'i artır (5-8) |
| Ölçeği düşür (%30-40)       | Ölçeği artır (%70-100) |
| Kaliteyi düşür (20-30)      | Kaliteyi artır (60-80) |
| Gri tonlamayı aç            | Gri tonlamayı kapat (renkli) |

Renkli görüntü istiyorsan server'da **Gri tonlama** kutusunun işaretini kaldır.
