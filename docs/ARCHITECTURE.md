# Remote Support — Mimari ve Geliştirme Yol Haritası

VB6 ile yazılmış, OCX/ActiveX kullanmayan, Windows API tabanlı modern bir
**LAN uzaktan destek** uygulaması. Bağlantılar **onay tabanlıdır**: host
tarafı gelen bağlantıyı görür ve (PHASE 2'den itibaren) açıkça izin verir;
kimlik doğrulama olmadan uzaktan kontrol başlamaz.

> **Durum:** PHASE 1 tamamlandı — Modern ana arayüz + LAN bağlantı altyapısı.

---

## Derleme ve çalıştırma

> Bu depo Linux tabanlı bir ortamda hazırlandı; **VB6 derleyicisi burada
> çalışmaz.** Aşağıdaki adımlar Windows + Visual Basic 6 IDE içindir.

1. `src\RemoteSupport.vbp` dosyasını VB6 IDE ile açın.
2. **Run ▶ (F5)** ile çalıştırın veya **File → Make RemoteSupport.exe** ile derleyin.
3. Ek bir runtime/DLL gerekmez (yalnızca her Windows'ta bulunan `msvbvm60.dll`).

### LAN testi (PHASE 1 el sıkışması)
1. İki bilgisayarda (veya aynı makinede iki örnekte) uygulamayı çalıştırın.
2. Her örnek açılışta `7070` portunu dinler (ayarlanabilir).
3. Bir örnekte **Partner ID / IP** alanına diğerinin IP'sini girin:
   `192.168.1.50` veya `192.168.1.50:7070`.
4. **BAĞLAN**'a basın. Durum sırasıyla *Bağlanıyor → Kimlik doğrulanıyor →
   Bağlandı* olur ve `HELLO/WELCOME` el sıkışması tamamlanır.
5. `Logs\app.log` teknik akışı gösterir (şifreler loglanmaz).

> VB6 IDE'de çalışırken **Stop** düğmesiyle durdurmak subclass'ı geri
> almadan kapatabilir. Temiz çıkış için pencereyi kapatın (X). Derlenmiş
> EXE'de bu sorun yoktur.

---

## Modül yapısı

| Dosya | Sorumluluk |
|-------|-----------|
| `modMain.bas` | `Sub Main` giriş noktası, uygulama sabitleri, tekil nesneler, düzenli kapanış |
| `modWindowsAPI.bas` | Tüm API bildirimleri ve paylaşılan `Type`/`Const` tanımları (tek yerde) |
| `modLogging.bas` | `Logs\app.log` — seviyeli, sır içermeyen log + rollover |
| `modConfig.bas` | `config.ini` oku/yaz (Unicode `W` API) — tema, port, timeout |
| `modSecurity.bas` | `CryptGenRandom` ile ID/şifre üretimi, SHA-256 (Phase 2 auth için) |
| `modUI.bas` | GDI çizim yardımcıları, DPI ölçekleme, font önbelleği (GDI leak yok) |
| `modNetwork.bas` | Winsock çekirdeği, message-only pencere + subclass, paket çerçeveleme, UTF-8 |
| `clsTheme.cls` | Windows 11 tarzı dark/light renk paleti (sistem teması algılama) |
| `clsPacket.cls` | Tek paketi temsil eden veri taşıyıcısı (payload yardımcıları) |
| `clsConnection.cls` | Asenkron TCP socket sarmalayıcısı (taşma korumalı parser, tx kuyruğu) |
| `clsSession.cls` | Host/Viewer orkestrasyonu, yerel kimlik, durum makinesi |
| `frmMain.frm` | Tamamen GDI ile çizilen modern ana pencere (owner-drawn butonlar) |

Katmanlar ayrıktır; ağ katmanı ileride **Internet Relay / Cloud Server /
NAT Traversal** eklenebilecek şekilde izole edilmiştir.

---

## Ağ protokolü

Her paket 16 baytlık sabit başlık + payload'dan oluşur (little-endian):

```
[0..1]   magic "RS" (0x52,0x53)
[2]      version
[3]      flags
[4..7]   type        (PacketType)
[8..11]  sessionId
[12..15] payloadLen  (0 .. 16 MB arası doğrulanır — taşma koruması)
```

Paket türleri (`modNetwork.PacketType`): `HELLO, WELCOME, AUTH_*,
PERMISSION_*, SCREEN_FRAME, MOUSE_*, KEYBOARD, CLIPBOARD, FILE_*, CHAT,
PING, PONG, DISCONNECT`. PHASE 1'de yalnızca `HELLO/WELCOME` ve `PING/PONG`
etkin; diğerleri sonraki fazlarda işlenecek.

### Asenkron model
`WSAAsyncSelect` socket olaylarını gizli bir **message-only** pencereye
yollar; pencere subclass edilir ve olaylar `clsConnection`'a dağıtılır.
Socket'ler non-blocking'dir, bu yüzden **UI thread'i ağ yüzünden donmaz**.
Gönderimde kısmi `send` sonuçları kuyruklanır ve `FD_WRITE`'ta boşaltılır.

---

## Güvenlik ilkeleri (uygulanmakta)

- **Erişim şifresi opsiyoneldir ve varsayılan olarak KAPALIDIR.** Ana
  ekrandaki "Erişim şifresi iste" anahtarıyla açılır/kapanır
  (`Security/RequirePassword`, `config.ini`).
  - **Kapalı (varsayılan) = katılımlı erişim:** şifre üretilmez; her gelen
    bağlantı host tarafından **Allow/Deny** ile onaylanır (PHASE 2).
  - **Açık = gözetimsiz erişim:** `CryptGenRandom` ile şifre üretilir; karşı
    taraf şifreyle (kimse onay vermeden) bağlanabilir — kendi bilgisayarlarınız
    için uygundur.
- Şifreler **düz metin saklanmaz** ve **loglanmaz**; referans olarak SHA-256
  hash tutulur. Şifre kapatılınca bellekten temizlenir.
- Bağlantı ID `CryptGenRandom` ile üretilir.
- Paket parser uzunluk sınırı uygular (bellek taşması/DoS koruması).
- Bağlantı kapanışında socket/kaynaklar düzgün serbest bırakılır.
- (PHASE 2) Host tarafında **Allow/Deny** onayı; şifre açıksa ayrıca
  challenge-response doğrulama.

---

## Faz yol haritası

| Faz | Kapsam | Durum |
|-----|--------|-------|
| **1** | Modern ana UI + LAN bağlantı altyapısı (handshake) | ✅ Tamamlandı |
| 2 | Host/Viewer kimlik doğrulama + Allow/Deny izni | ⏳ Sıradaki |
| 3 | Ekran yakalama + sıkıştırma + aktarım (değişen bölgeler) | ⏳ |
| 4 | Uzak fare kontrolü | ⏳ |
| 5 | Uzak klavye kontrolü | ⏳ |
| 6 | Tam ekran + kalite + FPS kontrolleri, floating toolbar | ⏳ |
| 7 | Dosya transferi (ilerleme çubuğu, iptal) | ⏳ |
| 8 | Pano paylaşımı + chat | ⏳ |
| 9 | Ayarlar penceresi + bağlantı geçmişi | ⏳ |
| 10 | Performans optimizasyonu + kararlılık testleri | ⏳ |

Her faz, önceki fazların çalışan özelliklerini bozmadan eklenir.

### Bilinen PHASE 1 sınırları
- **ID ile bağlanma** için LAN keşfi (UDP broadcast) henüz yok; şimdilik
  `IP` veya `IP:Port` ile bağlanılır (ID görüntü/kimlik amaçlı gösterilir).
- Ekran/giriş/dosya/chat özellikleri sonraki fazlarda gelecek.
