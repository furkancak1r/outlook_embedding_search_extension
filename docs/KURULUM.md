# HatırMail – Kurulum ve Kaldırma Kılavuzu

HatırMail, Klasik Outlook (Windows masaüstü) için bir COM eklentisidir. Hatırladığınız birkaç ayrıntıyı
Türkçe ya da İngilizce yazarsınız; eklenti ilgili e-postaları bulur ve Outlook'ta açar.
Yapay zekâ modeli yalnız sizin bilgisayarınızda çalışır. Postalarınız ve sorgularınız hiçbir sunucuya gönderilmez.

## Gereksinimler

| Gereksinim | Değer |
|---|---|
| İşletim sistemi | Windows 10 veya 11, 64 bit (x64) |
| Outlook | Klasik Outlook. Test edilen: Outlook 2016 64 bit (tür kitaplığı ve kayıt düzeyinde). Eklenti AnyCPU derlenir; 32 bit Outlook ve Microsoft 365 sürümleri doğrulanmadı. Yeni Outlook (New Outlook) desteklenmez. |
| .NET Framework | 4.8 (Windows 10/11 ile gelir) |
| Visual C++ çalışma kitaplığı | Microsoft Visual C++ 2015–2022 Redistributable (x64). Yapay zekâ kitaplığı (ONNX Runtime) bunu ister; çoğu Windows'ta yüklüdür. Test makinesinde 14.42 sürümü vardı. |
| Disk | Kurulum için yaklaşık 1,8 GB (paket 544 MB + kurulum klasörü 1,25 GB). İndeks: 316 sentetik posta için 1,1 MB ölçüldü; gerçek postalar daha uzun olduğu için 10.000 posta başına 40–80 MB tahmin edilir. |
| Bellek | Pencere açıkken arama işlemi yaklaşık 0,7 GB, indeksleme sırasında en çok yaklaşık 1,3 GB kullanır (2.19 ölçümü). Pencere kapanınca ya da "Yapay zekâyı hazır tut" süresi dolunca bu bellek boşalır |
| Yetki | Yönetici yetkisi gerekmez. Kurulum yalnız sizin kullanıcı hesabınızı etkiler. |

## Kurulum

> **Dikkat:** Kayıt adımından önce Outlook'u kapatın. Outlook açıksa kurulum betiği durur ve hiçbir kayıt yazmaz.

1. `HatirMail-win-x64-0.1.0.zip` dosyasını bir klasöre açın.
2. İsterseniz paketin bütünlüğünü denetleyin: `MANIFEST.sha256` dosyasındaki her satır bir dosyanın SHA-256 değeridir.
3. Outlook'u kapatın.
4. Açtığınız klasörde PowerShell'i başlatın ve şu komutu çalıştırın:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
   ```

5. Outlook'u açın. **Giriş** sekmesinde **HatırMail ile Ara** düğmesi görünür.

Kurulum betiği şunları yapar:

- Dosyaları `%LOCALAPPDATA%\HatirMail` klasörüne kopyalar.
- Gömülü Python'u bu klasöre açar ve bağımlılıkları internetsiz, SHA-256 denetimiyle kurar.
- Model dosyalarının SHA-256 değerlerini denetler.
- Eklentiyi yalnız geçerli kullanıcı için kaydeder (`HKCU\Software\Classes` ve
  `HKCU\Software\Microsoft\Office\Outlook\Addins\HatirMail.Connect`).

Kurulum betiği PATH, hizmet, zamanlanmış görev ya da Outlook profili değiştirmez. Sistem kayıt defterine (HKLM) yalnız
`-Machine` seçeneğiyle yazar (aşağıya bakın).

### UAC kapalı ve hesap yönetici ise

Bu durumda Outlook tam yönetici yetkisiyle çalışır. Windows COM, böyle süreçlerde kullanıcı başına (HKCU) sınıf kaydını
yok sayar ve eklenti yüklenmez. Kurulum betiği bu durumu "COM activation check" adımında algılar ve uyarı verir.
Hiçbir şeyi değiştirmeden yalnız denetlemek için: `install.ps1 -CheckOnly`.

İki çözüm vardır:

1. Outlook'u standart kullanıcı olarak çalıştırın (UAC açık). Kayıt defterine ek yazma gerekmez.
2. COM sınıfını makine düzeyinde kaydedin (aşağıdaki komut).

> **Dikkat:** Aşağıdaki komut sistem kayıt defterine (HKLM) yazar ve yönetici yetkisi ister.

Paketi açtığınız klasörde, yönetici PowerShell'inde:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -RegisterOnly -Machine
```

Kaldırırken `uninstall.ps1 -Machine` kullanın.

### Kayıt yazmadan önce inceleme

Kayıt defterine yazılacak değerleri önce görmek için:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -RegFile .\hatirmail-kayit.reg
```

Bu komut dosyaları kurar, kayıt değerlerini `.reg` dosyasına yazar ve kayıt defterini değiştirmez.

### Yalnız dosya kurulumu (test için)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -InstallDir C:\Test\HatirMail -SkipRegistration
```

## Kullanım

Pencerede üç sekme vardır: **Ara**, **İndeks** ve **Ayarlar**.

### İlk kullanım: indeksi oluşturun

1. **HatırMail ile Ara** düğmesine basın.
2. **İndeks** sekmesini açın. Posta kutusunu ve klasörleri seçin. Silinmiş Öğeler, Önemsiz E-Posta, Taslaklar ve Giden
   Kutusu varsayılan olarak seçili değildir.
3. **İndekslemeyi başlat**'a basın. Çubuklar okunan postaların, anlam indeksinin ve eklerin yüzdesini gösterir. Altında o
   anki aşama ve o aşamada kalan süre yazar.
4. İndeksleme sürerken Outlook'u kullanabilirsiniz. Pencereyi simge durumuna küçültebilirsiniz.

> **Not:** Pencereyi kapatırsanız ya da Outlook'u kapatırsanız indeksleme durur. Yapılan iş kaydedilir. Sonra
> **İndekslemeyi başlat**'a yeniden basın; iş kaldığı yerden sürer.

Büyük bir posta kutusunda tam indeksleme saatler sürebilir. 2.19 test makinesinde (Xeon E5-2620 v3, 4 iş parçacığı) birkaç
bin postalık gerçek bir PST'nin ilk indeksi 2 saatten uzun sürdü. Bu süre bir kez harcanır; sonraki çalışmalar yalnız yeni ve
değişen postaları işler.

### Arama

1. **Ara** sekmesinde hatırladıklarınızı yazın. Örnekler: `geçen ay Ahmet'ten gelen kargo faturası`,
   `mart başında gelen raf teklifi`, `mail from Sarah about the delayed shipment`. Ekin içinde geçen bir şeyi de
   yazabilirsiniz: `galvanizli raf dikmesi sayımı` (Excel), `soğuk zincir taşıma faturası` (PDF), `E-4012 hata kodu`
   (ekran görüntüsü).
2. **Ara**'ya basın. Önce kelime eşleşmeleri, sonra anlamsal sonuçlar görünür.
   Pencere açılınca yapay zekâ modeli arka planda hazırlanır (durum çubuğunda "Yapay zekâ hazır" yazar). Pencere açık
   kaldıkça sonraki aramalar modeli beklemez.
3. Bir sonuç seçin. Alttaki önizlemede konu, kişi, tarih ve eşleşen bölüm görünür; aranan sözcükler sarı zeminle
   işaretlenir. Eşleşme bir ekin içindeyse "Ekte bulundu: dosya adı" yazar.
4. Sonuca çift tıklayın ya da **Outlook'ta Aç**'a basın.

Arama, yeni postaları indekslemek için en çok **Aramada indeksleme** süresi kadar bekler (varsayılan 20 saniye). Kalan iş
sonraki aramaya ya da İndeks sekmesine kalır. Arama kutusunun altındaki satır, aranan klasörlerin indeks durumunu gösterir.
İndeks eksikse sonuçlar indeksteki postalarla sınırlıdır. Anlam indeksi hiç yoksa yalnız kelime eşleşmeleri gösterilir.

İndeksleme sürerken **Ara**'ya basarsanız HatırMail önce indekslemeyi durdurmak için onay ister. Aynı anda yalnız bir
arama işlemi çalışır.

### Ek içeriği

HatırMail eklerin içindeki metni de arar:

| Ek türü | Okunur mu |
|---|---|
| Excel (xlsx, xlsm, xls), Word (docx), PowerPoint (pptx), OpenDocument (odt, ods, odp) | Evet |
| PDF | Evet. Taranmış (metni olmayan) sayfalar OCR ile okunur, dosya başına en çok 3 sayfa. |
| Resim (png, jpg, bmp, gif, tif) | Evet, Windows OCR ile. Küçük imza logoları okunmaz. |
| Eklenmiş e-posta (msg), zip | Evet (zip içindeki desteklenen dosyalar) |
| txt, csv, html, xml | Evet |
| rar, 7z, eski doc/ppt, video, ses, CAD | Hayır, yalnız dosya adı aranır |

- Ekler indekslemede bir kez okunur. Sonraki çalışmalar yalnız yeni ekleri okur. Aynı dosya birden çok postada varsa bir
  kez okunur.
- 20 MB'tan büyük ekler, parola korumalı dosyalar ve bozuk dosyalar okunmaz. Posta yine bulunabilir.
- OCR, Windows'un kendi OCR özelliğidir. Windows'ta yüklü dil paketlerinin OCR dillerini kullanır (Türkçe Windows'ta
  Türkçe). İndeks sekmesinde "Ek dosyaları" çubuğu ve klasör listesindeki "Ekler" sütunu okunan/toplam eki gösterir.
- Ek okumayı kapatmak için **Ayarlar** sekmesinde "Ek dosyalarının içinde de ara" kutusunun işaretini kaldırıp
  **Kaydet**'e basın.

Sonuçlardaki **Güven** sütunu:

| Değer | Anlamı |
|---|---|
| Yüksek | Sorguyla güçlü anlam ya da kelime eşleşmesi var |
| Orta | Kısmi eşleşme var; içeriği kontrol edin |
| Zayıf (gri) | Yalnız uzak benzerlik var; büyük olasılıkla aradığınız posta değil |
| Kelime | Anlamsal arama bitmeden önce gösterilen tam kelime eşleşmesi |

## Ayarlar

**Ayarlar** sekmesinde değeri değiştirin ve **Kaydet**'e basın. Değerler `%LOCALAPPDATA%\HatirMail\settings.json`
dosyasına yazılır ve bir sonraki aramada ya da indekslemede geçerli olur. **Varsayılan değerler** düğmesi alanları
varsayılana döndürür (kaydetmek için yine **Kaydet**'e basın).

| Sekmedeki ad | `settings.json` | Varsayılan | Aralık | Anlamı |
|---|---|---|---|---|
| İşlemci çekirdeği | `threads` | 4 | 1–16 | Modelin kullanacağı en çok CPU çekirdeği. Çok çekirdek indekslemeyi hızlandırır ama Outlook'u yavaşlatabilir |
| Bellek sınırı (MB) | `memory_limit_mb` | 4096 | 2048–32768 | Arama işleminin bellek sınırı (Windows Job Object). Modelle indeksleme en çok yaklaşık 1.300 MB kullanır (2.19 ölçümü) |
| Aramada indeksleme (sn) | `search_index_budget_s` | 20 | 0–600 | Ara'ya basınca yeni postaları indekslemek için en çok bu kadar beklenir. 0: aramada indeksleme yok |
| Yapay zekâyı hazır tut (dk) | `keep_alive_min` | 10 | 0–120 | Pencere açıkken model son işten sonra bu kadar dakika bellekte kalır; aramalar modeli beklemez. Pencere kapanınca model de kapanır. 0: her aramada model yeniden yüklenir ve iş bitince bellek boşalır |
| Ek dosyalarının içinde de ara | `attachments` | true | – | Eklerin içeriğini oku ve ara. Kapalıysa yalnız ek adları aranır |
| (yalnız dosyada) | `idle_timeout_s` | 300 | 30–3600 | Eklenti yanıt vermezse arama işleminin kendini kapatma süresi |
| (yalnız dosyada) | `session_timeout_s` | 7200 | ≥ 60 | Bir aramanın en uzun süresi. Tam indekslemenin süre sınırı yoktur; 15 dakika hiç ilerleme olmazsa durur |

Eski `embed_budget_s` ayarı artık kullanılmaz. Var olan bir `settings.json` kurulumda değiştirilmez; yeni
`search_index_budget_s` yoksa varsayılan 20 saniye, `keep_alive_min` yoksa varsayılan 10 dakika geçerlidir.
İşlemci çekirdeği ya da bellek sınırı değişince hazır tutulan model kapanır; sonraki iş yeni değerlerle başlar.

### İndeksi silme

Yerel arama verilerini gizlilik için ya da bir sorunu gidermek için silebilirsiniz.

> **Dikkat:** Bu işlem tüm posta kutularının indeksini siler. Sonra arama yapmak için indekslemeyi baştan başlatmanız
> gerekir. Büyük bir posta kutusunda bu saatler sürebilir.

1. **Ayarlar** sekmesinde **İndeksi sil**'e basın. Arama ya da indeksleme sürerken düğme kapalıdır.
2. Onay penceresinde **Evet**'e basın. **Hayır** hiçbir şeyi değiştirmez.

- Silinenler: `data\index` (her posta kutusu için posta metni, kelime ve anlam indeksi, eklerden okunan metin) ve
  `data\tmp-att` (eklerin geçici kopyaları).
- Silinmeyenler: Outlook'taki postalar ve ekler, `settings.json`.
- Hazır tutulan yapay zekâ modeli de kapanır.
- Silme yarıda kalırsa (başka bir program bir dosyayı kullanıyorsa) durum çubuğunda "İndeks silinemedi" yazar.
  Biraz bekleyin ve yeniden deneyin.

## Kaldırma

> **Dikkat:** Outlook'u kapatın. `-KeepData` vermezseniz yerel indeks silinir. Outlook'taki postalarınız hiçbir durumda silinmez.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\HatirMail\uninstall.ps1"
```

- İndeksi saklamak için: `... uninstall.ps1 -KeepData`
- Kaldırma betiği yalnız kurulumun oluşturduğu kayıt anahtarlarını ve dosyaları siler.

## Gizlilik

- Bulut yapay zekâ, telemetri ve mail/sorgu kaydı (log) yoktur.
- Eklenti postaları yalnız okur. Gönderme, silme, taşıma ya da okundu işaretleme yapmaz. İndeksleme okundu durumunu
  değiştirmez. **Outlook'ta Aç** ile sizin açtığınız posta, Outlook'un kendi ayarına göre okundu olabilir (çift tıklayıp
  açmakla aynıdır).
- İndeks `%LOCALAPPDATA%\HatirMail\data\index` altındadır. Her posta kutusu (store) için ayrı bir dosya vardır.
  Eklerden okunan metin de yalnız bu indekste tutulur.
- Ek okunurken eklenti eki `%LOCALAPPDATA%\HatirMail\data\tmp-att` klasörüne kopyalar. Arama işlemi kopyayı okur ve
  hemen siler; arama bitince klasör de silinir. Asıl eke ve postaya dokunulmaz.
- İndeksi istediğiniz zaman **Ayarlar** sekmesindeki **İndeksi sil** düğmesiyle silebilirsiniz ("İndeksi silme"
  bölümüne bakın).
- Eklenti ile arama işlemi arasında ağ bağlantısı yoktur; iletişim yalnız işlemin kendi giriş/çıkış kanallarıyla yapılır.

## Sorun giderme

| Belirti | Çözüm |
|---|---|
| Düğme görünmüyor | Outlook > Dosya > Seçenekler > Eklentiler > COM Eklentileri: HatırMail işaretli olmalı. Devre dışı bırakılmışsa yeniden etkinleştirin. Kurulum çıktısında "COM activation check" uyarısı varsa "UAC kapalı ve hesap yönetici ise" bölümüne bakın. |
| "Arama bileşeni bulunamadı" | Kurulumu yeniden çalıştırın. |
| Arama çok yavaş | İndeks sekmesinden tam indekslemeyi bir kez yapın. Ayarlar sekmesinde "Aramada indeksleme" süresini azaltın ya da çekirdek sayısını artırın. |
| "İndeks eksik" yazıyor | Aramalar süre sınırlıdır. İndeks sekmesinde **İndekslemeyi başlat**'a basın. |
| "Posta açılamadı" | Posta taşınmış ya da silinmiş olabilir. Yeni bir arama indeksi günceller. |
| Resimdeki yazı bulunmuyor | Windows'ta o dilin OCR'ı yüklü olmayabilir (Ayarlar > Saat ve dil > Dil). Çok küçük ya da el yazısı metin okunamayabilir. |
| Güncellemeden sonra "Okunan postalar" %0 görünüyor | İndeks sürüm 2'ye geçerken her posta bir kez yeniden okunur (ek listesi için). Var olan anlam indeksi korunur. İndeks sekmesinden indekslemeyi başlatın. |
| Sonuçlar tutarsız ya da indeks hata veriyor | **Ayarlar** sekmesinde **İndeksi sil**'e basın, sonra İndeks sekmesinde **İndekslemeyi başlat**'a basın. |
| Pencere yazıları kesik ya da üst üste | 0.1.0 sürümünün 15 Ekim 13:42 ve sonraki paketi yüksek DPI ekranlara (ör. %200 ölçek) uyar. Daha eski bir paket kuruluysa Outlook'u kapatın ve `install.ps1 -SkipRegistration` ile dosyaları güncelleyin. |
