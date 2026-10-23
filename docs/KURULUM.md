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

## Kurulum (kurulum programıyla, önerilen yol)

> **Dikkat:** Kurulumdan önce Outlook'u kapatın. Outlook açıksa kurulum programı bekler ve kapatmanızı ister.

1. `HatirMail-Kurulum-0.1.0.exe` dosyasını bilgisayarınıza kopyalayın ve çift tıklayın.
2. Windows "Bilgisayarınız korundu" derse **Ek bilgi**'ye, sonra **Yine de çalıştır**'a basın. Kurulum programı imzasız
   olduğu için Windows bu uyarıyı gösterir.
3. **İleri**'ye basın. Kurulum birkaç dakika sürer.
4. Kurulumun sonunda Windows yönetici onayı isterse **Evet**'e basın. Bu onay güvenlik duvarı kuralı içindir (aşağıya
   bakın). **Hayır**'a basarsanız HatırMail yine çalışır, yalnız kural eklenmez.
5. **Bitir**'e basın. Outlook'u açın. **Giriş** sekmesinde **HatırMail ile Ara** düğmesi görünür.

Kurulum programı yalnız sizin kullanıcı hesabınıza kurar. Aşağıdaki PowerShell kurulumuyla aynı adımları yapar ve
HatırMail'i Windows **Ayarlar** > **Uygulamalar** listesine ekler. UAC kapalıysa ve hesap yöneticiyse gereken makine
kaydını ("UAC kapalı ve hesap yönetici ise" bölümü) kendisi yapar. Kurulum sırasında yaklaşık 2,5 GB boş disk alanı
gerekir; kurulumdan sonra 1,3 GB kalır.

**Güvenlik duvarı kuralı:** Kurulum programı Windows Güvenlik Duvarı'na bir kural ekler. Kural, HatırMail'in arama
işleminin (`%LOCALAPPDATA%\HatirMail\runtime\python\python.exe`) internete bağlanmasını engeller. HatırMail ağ
kullanmadığı için kuralla aynı biçimde çalışır. Outlook bu kuraldan etkilenmez. Kural eklemek yönetici onayı ister. UAC
kapalıysa ve hesap yöneticiyse onay sorulmaz. Kurulum sonunda kuralı denetler: ayrıntılarda "Güvenlik duvarı kuralı
etkin" yazar. Kural adı "HatırMail: arama işlemi internete çıkamaz"dır. Onu Windows Defender Güvenlik Duvarı > Gelişmiş
ayarlar > Giden Kuralları altında görebilirsiniz. Onay vermediyseniz kuralı sonra eklemek için kurulum programını
yeniden çalıştırın; indeksiniz korunur.

Kaldırmak için: Windows **Ayarlar** > **Uygulamalar** > **HatırMail** > **Kaldır**. Yerel indeks, ayarlar ve kayıtlı
aramalar da silinir. Güvenlik duvarı kuralı da silinir; bunun için Windows yeniden yönetici onayı isteyebilir. Outlook'taki
postalarınız değişmez.

**Yeni sürüme geçmek için** yeni kurulum programını eski sürümün üzerine çalıştırın. İndeksiniz, ayarlarınız ve kayıtlı
aramalarınız korunur. Önce eski sürümü kaldırmayın: kaldırma indeksi siler ve ilk indeksleme yeniden gerekir.

## Kurulum (zip paketi ve PowerShell ile)

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

Kurulum betiği güvenlik duvarı kuralını eklemez. Kuralı eklemek için kurulumdan sonra şu komutu çalıştırın. Windows
yönetici onayı ister:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\HatirMail\firewall.ps1"
```

Çıktıda "check OK" yazıyorsa kural etkindir. `uninstall.ps1` kuralı da siler.

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
2. **İndeks** sekmesini açın. Posta kutusunu ve klasörleri seçin. Silinmiş Öğeler, Önemsiz E-Posta, Taslaklar, Giden
   Kutusu ve Eşitleme Sorunları varsayılan olarak indeksten hariç tutulur ve bu listede görünmez. Hariç tutulan
   klasörleri **Ayarlar** sekmesinden değiştirebilirsiniz ("Klasörleri indeksten hariç tutma" bölümüne bakın).
3. **İndekslemeyi başlat**'a basın. Çubuklar okunan postaların, anlam indeksinin ve eklerin yüzdesini gösterir. Altında o
   anki aşama ve o aşamada kalan süre yazar.
4. İndeksleme sürerken Outlook'u kullanabilirsiniz. Pencereyi simge durumuna küçültebilirsiniz.

> **Not:** Pencereyi kapatırsanız ya da Outlook'u kapatırsanız indeksleme durur. Yapılan iş kaydedilir. Sonra
> **İndekslemeyi başlat**'a yeniden basın; iş kaldığı yerden sürer.

Büyük bir posta kutusunda tam indeksleme saatler sürebilir. 2.19 test makinesinde (Xeon E5-2620 v3, 4 iş parçacığı) birkaç
bin postalık gerçek bir PST'nin ilk indeksi 2 saatten uzun sürdü. Bu süre bir kez harcanır; sonraki çalışmalar yalnız yeni ve
değişen postaları işler.

### İndeksi güncel tutma: Güncelle

Tam indekslemeden sonra İndeks sekmesinde **Güncelle**'ye basın (kısayol: Alt+G). HatırMail yalnız son indekslemeden sonra
gelen, değişen, taşınan ve silinen postaları işler. Güncel bir indekste bu birkaç saniye sürer.

- Bitince durum çubuğunda sonuç yazar: "Güncelleme bitti: 3 yeni, 1 değişen posta işlendi; 2 posta indeksten çıktı".
  Değişiklik yoksa "İndeks güncel" yazar.
- Hiç tam indekslenmemiş bir klasördeki postaları Güncelle okumaz. Durum çubuğu bunların sayısını yazar. Bu klasörler için
  **İndekslemeyi başlat**'a basın.
- **Durdur**'a basarsanız yapılan iş kaydedilir. Sonra **Güncelle**'ye yeniden basın.
- Güncelleme sürerken **Ara**'ya basarsanız arama güncelleme bitince başlar.

### Arama

1. **Ara** sekmesinde hatırladıklarınızı yazın. Örnekler: `geçen ay Ahmet'ten gelen kargo faturası`,
   `mart başında gelen raf teklifi`, `mail from Sarah about the delayed shipment`. Ekin içinde geçen bir şeyi de
   yazabilirsiniz: `galvanizli raf dikmesi sayımı` (Excel), `soğuk zincir taşıma faturası` (PDF), `E-4012 hata kodu`
   (ekran görüntüsü).
2. **Ara**'ya basın. Önce kelime eşleşmeleri, sonra anlamsal sonuçlar görünür. Yeni postalar varsa önce indeksteki
   postalarla sonuçlar gelir. Yeni postalar indekslenince liste bir kez daha yenilenir. Seçtiğiniz posta seçili kalır.
   Pencere açılınca yapay zekâ modeli arka planda hazırlanır (durum çubuğunda "Yapay zekâ hazır" yazar). Pencere açık
   kaldıkça sonraki aramalar modeli beklemez.
3. Bir sonuç seçin. Alttaki önizlemede konu, kişi, tarih ve eşleşen bölüm görünür; aranan sözcükler sarı zeminle
   işaretlenir. Eşleşme bir ekin içindeyse önizlemenin üstünde "Ekte bulundu: dosya adı" yazar. Ek bir zip ise zip
   içindeki dosyanın adı da yazar (`Arsiv.zip › Klasör/rapor.docx`).
4. Postayı açmak için sonuca çift tıklayın ya da **Outlook'ta Aç**'a basın. Yalnız eki açmak için **Eki aç**'a basın
   (kısayol: Alt+E). Ek, bilgisayarınızda o dosya türünü açan programla açılır (Excel, Word, PDF okuyucu…). Zip içindeki
   bir dosyada yalnız o dosya açılır.

**Eki aç** hakkında:

- Açılan dosya ekin bir kopyasıdır. Kopya salt okunurdur. Office onu, Outlook'tan açılan bir ek gibi Korumalı Görünüm'de
  açar.
- Değişiklik yaparsanız dosyayı **Farklı Kaydet** ile başka bir klasöre kaydedin. Kopya HatırMail penceresi kapanınca
  silinir. Postadaki asıl ek hiçbir durumda değişmez.
- Eşleşme postanın kendi metnindeyse düğme görünmez. Posta taşınmış ya da değişmişse durum çubuğunda "Ek bulunamadı"
  yazar; postayı **Outlook'ta Aç** ile açın. Zip içindeki dosya artık yoksa zip'in kendisi açılır.

Arama, yeni postaları indekslemek için en çok **Aramada indeksleme** süresi kadar bekler (varsayılan 20 saniye). Kalan iş
sonraki aramaya ya da İndeks sekmesine kalır. Arama kutusunun altındaki satır, aranan klasörlerin indeks durumunu gösterir.
İndeks eksikse sonuçlar indeksteki postalarla sınırlıdır. Anlam indeksi hiç yoksa yalnız kelime eşleşmeleri gösterilir.

İndeksleme sürerken **Ara**'ya basarsanız HatırMail önce indekslemeyi durdurmak için onay ister. Aynı anda yalnız bir
arama işlemi çalışır.

### Filtreler: gönderen, alıcı, klasör, tarih ve ek

Arama kutusunun altında filtre alanları vardır. Alanlar üç eşit sütunda, alt alta hizalıdır:

- Birinci satır: **Gönderen** (Alt+N), **Alıcı** (Alt+L), **Klasör** (Alt+S).
- İkinci satır: **Tarih** (Alt+T), **Ek** (Alt+K) ve **Filtreleri temizle** düğmesi (Alt+F).
- Üçüncü satır yalnız **Özel aralık…** seçilince görünür: **Başlangıç** ve **Bitiş** tarihleri.

1. Hatırladığınız değeri seçin.
2. **Ara**'ya basın. Sonuçlarda yalnız bu filtrelere uyan postalar görünür. Durum çubuğu etkin filtreleri yazar.

**Gönderen, Alıcı ve Klasör listelerinde arama:** Liste açılınca en üstte bir arama alanı vardır.

1. Listeye tıklayın ya da liste seçiliyken yazmaya başlayın.
2. Adın ya da e-posta adresinin bir bölümünü yazın. Büyük-küçük harf ve Türkçe harfler fark etmez: "cakir" yazınca
   "Çakır" da bulunur. Birden çok sözcük yazarsanız hepsini içeren satırlar kalır.
3. Ok tuşlarıyla seçin ve Enter'a basın ya da satıra tıklayın. Esc listeyi kapatır, seçim değişmez.

**Birden çok klasör:** Klasör listesinde her satırın bir onay kutusu vardır.

1. Aramak istediğiniz klasörleri işaretleyin (tıklayın ya da Enter'a basın). Liste açık kalır.
2. Esc'e basın ya da listenin dışına tıklayın. Kutuda seçtiğiniz klasör ya da "3 klasör seçili" yazar.
3. **Ara**'ya basın. Yalnız işaretli klasörler aranır.

"(tümü)" satırını işaretlemek bütün işaretleri kaldırır.

- **Tarih:** Son 7 gün, Son 30 gün, Son 3 ay, Son 12 ay, Bu yıl, Geçen yıl ya da **Özel aralık…**. Özel aralıkta
  Başlangıç ve Bitiş tarihleri görünür; iki gün de aralığa dahildir. Tarih filtresi kesindir: aralık dışındaki posta
  görünmez. Sorguya yazdığınız tarih ("geçen ay") ise yalnız sıralamayı etkiler.
- **Ek:** Ekli, Eksiz, Excel, PDF, Word, PowerPoint, Resim ya da Zip, RAR. Postadaki ek adlarına bakılır. Metnin içine
  yapıştırılmış resimler (image001.png gibi) sayılmaz.
- Gönderen ve alıcı listeleri indeksteki postalardan dolar (en sık 1.000 kişi). Yeni bir kişi ilk indekslemeden ya da
  Güncelle'den sonra görünür.
- Klasör listesi İndeks sekmesinde işaretli klasörlerdir. Alt klasörler ayrı satırdır ve seçilen klasöre dahil değildir;
  bir alt klasörü de aramak için onu da işaretleyin.
- "(tümü)" o filtreyi kaldırır. **Filtreleri temizle** hepsini kaldırır. Filtre değişince arama kendiliğinden yeniden
  başlamaz; **Ara**'ya basın.

### Sonuçları sıralama ve sağ tık menüsü

- Sonuç listesinde bir sütun başlığına tıklayın: liste o sütuna göre sıralanır. Aynı başlığa yeniden tıklarsanız sıra
  tersine döner. Başlıktaki ▼ ya da ▲ işareti geçerli sıralamayı gösterir.
- İlk sıralama **Güven ▼**'dir: en uygun sonuç en üsttedir. **Tarih**'e tıklayınca en yeni posta üste gelir.
- Seçtiğiniz sıralama sonraki aramalarda da geçerlidir.
- Bir sonuca sağ tıklayın (ya da klavyede Menü tuşuna basın): **Outlook'ta aç**, **Eki aç** (yalnız eşleşme bir ekteyse),
  **Konuyu kopyala** ve **Yalnız bu gönderenle yeniden ara**. Sonuncusu Gönderen filtresini seçer ve son aramayı
  yeniden başlatır: Ara ile başlattığınız aramayı aynı sorguyla, **Benzerlerini bul**'u aynı e-postayla. Arama sürerken
  bu komut kapalıdır.

### Benzerlerini bul

Outlook'ta bir e-postaya sağ tıklayın ve **Benzerlerini bul**'a basın. HatırMail penceresi açılır ve aynı konudaki e-postaları
listeler. Sorgu yazmanız gerekmez. Arama, e-postanın indeksteki anlam vektörlerini kullanır; yapay zekâ modeli bu arama
için yüklenmez. Pencere bu komutla yeni açıldıysa arama, pencerenin arka planda başlattığı model hazırlığının bitmesini
bekleyebilir (birkaç saniye).

- Sonuçlar Ara sekmesindeki listededir. "Neden" sütununda benzerlik değeri yazar ("benzerlik 0.91"). Gri satırlar yalnız
  zayıf benzerliktir.
- Ara sekmesindeki filtreler burada da geçerlidir.
- E-posta henüz indekste yoksa ya da anlam indeksi yoksa durum çubuğu bunu yazar. İndeks sekmesinde **Güncelle**'ye basın
  ve yeniden deneyin.

### Son aramalar ve kayıtlı aramalar

HatırMail başlattığınız son 20 aramayı bu bilgisayarda saklar. Sık yaptığınız bir aramayı bir adla kaydedebilirsiniz.

- **Öneri:** Arama kutusuna yazmaya başlayın. Aynı harflerle başlayan son ve kayıtlı aramalar bir listede görünür. Ok
  tuşlarıyla birini seçin ve Enter'a basın.
- **Tek tıklamayla arama:** Arama kutusunun yanındaki **Aramalar ▼** düğmesine basın (kısayol: Alt+M). Kayıtlı ya da son
  aramalardan birine tıklayın. Arama hemen başlar.
- **Kaydetme:** Aramayı kutuya yazın, **Aramalar ▼** > **Bu aramayı kaydet…** öğesine basın, bir ad yazın ve **Kaydet**'e
  basın. Aynı adla yeniden kaydederseniz eski kaydın sorgusu değişir.
- **Silme:** **Aramalar ▼** > **Kayıtlı aramayı sil** ya da **Son aramaları temizle**. Son aramaları temizlemek kayıtlı
  aramaları silmez.

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
| Klasörleri hariç tut… | `excluded_folders` | (yok) | – | Posta kutusu kimliği → hariç tutulan klasör kimlikleri. Bir posta kutusu için kayıt yoksa varsayılan klasörler hariç tutulur (aşağıya bakın). Boş liste: hiçbir klasör hariç tutulmaz |
| (yalnız dosyada) | `idle_timeout_s` | 300 | 30–3600 | Eklenti yanıt vermezse arama işleminin kendini kapatma süresi |
| (yalnız dosyada) | `session_timeout_s` | 7200 | ≥ 60 | Bir aramanın en uzun süresi. Tam indekslemenin süre sınırı yoktur; 15 dakika hiç ilerleme olmazsa durur |

Eski `embed_budget_s` ayarı artık kullanılmaz. Var olan bir `settings.json` kurulumda değiştirilmez; yeni
`search_index_budget_s` yoksa varsayılan 20 saniye, `keep_alive_min` yoksa varsayılan 10 dakika geçerlidir.
İşlemci çekirdeği ya da bellek sınırı değişince hazır tutulan model kapanır; sonraki iş yeni değerlerle başlar.

### Klasörleri indeksten hariç tutma

Gizli tutmak istediğiniz postaları aramanın dışında bırakabilirsiniz. Hariç tutulan bir klasör indekslenmez ve
aranmaz. İndeks de küçülür.

> **Dikkat:** Kaydettiğinizde hariç tutulan klasörlerin postaları, ek metinleri ve anlam indeksi bu bilgisayardaki
> indeksten silinir. Klasörü sonra yeniden eklerseniz bu postalar baştan indekslenir. Outlook'taki postalarınız
> hiçbir durumda değişmez.

1. **İndeks** sekmesinde posta kutusunu seçin.
2. **Ayarlar** sekmesinde **Klasörleri hariç tut…**'a basın (kısayol: Alt+H). Arama ya da indeksleme sürerken düğme
   kapalıdır.
3. Açılan pencerede indekslenmeyecek klasörleri işaretleyin. Bir klasörü işaretlerseniz alt klasörleri de işaretlenir.
4. **Kaydet**'e basın. Durum çubuğunda indeksten kaç posta silindiği yazar.

- Varsayılan olarak Silinmiş Öğeler, Önemsiz E-Posta, Taslaklar, Giden Kutusu ve Eşitleme Sorunları hariç tutulur.
  Bir posta kutusu için seçim yapıp kaydettiğinizde varsayılanlar yerine sizin seçiminiz geçerli olur.
- Hariç tutulan bir klasörün sonradan oluşturulan alt klasörleri de hariç tutulur. Outlook'ta sildiğiniz bir klasör
  Silinmiş Öğeler'in altına taşınır; Silinmiş Öğeler hariç tutuluyorsa o klasör de indeksten çıkar.
- Bir postayı hariç tutulan klasöre taşırsanız posta sonraki aramada ya da indekslemede indeksten silinir.
- Seçim `settings.json` dosyasında saklanır. **İndeksi sil** seçimi silmez.
- Daha önce İndeks sekmesinde Silinmiş Öğeler gibi bir varsayılan klasörü elle işaretleyip indekslediyseniz bu sürümle
  o klasörün postaları pencere ilk açıldığında indeksten silinir. İndekslemek için klasörün işaretini **Klasörleri
  hariç tut…** penceresinde kaldırın.
- Silme, indeks dosyasını da sıkıştırır (SQLite VACUUM). Büyük bir indekste bu birkaç saniye sürebilir ve geçici
  olarak indeks dosyası kadar boş disk alanı gerekir.

### İndeksi silme

Yerel arama verilerini gizlilik için ya da bir sorunu gidermek için silebilirsiniz.

> **Dikkat:** Bu işlem tüm posta kutularının indeksini siler. Sonra arama yapmak için indekslemeyi baştan başlatmanız
> gerekir. Büyük bir posta kutusunda bu saatler sürebilir.

1. **Ayarlar** sekmesinde **İndeksi sil**'e basın. Arama ya da indeksleme sürerken düğme kapalıdır.
2. Onay penceresinde **Evet**'e basın. **Hayır** hiçbir şeyi değiştirmez.

- Silinenler: `data\index` (her posta kutusu için posta metni, kelime ve anlam indeksi, eklerden okunan metin) ve
  `data\tmp-att` (eklerin geçici kopyaları).
- Silinmeyenler: Outlook'taki postalar ve ekler, `settings.json`, son ve kayıtlı aramalar (`data\searches.json`).
- Hazır tutulan yapay zekâ modeli de kapanır.
- Silme yarıda kalırsa (başka bir program bir dosyayı kullanıyorsa) durum çubuğunda "İndeks silinemedi" yazar.
  Biraz bekleyin ve yeniden deneyin.

## Kaldırma

> **Dikkat:** Outlook'u kapatın. `-KeepData` vermezseniz yerel indeks, ayarlar ve son/kayıtlı aramalar silinir. Outlook'taki postalarınız hiçbir durumda silinmez.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\HatirMail\uninstall.ps1"
```

- İndeksi, ayarları ve aramaları saklamak için: `... uninstall.ps1 -KeepData`
- Kaldırma betiği yalnız kurulumun oluşturduğu kayıt anahtarlarını, dosyaları ve bu kurulumun güvenlik duvarı kuralını
  siler. Kural varsa Windows yönetici onayı ister.

## Gizlilik

- Bulut yapay zekâ, telemetri ve mail/sorgu kaydı (log) yoktur.
- Her aramanın süreleri `%LOCALAPPDATA%\HatirMail\data\timing.log` dosyasına yazılır. Dosyada yalnız sayılar vardır
  (milisaniye, posta sayısı). Arama metni, posta metni, klasör adı ve posta kimliği yazılmaz. Dosya hiçbir yere
  gönderilmez. En çok iki dosya (`timing.log`, `timing.log.1`), her biri en çok 1 MB tutulur. Dosyayı silebilirsiniz.
- Eklenti postaları yalnız okur. Gönderme, silme, taşıma ya da okundu işaretleme yapmaz. İndeksleme okundu durumunu
  değiştirmez. **Outlook'ta Aç** ile sizin açtığınız posta, Outlook'un kendi ayarına göre okundu olabilir (çift tıklayıp
  açmakla aynıdır).
- İndeks `%LOCALAPPDATA%\HatirMail\data\index` altındadır. Her posta kutusu (store) için ayrı bir dosya vardır.
  Eklerden okunan metin de yalnız bu indekste tutulur.
- Ek okunurken eklenti eki `%LOCALAPPDATA%\HatirMail\data\tmp-att` klasörüne kopyalar. Arama işlemi kopyayı okur ve
  hemen siler; arama bitince klasör de silinir. Asıl eke ve postaya dokunulmaz.
- **Eki aç** eki `%LOCALAPPDATA%\HatirMail\data\open-att` klasörüne kopyalar ve kopyayı açar. Bu kopyalar pencere
  kapanınca silinir. O sırada bir program kopyayı açık tutuyorsa kopya, pencere bir sonraki açılışta silinir.
- İndeksi istediğiniz zaman **Ayarlar** sekmesindeki **İndeksi sil** düğmesiyle silebilirsiniz ("İndeksi silme"
  bölümüne bakın).
- Gizli tutmak istediğiniz klasörleri **Ayarlar** > **Klasörleri hariç tut…** ile indeksin dışında bırakabilirsiniz.
  Bu klasörlerin metni indeks dosyasından da silinir ("Klasörleri indeksten hariç tutma" bölümüne bakın).
- Son ve kayıtlı aramalar yalnız `%LOCALAPPDATA%\HatirMail\data\searches.json` dosyasında düz metin olarak durur. Bu
  dosya hiçbir yere gönderilmez ve günlüğe yazılmaz. **Aramalar ▼** menüsünden silebilirsiniz ("Son aramalar ve kayıtlı
  aramalar" bölümüne bakın).
- Eklenti ile arama işlemi arasında ağ bağlantısı yoktur; iletişim yalnız işlemin kendi giriş/çıkış kanallarıyla yapılır.
- Kurulum programı, arama işleminin internete bağlanmasını Windows Güvenlik Duvarı kuralıyla da engeller ("Kurulum
  (kurulum programıyla, önerilen yol)" bölümüne bakın).
- İndeks şifreli değildir. Bilgisayar kaybolursa ya da çalınırsa postalarınızı ve indeksi korumak için disk şifrelemesini
  (BitLocker ya da Cihaz şifrelemesi) açık tutun.

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
