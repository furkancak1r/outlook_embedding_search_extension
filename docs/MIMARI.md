# HatırMail – Mimari, kararlar ve sınırlar

## Özet

HatırMail, Klasik Outlook için bir COM eklentisidir. Arama ve yapay zekâ kullanıcının kendi bilgisayarında, tek bir alt
süreçte (worker) çalışır. Worker yalnız kullanıcı eylemiyle başlar: arama penceresini açmak, **Ara**,
**İndekslemeyi başlat**, posta kutusu değiştirmek ya da **Durumu yenile**. Pencere açıkken model bellekte hazır kalır
(ayar: **Yapay zekâyı hazır tut**, varsayılan 10 dk). Pencere kapanınca worker ve model de kapanır. Ortak sunucu yoktur.

```
Outlook.exe (Klasik)
 └─ HatirMail.dll  (COM eklentisi, .NET Framework 4.8, geç bağlama; sekmeler: Ara, İndeks, Ayarlar)
      │  pencere açıkken: tek alt süreç (Job Object: kapanınca öldür, bellek sınırı, alt süreç yasağı)
      │  stdin/stdout üzerinden JSON satırları (ağ yok, port yok)
      ▼
 python.exe hm_worker.py  (gömülü CPython 3.13.15, kurulum klasöründe)
      ├─ SQLite: FTS5 (kelime) + float16 vektörler (anlam), her posta kutusu için ayrı dosya
      └─ EmbeddingGemma 2, yalnız 270M metin bileşeni, ONNX Runtime, CPU, float32
         (pencere kapanınca ya da ayarlanan süre işlem olmayınca süreç kapanır; ağırlıklar bellekte kalmaz)
```

## Doğrulanan ortam (2.19 test makinesi, salt okunur kontrol)

| Konu | Sonuç |
|---|---|
| Windows | 10 Pro 22H2 (19045), x64, 2× Xeon E5-2620 v3 (AVX2), 54 GB RAM |
| Outlook | Office Standard 2016 MSI 16.0.4266.1001, Outlook **64 bit**, Klasik |
| Yeni Outlook | Paket kurulu (1.2026.818.0), veri klasörü 2025-02-18'den beri değişmemiş |
| Hesap | Tek profil: PST deposu (`MSUPST MS`) + kişi listesi. Exchange/Microsoft 365/IMAP hesabı yok |
| .NET | Framework 4.8 (Release 533325), Outlook PIA 15 GAC'ta, RegAsm mevcut |
| Ağ | huggingface.co ve pypi.org erişilebilir (yalnız paket hazırlığı için; çalışma anında ağ yok) |

## Kararlar

### 1. Klasik Outlook: COM eklentisi (seçilen)

- Outlook Nesne Modeli (OOM) PST, OST ve Exchange önbelleğini çevrimdışı okur. Tüm arşiv erişilebilir.
- İşlem içi (in-process) eklenti, Outlook'un "Nesne Modeli Koruması" uyarısına takılmaz.
- Geç bağlama (late binding): PIA gerekmez. AnyCPU derleme 32 ve 64 bit Outlook'ta çalışır.
- Kayıt yalnız HKCU altındadır; yönetici yetkisi gerekmez.
- Klasör listesi `Folder.GetTable()` ile okunur (öğe açmadan EntryID, değişiklik zamanı, sınıf, tarih).
  Yalnız yeni ya da değişen öğelerin içeriği okunur.

### 2. Yeni Outlook: değerlendirildi, v1'de yok

- Yeni Outlook yalnız Office.js web eklentilerini yükler. COM eklentisi çalışmaz.
- Office.js yalnız açık öğeye ve sınırlı REST/EWS çağrılarına erişir. Tüm arşivi doğrudan okuyamaz.
- Tüm kutuyu okumak için Microsoft Graph + MSAL + delegated `Mail.Read` gerekir. Bu yol:
  - Exchange Online / Microsoft 365 hesabı ister (test makinesindeki hesap yalnız PST; bu yol burada çalışamaz),
  - Entra uygulama kaydı ve kiracı onayı ister,
  - postaları buluttan indirir; tamamen çevrimdışı çalışma sözü verilemez,
  - yerel AI için ayrıca bir yerel köprü (loopback + oturum belirteci + origin denetimi) gerektirir.
- Bu nedenle v1 kapsamı Klasik Outlook'tur. Yeni Outlook için olası yol (uygulanmadı): Office.js görev bölmesi → MSAL (delegated Mail.Read) → Graph `/me/messages` delta sorgusu →
  yerel köprü (yalnız 127.0.0.1, rastgele oturum belirteci, `Origin` doğrulaması) → aynı worker.

### 3. Model: google/embeddinggemma-2, yalnız metin bileşeni

| Konu | Değer | Kanıt |
|---|---|---|
| Revizyon | `914f7f89142e33e77833254d9c9b90c3cef7303b` | HF API, `packaging/model.lock.json` |
| Lisans | Apache-2.0 | Model kartı, HF etiketleri |
| Yükleme | `vision_config=None`, `audio_config=None`, float32, CPU | Model kartı "Selective Encoder Loading", `eval/export_onnx.py` |
| Çalışma | ONNX Runtime 1.31.0, CPU; torch ve transformers çalışma anında yok | `hm_embed.py`, `eval/parity_onnx.py` |
| Parametre | 271.002.624 (metin bileşeni) | `eval/probe_model.py` |
| Sorgu biçimi | `task: search result \| query: {sorgu}` | Model kartı, `config_sentence_transformers.json` |
| Belge biçimi | `title: {konu} \| text: {içerik}` (konu yoksa `none`) | Model kartı |
| Boyut | 512 (MRL kırpma + yeniden normalleştirme) | Model kartı: 512d MTEB 61.17, 768d 61.36 |
| Hassasiyet | float32 (CPU'da bf16 yok; float16 yasak) | Model kartı "Numerical Precision" |

- Metin bileşeni yalnız `language_model.*` tensörleridir. Bayt düzeyinde aynıdır; resmi dosyalarla çıktı farkı **0,0**
  (`eval/check_text_only.py`). Resmi sentence-transformers 6.1.0 yoluyla eşdeğerlik: en büyük fark **3e-8**, kosinüs
  1,0 (`eval/parity_st.py`).
- **ONNX (20 Ekim):** metin bileşeni geliştirme makinesinde bir kez ONNX'e aktarılır (`eval/export_onnx.py`; torch 2.14.1,
  onnx 1.23.2, onnxscript 0.7.2). Grafik: `input_ids`, `attention_mask` → model → maske ile ortalama → ilk 512 boyut →
  L2 normalleştirme. Ağırlıklar float32'dir (bf16 değerlerin tam karşılığı), 1,08 GB.
- **Eşdeğerlik (torch ↔ ONNX Runtime):** 86 sorgu, 316 belge parçası ve uç durumlar (boş metin, 512 token, dolgu):
  en büyük fark **1,1e-6**, en küçük kosinüs **0,9999999** (Mac, `eval/parity_onnx.py`). Aynı indekste iki kodlayıcıyla
  86 sorgunun ilk 10 sonucu birebir aynıdır. Model revizyonu değişmedi: var olan vektörler geçerli kalır, yeniden indeks
  gerekmez.
- **Neden ONNX:** 2.19'da yükleme süresinin çoğu `import torch` (3,6–6,2 sn) ve `import transformers` (8,4–30,7 sn)
  idi. ONNX Runtime ile worker yalnız `onnxruntime`, `tokenizers` ve `numpy` yükler.
- Tokenizer resmi `tokenizer.json` dosyasıdır; `tokenizers` kütüphanesi doğrudan kullanılır.
- Sohbet modeli, yeniden sıralayıcı (reranker), görüntü ve ses kodlayıcı yüklenmez.

### 4. İndeks ve arama

- **Kimlik ve sürümler:** `items` tablosu EntryID, değişiklik zamanı, içerik hash'i, model revizyonu ve indeks sürümü tutar.
- **Değişiklik algılama:** klasör listesi (EntryID + değişiklik zamanı) her aramada karşılaştırılır.
  Yeni/değişen öğeler okunur; listede olmayanlar silinir; başka klasöre geçenlerin klasörü güncellenir.
- **Vektör yeniden kullanımı:** vektörler parça metninin hash'iyle saklanır. Taşınan ya da yalnız bayrağı değişen posta
  yeniden kodlanmaz.
- **Metin hazırlığı:** Türkçe uyumlu katlama (İ/I/ı/i → i, ş → s …), alıntı zinciri, imza, mobil alt satır ve yasal uyarı
  ayıklama; uzun metin 1.200 karakterlik parçalara (150 karakter örtüşme) bölünür. İletilen (FW/İLT) postada iletilen
  metin içerik sayılır.
- **Kelime araması:** FTS5 (contentless, `unicode61 remove_diacritics 2`, önek dizinleri); hafif Türkçe/İngilizce kök
  ayıklama, sözlükten yazım hatası genişletmesi (Damerau-Levenshtein 1–2), ek türü eşanlamlıları (excel → xlsx).
- **Tarih:** "geçen ay", "mart başında", "2 hafta önce", "yazın", "geçen yıl kasımda", "last month", "before Christmas" …
  yumuşak tarih penceresine çevrilir. Pencere kesin süzgeç değildir; yalnız sıralamayı etkiler.
  "Kasım'ın", "aralığı", "nişan", "Martin" gibi belirsiz sözcükler tarih sayılmaz.
- **Sıralama:** `kosinüs(tam sorgu) + 0,08 × kelime kapsamı + 0,04 × tarih puanı + 0,5 / (10 + BM25 sırası)`.
  Ağırlıklar sentetik sorguların tek numaralı yarısında seçildi, çift numaralı yarısında sınandı.
- **Güven etiketi:** Yüksek / Orta / Zayıf. Zayıf sonuçlar gri gösterilir; hiçbir sonuç "bulundu" diye sunulmaz.
  Anlamsal aşama bitmeden yalnız tüm sorgu sözcüklerini içeren postalar ("Kelime") gösterilir.

### 4a. Ek içeriği (indeks sürümü 2)

- **Akış:** kelime sonuçlarından sonra worker okunacak ekleri dörderli ister (`need_att`). Eklenti her eki
  `Attachment.SaveAsFile` ile `data\tmp-att` klasörüne sayaç adıyla kaydeder (`att_files`). Worker dosyayı okur ve hemen
  siler. Oturum sonunda eklenti bu klasörü siler. Okuma Outlook sürecinde değil, worker'da yapılır.
- **Okuyucular** (`worker/hm_attach.py`): xlsx/xlsm (akışla, bellek sabit), xls (xlrd), docx, pptx, OpenDocument,
  PDF metin katmanı (pypdfium2), Outlook iletisi (.msg, olefile), zip (bir düzey), txt/csv/html.
  Resim ve metni olmayan PDF sayfası (tarama): `Windows.Media.Ocr` (Windows'un kendi OCR'ı, kullanıcının OCR dilleri;
  2.19'da `tr`). OCR alt süreç başlatmaz.
- **Sınırlar:** ek başına 20 MB dosya ve 100.000 karakter metin, posta başına 300.000 karakter; PDF'te 50 sayfa,
  OCR için 3 sayfa; zip'te üye başına 20 MB ve toplam 100 MB. Küçük satır içi resimler (40 KB altı `image001.png` gibi
  logolar) okunmaz.
- **İndeks:** `atts` (her ekin durumu), `att_texts` (dosya SHA-256'sı → metin; aynı dosya bir kez okunur, taşınan posta
  yeniden okunmaz), `fts_att` (ek metninin kelime dizini), `chunks.src` (0 = gövde, n = n. ek). Her ekin ilk 3 parçası
  anlam dizinine girer (posta başına en çok 12); metnin tamamı kelimeyle aranır. Silinen postanın ek metni oturum
  sonunda silinir.
- **Çökme koruması:** okunan ek önce `extracting` yazılır. Worker o dosyada ölürse (bozuk dosya, bellek sınırı), sonraki
  arama eki `failed` yapar ve aynı dosyayı bir daha denemez.
- **Sıralama:** ekteki BM25 sırası postanın kendi sırasının yarı ağırlığıyla eklenir (`0,5 × 0,5 / (10 + sıra)`);
  kelime kapsamı ekteki sözcükleri de sayar. Bu ağırlık kalibre edilmedi. Özet, sorgu sözcüklerinin
  en çok geçtiği metinden (posta ya da ek) alınır ve `[Ek: ad]` ile gösterilir.
- **Geçiş:** sürüm 1 indeksi açılınca her posta bir kez yeniden okunur (ek listesi için); var olan vektörler korunur.

### 5. Süreç yaşam döngüsü ve kaynak sınırları

- Outlook açılırken model yüklenmez, alt süreç başlamaz, zamanlayıcı ya da yoklama (polling) yoktur. Arka planda
  kendiliğinden indeksleme yoktur.
- Worker yalnız kullanıcı eylemiyle başlar. Üç çalışma modu vardır (`docs/PROTOKOL.md`):
  - `search` (**Ara**): indeksleme işi `search_index_budget_s` ile sınırlıdır (varsayılan 20 sn). Eklenti yeni postaları
    sürenin en çok yarısı kadar okur; worker kalan sürede ekleri okur ve anlam indeksini yapar. Modelin yüklenme süresi
    sayılmaz. Kapsamda hiç anlam indeksi yoksa model yüklenmez, yalnız kelime sonuçları gelir.
  - `index` (**İndekslemeyi başlat**): süre sınırı yok, sorgu yok. Kullanıcı **Durdur**'a basabilir.
  - `status`: yalnız indeks durumu (klasör başına okunan/anlam/ek sayıları). Tarama, posta okuma ve model yok;
    bilinmeyen posta kutusu için dosya da açılmaz. Pencere açılınca, posta kutusu değişince, **Durumu yenile**'ye
    basılınca ve ek okuma ayarı değişince bir kez çalışır.
- **Hazır model (20 Ekim, kullanıcı isteği):** pencere açıkken worker bir işten sonra kapanmaz, sonraki işi bekler; model
  bellekte kalır. Pencere açılınca yapılan durum okumasından sonra, kapsamda anlam indeksi varsa worker modeli arka planda
  yükler (`preload`). İlk arama da modeli beklemez. Bellek: pencere açıkken ~1–1,5 GB (2.19 ölçümü).
  Ayar **Yapay zekâyı hazır tut (dk)**: varsayılan 10, 0–120. 0 eski davranıştır: her iş kendi worker'ında, iş bitince
  süreç kapanır, ön yükleme yok. Süre son işten (ya da ön yüklemenin bitişinden) sayılır.
- Aynı anda tek worker çalışır. Arama ve indeksleme birbirini dışlar: indeksleme sürerken **Ara**, onaydan sonra
  indekslemeyi durdurur ve aramayı başlatır. Durum okuması bir saniye kadar sürer; o sırada basılan **Ara** ve
  **İndekslemeyi başlat** onun bitişini bekler.
- Worker şu durumlarda kapanır: pencere kapandı, Outlook kapandı ya da eklentiyi kaldırdı, hazır tutma süresi doldu,
  **Durdur** (iptal), hata, eklenti sessiz kaldı (`idle_timeout_s`), aramanın süresi doldu (`session_timeout_s`), worker
  15 dakika hiç ileti göndermedi, stdin kapandı (Outlook çöktü), iş parçacığı ya da bellek ayarı değişti (bu değerler
  süreç başlarken verilir). Worker'ın kendi güvenlik sınırı: hazır tutma süresi + 10 dk içinde yeni iş gelmezse kendisi
  çıkar. Pencere indeksleme sürerken kapanırsa kullanıcıdan onay istenir.
- Windows Job Object: `KILL_ON_JOB_CLOSE` (Outlook ölürse worker da ölür), süreç başına bellek sınırı (4 GB),
  `ActiveProcessLimit=1` (worker başka süreç başlatamaz). Öncelik: BelowNormal. İş parçacığı: 4 (Ayarlar sekmesi).
- Durdurulan indeksleme kaybolmaz: vektörler her model partisinde (8 parça) kaydedilir. Değişen ama henüz yeniden
  okunmayan posta eski metniyle aranabilir kalır ve "okunacak" olarak işaretlenir; sonraki tarama onu yeniden ister.
- Eklenti postaları Outlook'un arayüz iş parçacığında okur ve her 100 ms'de bir pencere iletilerini işler. Böylece uzun
  bir indeksleme sırasında Outlook ve **Durdur** düğmesi yanıt verir.

### 6. Gizlilik

- Çalışma anında ağ kullanılmaz (`HF_HUB_OFFLINE=1`, telemetri kapalı).
- **ONNX Runtime telemetrisi:** resmî ONNX Runtime 1.29+ sürümleri varsayılan olarak kullanım olayları toplar ve
  kullanıcı profiline bir cihaz kimliği yazar ([Privacy.md](https://github.com/microsoft/onnxruntime/blob/main/docs/Privacy.md)).
  Mac'te bayraksız denemede `Library/Application Support/Microsoft/DeveloperTools/.onnxruntime` altında `deviceid` ve
  `onnxruntime.db` oluştu. HatırMail worker ortamına `ORT_DISABLE_TELEMETRY=1` koyar; worker bunu ONNX Runtime
  başlamadan yeniden yazar ve `disable_telemetry_events()` çağırır. Birim testi ve 2.19 testi profil klasöründe hiçbir
  dosya oluşmadığını denetler.
- Mail ya da sorgu metni hiçbir log dosyasına yazılmaz. Worker stderr'i yalnız boşaltılır, saklanmaz.
  Protokol testi, sorgu metninin stderr'de görünmediğini denetler.
- Eklenti postaları yalnız okur: `Send`, `Delete`, `Move`, `UnRead` gibi üyeleri hiç çağırmaz.
  `Body` okumak öğeyi okundu yapmaz; "Outlook'ta Aç" yalnız `Display()` çağırır. `Attachment.SaveAsFile` eki
  değiştirmez; kopyası yalnız worker okuyana kadar yerel veri klasöründe durur.
- Ek metni de posta metni gibi yalnız yerel indekste tutulur; log'a ya da başka bir yere yazılmaz.
- Her posta kutusu (store) ayrı SQLite dosyasıdır (`data\index\<store hash>.sqlite`).
- **İndeksi sil** (Ayarlar sekmesi, onaydan sonra): eklenti `data\index` ve `data\tmp-att` klasörlerini siler. Vektörler
  ve ek metni önbelleği (`att_texts`, `fts_att`) aynı SQLite dosyasındadır. Silmeden önce hazır tutulan worker kapanır.
  Oturum sürerken düğme kapalıdır; bu yüzden hiçbir worker dosyaları açık tutmaz. `settings.json` kalır. Worker ve
  protokol değişmez. Silmeden sonraki durum okuması yeni dosya açmaz (bilinmeyen posta kutusu kuralı).
- Eklenti ile worker arasında soket yoktur. Oturum belirteci (`HM_SESSION`) ortam değişkeniyle verilir;
  worker eşleşmeyen `begin` mesajını reddeder ve hiçbir dosya yazmaz.

## Bilinen sınırlar

| Sınır | Etki | Yükseltme yolu |
|---|---|---|
| Bazı ek türleri okunmaz | rar/7z, eski .doc/.ppt, ses, video ve CAD eklerinde yalnız ek adı aranır. Resim ve taranmış PDF yalnız Windows OCR ile okunur. 20 MB üstü ek okunmaz. | `worker/hm_attach.py` içine yeni bir okuyucu |
| Yalnız ekte geçen sözcükte yazım hatası düzeltilmez | Yazım düzeltme sözlüğü yalnız postaların kendi metninden kurulur (tablolardaki çok sayıda kod taramayı yavaşlatır) | Ek sözcüklerini de içeren, daha hızlı bir sözlük |
| Ek sıralama ağırlığı kalibre edilmedi | 13 sorguluk küçük ek test kümesiyle sınandı | Gerçek kutudan etiketli sorgularla kalibrasyon |
| İlk indeks CPU'ya bağlı | Birkaç bin postalık gerçek bir PST'nin ilk indeksi 2.19'da (Haswell, 4 iş parçacığı) 2 saatten uzun sürdü (kullanıcı gözlemi; dosya zamanlarına göre ~2 sa 10 dk; torch ile). ONNX Runtime ile sentetik ilk indeks 62 → 37 sn
oldu; gerçek kutuda süre henüz ölçülmedi. Arama artık bunu beklemez (süre sınırı); tam indeks İndeks sekmesinden bir kez yapılır | Daha hızlı CPU, daha çok çekirdek ya da daha küçük model |
| İndeksleme yalnız Outlook açıkken çalışır | Postaları Outlook nesne modeli okur; Outlook kapanınca indeksleme durur (iş kaydedilir) | Arka plan hizmeti (kapsam dışı: yalnız kullanıcı eylemiyle çalışma kuralı) |
| Kalan süre aşama başınadır | İndeks sekmesi o anki aşamanın (okuma, ek, anlam) kalan süresini gösterir, toplamı göstermez | Aşama hızlarını kaydedip toplam tahmin |
| Hazır model bellekte yer tutar | Pencere açıkken worker ~1–1,5 GB bellek kullanır | Ayarlar: hazır tutma süresini kısaltın ya da 0 yapın |
| Vektör araması kaba kuvvet (tüm vektörler bellekte) | 100.000 parça için ~200 MB, birkaç yüz ms | ANN dizini (ör. HNSW) |
| Yazım hatası genişletmesi sözlükte doğrusal tarama | ~10^5 terimde birkaç yüz ms | Sözlük üzerinde trigram FTS |
| Alıntı/imza ayıklama sezgiseldir | Alışılmadık biçimler ana metinde kalabilir | Daha çok kalıp |
| Yeni Outlook desteklenmez | Yalnız Klasik Outlook | Bölüm 2'deki Graph + köprü yolu |
| Gerçek Outlook arayüz testi yapılmadı | SSH oturumu (Session 0) masaüstü arayüzü açamaz | Kullanıcı makinesinde el ile kabul testi (`docs/KURULUM.md`) |
