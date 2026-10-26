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
  Satırlar `Table.GetArray(500)` ile 500'lük paketlerle alınır: değerler `GetNextRow` ile aynıdır, Outlook'a
  satır başına bir çağrı yerine paket başına bir çağrı gider. Yalnız yeni ya da değişen öğelerin içeriği okunur.

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
- **Erken anlamsal liste (23 Ekim):** aramanın indeksleme işi varsa (vektörü olmayan yeni parçalar ya da okunmamış
  ekler), worker önce indeksteki vektörlerle anlamsal listeyi gönderir (`semantic` aşaması), sonra işi yapar ve son
  listeyi (`final`) eskisi gibi baştan hesaplar. Son liste değişmez; erken liste yalnız eklenir. İndeksleme süresi bu
  adımın süresi kadar uzar. Ölçüm (Mac, 28.632 posta, 50 yeni posta): ilk anlamsal liste 2,57 s yerine 0,24 s'de geldi;
  5 sorgunun 5'inde son liste (posta, puan, güven) eski worker ile aynı.
- **Tarama (23 Ekim):** worker artık her taramada her satıra `seen_gen` yazmaz. Görülen satırlar bellekte tutulur;
  yalnız taşınan postaların klasörü yazılır. Silinecekler: listelenen klasörlerde olup bu taramanın hiçbir manifestinde
  olmayan satırlar (silme kuralı aynı). Ölçüm: tarama 0,22 s'den 0,09 s'ye indi.

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
  en çok geçtiği metinden (posta ya da ek) alınır. Önizlemenin üstünde "Ekte bulundu: ad" yazar.
- **Eki aç (20 Ekim, kullanıcı isteği):** sonuç ekin Outlook sırasını (`match_att_index`) ve zip ekinde eşleşen üyenin
  yolunu (`match_att_member`) taşır. Üye, `zip_text` metninde özetin ilk eşleşmesinden önceki son `[üye yolu]`
  satırıdır. Eklenti eki `Attachment.SaveAsFile` ile `data\open-att\<8 hex>\` klasörüne kaydeder. Zip ise yalnız o üyeyi
  `System.IO.Compression` ile çıkarır (üye başına en çok 20 MB, worker sınırıyla aynı). Kopyayı salt okunur yapar,
  Mark of the Web ekler (`Zone.Identifier`, ZoneId=3: Office Korumalı Görünüm) ve Windows'un dosya türü ilişkisiyle
  açar. Eklenti açmadan önce ekin adını o sıradaki ekle karşılaştırır (eski indeks yanlış eki açmaz). Kopyalar pencere
  açılınca ve kapanınca silinir.
- **Geçiş:** sürüm 1 indeksi açılınca her posta bir kez yeniden okunur (ek listesi için); var olan vektörler korunur.

### 4b. Güncelle, filtreler ve benzer e-postalar (21 Ekim, kullanıcı isteği)

- **Güncelle (İndeks sekmesi, worker modu `update`):** yalnız son tam indekslemeden sonra gelen, değişen, taşınan ve
  silinen postaları işler.
  - Bir indeksleme okumayı durdurulmadan bitirince klasör `folders.synced_gen` alır. Her posta, eklendiği ya da
    değiştiği oturumun numarasını `items.added_gen` alanında tutar.
  - Güncelle değişen postaları her zaman okur. Yeni kimlikleri yalnız tam indekslenmiş klasörlerde okur; öteki
    klasörlerdeki yeni postalar sayılır ve atlanır (`unsynced_skipped`). İndekste olmayan bir posta, o klasör hiç tam
    okunmadıysa eski bir posta da olabilir.
  - Ekler ve anlam indeksi yalnız bu fark (delta) için yapılır. Eski birikmiş işe İndekslemeyi başlat bakar.
  - Klasör listesindeki kimlikler kullanılır, değişiklik zamanı kullanılmaz. Bir PST içinde taşıma EntryID'yi
    değiştirir. Taşımanın değişiklik zamanını koruyup korumadığı doğrulanmadı.
  - Eski indeks dosyalarında tamamen okunmuş klasör, son taramasında indekslenmiş sayılır. İndeks sürümü 2 olarak kalır.
- **Filtreler (Ara sekmesi):** Gönderen, Alıcı ve Klasör açılır listeleri. Gönderen ve alıcı listeleri her `done`
  olayıyla gelen `facets` alanından dolar (indeks dosyasının tamamı, en çok 1.000 değer). Klasör listesi İndeks
  sekmesindeki işaretli klasörlerdir; alt klasörler ayrı satırdır ve dahil değildir.
  - Worker Türkçe katlamayla eşleştirir ("ÇAKIR" = "Çakır").
  - Filtre aday kümesini daraltır. Güven etiketi filtreli ve filtresiz aramada aynıdır (z istatistiği tüm kapsamdan).
  - Filtre değişince arama yeniden başlamaz; durum çubuğu Ara'ya basmayı söyler.
  - **Tarih ve Ek filtreleri (22 Ekim):** ikinci filtre satırı. Tarih hazır aralıklar ya da "Özel aralık…" (iki tarih
    kutusu, iki gün dahil) verir; eklenti bunu yerel saatle `[date_from, date_to)` olarak gönderir. Bu kesin bir
    süzgeçtir; sorgudaki tarih sözcükleri ise yalnız sıralamayı etkiler. Ek filtresi `items.attachments` içindeki ek
    adlarına bakar (ekli, eksiz ya da dosya türü). Metne gömülü küçük resimler bu listede yoktur.
  - **Yerleşim ve liste araması (22 Ekim):** bütün filtre alanları tek bir tablodadır: üç eşit sütun, sabit genişlikte
    etiket sütunları. Özel aralığın tarihleri üçüncü satırda, üstteki listelerle aynı hizadadır. Gönderen, Alıcı ve
    Klasör listeleri `SearchList` ile açılır: üstte arama alanı, altta süzülen satırlar (sözcüklerin hepsi, harf ve
    Türkçe işaret farkı yok). Açılır pencere `ToolStripDropDown`'dır. Outlook'un ileti döngüsünde WinForms bu pencere
    açıkken tuşları kendi kancasıyla görür ve dışarı tıklanınca pencereyi kapatır.
  - **Birden çok klasör (23 Ekim):** Klasör listesi onay kutuludur (`SearchList(multi)`, `CheckedListBox`). Seçilen
    klasörler aramanın kapsamıdır (eklenti kapsamı daraltır; worker'a giden protokol değişmedi). İndeks sekmesinde
    işareti kalkan klasör filtreden de çıkar.
- **Sonuç listesi (22 Ekim):** sütun başlığına tıklamak listeyi o sütuna göre sıralar, ikinci tıklama sırayı çevirir.
  Güven sütunu worker'ın sırasıdır (`rank`). Seçilen sıralama yeni sonuçlarda da kalır. Sağ tık menüsü: Outlook'ta aç, Eki
  aç, Konuyu kopyala, Yalnız bu gönderenle yeniden ara (sonuçtaki `sender_key` ile Gönderen filtresi). Bu komut son
  başlayan işi yineler: metin aramasını aynı sorguyla, Benzerlerini bul'u aynı postayla. Arama ya da benzer arama
  sürerken ve henüz hiç iş başlamadıysa kapalıdır.
- **Benzerlerini bul (Outlook'ta bir postaya sağ tık, worker modu `similar`):** postanın indeksteki vektörleriyle arar.
  Model yüklenmez, sorgu yazılmaz, tarama ve posta okuma yoktur.
  - Kaynak vektör, gövde parçalarının ortalamasıdır. Aday, en iyi parçasıyla puanlanır. Kaynak posta sonuçlarda yoktur.
  - Güven: kosinüs ≥ 0,85 Yüksek, ≥ 0,80 Orta, altı Zayıf (gri). Eşikler sentetik verideki ölçümden alındı (aşağıda).
  - Ara filtreleri benzer aramada da uygulanır.
  - Menü `ContextMenuMailItem` içindedir. `IRibbonControl.Context` bir `Selection` nesnesidir; ilk öğe okunur. Eklenti
    yalnız `EntryID`, `Parent.EntryID` ve `Parent.StoreID` değerlerini okur.
  - Posta indekste yoksa ya da anlam indeksi yoksa durum çubuğu Güncelle'yi önerir.
- **Ölçüm (21 Ekim, Mac, ONNX vektörleri, 310 posta, 32 konu ailesi, `eval/similar_eval.py`):** 100 aile postasında ilk 5
  sonucun 0,98'i aynı ailedendir. "Yüksek" etiketli 190 sonucun 172'si (0,905) aynı ailedendir. Kaynak posta hiçbir
  sonuçta yoktur.

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
    basılınca, ek okuma ayarı değişince ve hariç tutulan klasörler kaydedilince bir kez çalışır. Hariç tutulan
    klasörleri indeksten siler (bölüm 6); bütün klasörler hariç tutulmuşsa kapsamı boş olabilir.
- **Hazır model (20 Ekim, kullanıcı isteği):** pencere açıkken worker bir işten sonra kapanmaz, sonraki işi bekler; model
  bellekte kalır. Pencere açılınca yapılan durum okumasından sonra, kapsamda anlam indeksi varsa worker modeli arka planda
  yükler (`preload`). İlk arama da modeli beklemez. Bellek: pencere açıkken ~1–1,5 GB (2.19 ölçümü).
  23 Ekim'den beri yükleme ayrı bir iş parçacığında (thread) çalışır. Yükleme sürerken başlayan arama klasörleri tarar ve
  kelime sonuçlarını hemen gösterir; yalnız anlamsal adım yüklemenin bitmesini bekler. Model bir kez yüklenir.
- **Bellek önbelleği (23 Ekim):** hazır bekleyen worker, aranan klasörlerin vektör matrisini, kelime listesini (FTS
  sözlüğü) ve filtre listelerini (`done.facets`) bellekte tutar. İndeks dosyasındaki tetikleyiciler (trigger) her posta,
  parça ve vektör değişikliğinde `meta.data_gen` değerini artırır; değer değişince önbellek yeniden okunur. Sonuçlar yeni
  sorguyla aynıdır. Bellek: parça başına 2 KB (28.632 parça: 56 MB); kelime listesi 67 bin kelimede ~2 MB.
- **Bir aramanın aşamaları (23 Ekim):** kelime, erken anlamsal ve son liste aşaması sorgu kelimelerini, sorgu vektörünü
  ve bir posta metnindeki kelime eşleşmelerini bir kez hesaplar. Kelime kapsamı (sorgu kelimelerinin postada geçip
  geçmediği) yalnız ilk sonuçlara girebilecek adaylar için hesaplanır: adaylar, kapsam tam sayılınca alacakları puana
  göre sıralanır; bu puan sıradaki sonucun puanının altına düşünce okuma durur. Sonuçlar tüm adayları okumakla aynıdır.
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
- Süre günlüğü `data\timing.log`: her arama için tek bir JSON satırı. Satırda yalnız sayılar ve sabit sözcükler vardır:
  klasör başına manifest satır sayısı ve süresi (klasörler kapsam sırasıyla numaralı, ad yok), okuma ve ek kaydetme
  süresi, olay zamanları (`hello`, `need`, `keyword`, `semantic`, `final`, `done`), worker `done.timing` ve
  `done.stats` içindeki sayılar. Sorgu, posta metni, posta kimliği, klasör adı ve yol yazılmaz. İki dosya, her biri en
  çok 1 MB (`timing.log`, `timing.log.1`). Kaldırma `-KeepData` olmadan `data` klasörüyle birlikte siler.
- Eklenti postaları yalnız okur: `Send`, `Delete`, `Move`, `UnRead` gibi üyeleri hiç çağırmaz.
  `Body` okumak öğeyi okundu yapmaz; "Outlook'ta Aç" yalnız `Display()` çağırır. `Attachment.SaveAsFile` eki
  değiştirmez; kopyası yalnız worker okuyana kadar yerel veri klasöründe durur. "Eki aç" kopyası pencere kapanana kadar
  `data\open-att` klasöründe durur; açan program kopyayı tutuyorsa sonraki pencere açılışında silinir.
- Ek metni de posta metni gibi yalnız yerel indekste tutulur; log'a ya da başka bir yere yazılmaz.
- Her posta kutusu (store) ayrı SQLite dosyasıdır (`data\index\<store hash>.sqlite`).
- **İndeksi sil** (Ayarlar sekmesi, onaydan sonra): eklenti `data\index` ve `data\tmp-att` klasörlerini siler. Vektörler
  ve ek metni önbelleği (`att_texts`, `fts_att`) aynı SQLite dosyasındadır. Silmeden önce hazır tutulan worker kapanır.
  Oturum sürerken düğme kapalıdır; bu yüzden hiçbir worker dosyaları açık tutmaz. `settings.json` kalır. Worker ve
  protokol değişmez. Silmeden sonraki durum okuması yeni dosya açmaz (bilinmeyen posta kutusu kuralı).
- **Klasörleri hariç tutma (21 Ekim, kullanıcı isteği):** Ayarlar > **Klasörleri hariç tut…** penceresinde seçilen
  klasörler (alt klasörleriyle) İndeks listesinde görünmez ve hiçbir aramanın ya da indekslemenin kapsamına girmez.
  Seçim `settings.json` içinde `excluded_folders` (posta kutusu kimliği → klasör EntryID'leri) olarak durur; kayıt
  yoksa eski ad kuralı (Silinmiş/Deleted, Önemsiz/Gereksiz/Junk, Giden/Outbox, Taslak/Drafts, Eşitleme/Sync Issues)
  varsayılandır. Eklenti her oturumda listeyi `begin.exclude` ile gönderir. Worker oturum başında bu klasörlerin
  postalarını (`fts`, `fts_att`, `chunks`, `atts`), başka postanın kullanmadığı ek metnini ve vektörlerini siler. Metin
  dosyadan da çıkar: FTS5 `optimize` (silinen sözcükler segmentlerde kalır), sonra `VACUUM` ve
  `wal_checkpoint(TRUNCATE)`. 20 Ekim denemesi: `optimize` olmadan `VACUUM` sonrasında sözcük dosyada kaldı. Yarıda
  kalan sıkıştırmayı `meta.compact` sonraki oturumda yeniden başlatır. Kapsamla çakışan `exclude` protokol hatasıdır.
  Birim testi ham dosya baytlarını denetler.
- **Son ve kayıtlı aramalar (20 Ekim, kullanıcı isteği):** eklenti başlayan her aramanın metnini `data\searches.json`
  dosyasına yazar (en çok 20 son arama ve adlandırılmış kayıtlı aramalar). Öneriler Windows'un otomatik tamamlama
  listesiyle gösterilir. Dosya yalnız eklentide okunur ve yazılır; worker ve protokol değişmez. **İndeksi sil** bu dosyayı
  silmez; `-KeepData` olmadan kaldırma siler. Okunamayan (kilitli) dosyanın üstüne yazılmaz; yazma hatasında pencere
  dosyadaki durumu gösterir ve geçici kopya kalmaz.
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
| Güncelle, hiç tam indekslenmemiş klasördeki yeni postayı okumaz | Bu postalar sayılır; durum metni İndekslemeyi başlat'ı önerir | Klasör başına son okunan değişiklik zamanı (PST taşımasında davranış doğrulandıktan sonra) |
| Filtreli kelime araması dosyanın en iyi 400 sonucuyla sınırlı | Çok yaygın bir sözcükte, filtredeki bir posta yalnız anlamıyla bulunabilir | FTS sorgusuna rowid kümesi koşulu |
| Filtre listeleri indeks değişince bütün postaları tarar | 100.000 postada ~0,2 sn; değişiklik yoksa hazır worker listeyi bellekten verir | Kişiler tablosu |
| Benzerlik eşikleri kalibre edilmedi | Yalnız sentetik 32 aileyle ölçüldü | Gerçek kutudan etiketli örneklerle kalibrasyon |
| Sağ tık menüsü SSH ile sınanamaz | `Context` nesnesinin `Selection` olduğu bir topluluk yanıtına dayanır; yedek yol (doğrudan öğe) ve OlbCheck üyeleri vardır | Kullanıcı kabul adımı (`docs/KURULUM.md`) |
| Yazım hatası genişletmesi sözlükte doğrusal tarama | ~10^5 terimde birkaç yüz ms | Sözlük üzerinde trigram FTS |
| Alıntı/imza ayıklama sezgiseldir | Alışılmadık biçimler ana metinde kalabilir | Daha çok kalıp |
| Yeni Outlook desteklenmez | Yalnız Klasik Outlook | Bölüm 2'deki Graph + köprü yolu |
| Gerçek Outlook arayüz testi yapılmadı | SSH oturumu (Session 0) masaüstü arayüzü açamaz | Kullanıcı makinesinde el ile kabul testi (`docs/KURULUM.md`) |
