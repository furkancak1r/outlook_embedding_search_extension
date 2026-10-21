# HatırMail

Klasik Outlook için yerel, hatırlamaya dayalı e-posta araması. Kullanıcı hatırladığı birkaç ayrıntıyı Türkçe ya da
İngilizce yazar; eklenti ilgili gerçek e-postaları bulur ve Outlook'ta açar. Eklerin içeriği de aranır (Excel, Word,
PowerPoint, PDF, resimler Windows OCR ile, eklenmiş iletiler, zip). Yapay zekâ (EmbeddingGemma 2, yalnız 270M metin
bileşeni, ONNX Runtime) kullanıcının kendi bilgisayarında çalışır. Model arama penceresi açılınca yüklenir ve
pencere açıkken hazır kalır; aramalar modeli beklemez. Pencere kapanınca model de kapanır. Pencerede üç sekme vardır: **Ara**, **İndeks** (indeks durumu, yüzdeler, tam indekslemeyi başlatma) ve
**Ayarlar** (indekslenmeyecek klasörleri seçme dahil).

| Belge | İçerik |
|---|---|
| [docs/KURULUM.md](docs/KURULUM.md) | Kurulum, kullanım, ayarlar, kaldırma |
| [docs/MIMARI.md](docs/MIMARI.md) | Mimari kararlar, model doğrulaması, gizlilik, bilinen sınırlar |
| [docs/PROTOKOL.md](docs/PROTOKOL.md) | Eklenti ↔ worker JSON-satır protokolü |

## Klasör yapısı

| Yol | İçerik |
|---|---|
| `addin/` | C# COM eklentisi (`HatirMail.cs`, `SearchForm.cs`), Outlook'suz test koşucusu (`TestHost.cs`), tür kitaplığı denetimi (`OlbCheck.cs`), `build.sh` |
| `worker/` | Python worker: metin hazırlığı, tarih ayrıştırma, ek metni okuma (`hm_attach.py`), SQLite FTS5 + vektör deposu, hibrit arama, protokol; `tests/` |
| `eval/` | Sentetik korpus ve ek dosyası üreticileri (`att/`), uçtan uca değerlendirme, kalibrasyon, model doğrulama, ONNX dışa aktarma ve eşdeğerlik (`export_onnx.py`, `parity_onnx.py`) |
| `packaging/` | Model kilidi ve indirme, yalnız metin model türetimi, hash kilitli bağımlılıklar (`make_locks.py`), Windows paket üreticisi ve kurulum betikleri |

## Geliştirici komutları (macOS)

```sh
mkdir -p reports                                                         # local results (git-ignored)
python3.13 -m venv .venv                                                 # runtime venv: ONNX Runtime, no torch
.venv/bin/python packaging/make_locks.py                                 # wheels + hashed locks (checked against PyPI)
.venv/bin/python -m pip install --no-index --find-links build/wheels-mac --require-hashes -r packaging/requirements-mac.lock
python3.13 packaging/fetch_model.py models/embeddinggemma-2              # pinned revision, sha256-checked
python3.13 packaging/make_text_only.py models/embeddinggemma-2 models/text-only
# ONNX export + parity (dev venv .venv-onnx: torch 2.14.1, transformers 5.19.0, onnx 1.23.2, onnxscript 0.7.2, onnxruntime 1.31.0)
.venv-onnx/bin/python eval/export_onnx.py models/text-only models/onnx
.venv-onnx/bin/python eval/parity_onnx.py models/text-only models/onnx reports/parity-mac.json 4
.venv/bin/python -m unittest discover -s worker/tests                    # 51 tests (1 is Windows OCR only), no model
python3.13 eval/make_corpus.py                                           # deterministic synthetic data
.venv/bin/python eval/run_eval.py --model models/onnx --data tmp/eval --out reports/eval-mac.json
.venv/bin/python eval/run_eval.py --att --model models/onnx --data tmp/eval-att --out reports/eval-att-mac.json
sh addin/build.sh                                                        # mono csc, .NET Framework 4.8 refs
sh packaging/build_win.sh                                                # dist/HatirMail-win-x64-0.1.0.zip
```
