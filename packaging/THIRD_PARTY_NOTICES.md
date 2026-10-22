# Third-party components in the HatırMail Windows package

| Component | Version | License | Source |
|---|---|---|---|
| EmbeddingGemma 2 (text component only) | revision 914f7f89142e33e77833254d9c9b90c3cef7303b | Apache-2.0 | https://huggingface.co/google/embeddinggemma-2 |
| CPython (embeddable package, amd64) | 3.13.15 | PSF-2.0 | https://www.python.org/ftp/python/3.13.15/ |
| ONNX Runtime (CPU; telemetry turned off by HatırMail, see below) | 1.31.0 | MIT | https://pypi.org/project/onnxruntime/ |
| tokenizers | 0.23.3 | Apache-2.0 | https://pypi.org/project/tokenizers/ |
| numpy | 2.5.4 | BSD-3-Clause | https://pypi.org/project/numpy/ |
| protobuf, flatbuffers (ONNX Runtime dependencies) | 7.36.2, 25.12.19 | BSD-3-Clause, Apache-2.0 | https://pypi.org/ |
| huggingface-hub | 2.2.0 | Apache-2.0 | https://pypi.org/project/huggingface-hub/ |
| pypdfium2 (PDF text and page rendering; includes PDFium) | 5.14.0 | BSD-3-Clause, Apache-2.0 (PDFium and its dependencies: see the wheel's licenses folder) | https://pypi.org/project/pypdfium2/ |
| xlrd (legacy .xls) | 2.0.2 | BSD | https://pypi.org/project/xlrd/ |
| olefile (Outlook .msg) | 0.47 | BSD-2-Clause | https://pypi.org/project/olefile/ |
| winrt-runtime and winrt-Windows.* (Windows OCR API binding) | 3.2.1 | MIT | https://pypi.org/project/winrt-runtime/ |
| other transitive wheels | see runtime/requirements-win.lock | each wheel holds its license file | https://pypi.org/ |
| NSIS (the setup program `HatirMail-Kurulum-*.exe` only: installer stub, plug-ins, LZMA module) | 3.13 | zlib/libpng; LZMA module: CPL-1.0 with the NSIS linking exception | https://nsis.sourceforge.io/ |

OCR uses `Windows.Media.Ocr`, a part of Windows 10/11. No OCR model is in this package; Windows uses the OCR
languages of the user profile.

The model files `model/model.onnx` and `model/model.onnx.data` hold only the text component (`language_model.*`
tensors) of the official checkpoint, in float32 (exact conversion of the bf16 values), exported to ONNX with the
pooling and normalization steps. The vision and audio tensors were removed. See `model/NOTICE.txt`.

ONNX Runtime official builds collect usage events by default (Microsoft privacy statement:
https://go.microsoft.com/fwlink/?LinkID=824704). HatırMail sets `ORT_DISABLE_TELEMETRY=1` before ONNX Runtime starts
and calls `disable_telemetry_events()`: no events, no uploader, no device identifier.
Use of the model must follow the Gemma Prohibited Use Policy: https://ai.google.dev/gemma/prohibited_use_policy
