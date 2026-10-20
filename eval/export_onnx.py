"""Export the EmbeddingGemma 2 text model to ONNX (once, on the dev machine; not part of the package).

Graph: input_ids, attention_mask (int64, [batch, tokens]) -> sentence_embedding (float32, [batch, 512]).
Inside: the text model (fp32, as hm_embed loads it), mean pooling over the mask, Matryoshka cut to 512, L2 norm.
The worker then needs only onnxruntime + tokenizers + numpy (no torch, no transformers at run time).
Usage: .venv-onnx/bin/python eval/export_onnx.py models/text-only models/onnx
Check: eval/parity_onnx.py (torch vs ONNX Runtime).
"""
import hashlib
import json
import shutil
import sys
import time
from pathlib import Path

import torch
from transformers import AutoModel

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "worker"))
import hm_embed  # noqa: E402

src, dst = Path(sys.argv[1]), Path(sys.argv[2])
model = AutoModel.from_pretrained(src, local_files_only=True, vision_config=None, audio_config=None, dtype=torch.float32).eval()


class Embed(torch.nn.Module):
    def __init__(self, m):
        super().__init__()
        self.m = m

    def forward(self, input_ids, attention_mask):
        h = self.m(input_ids=input_ids, attention_mask=attention_mask).last_hidden_state
        w = attention_mask.unsqueeze(-1).to(h.dtype)
        pooled = (h * w).sum(1) / w.sum(1).clamp(min=1)
        return torch.nn.functional.normalize(pooled[:, :hm_embed.DIM], dim=-1)


dst.mkdir(parents=True, exist_ok=True)
ids = torch.tensor([[2, 1000, 2000, 3000, 1], [2, 1500, 1, 0, 0]], dtype=torch.long)
mask = (ids != 0).long()
batch, tokens = torch.export.Dim("batch", min=1, max=64), torch.export.Dim("tokens", min=2, max=hm_embed.MAX_DOC_TOKENS)
t = time.perf_counter()
with torch.inference_mode():
    prog = torch.onnx.export(Embed(model), (ids, mask), input_names=["input_ids", "attention_mask"], output_names=["sentence_embedding"],
                             dynamic_shapes={"input_ids": {0: batch, 1: tokens}, "attention_mask": {0: batch, 1: tokens}},
                             dynamo=True, external_data=True, optimize=True)
prog.save(str(dst / "model.onnx"), external_data=True)
print("export_s", round(time.perf_counter() - t, 1))
for f in ("tokenizer.json", "README.md"):
    shutil.copy2(src / f, dst / f)
(dst / "NOTICE.txt").write_text(
    (src / "NOTICE.txt").read_text().splitlines()[0] + "\n"
    "Source: https://huggingface.co/google/embeddinggemma-2 revision %s\n"
    "Modified by HatirMail: the text component only (the language_model.* tensors; vision and audio removed), loaded\n"
    "in float32 and exported to ONNX (model.onnx + model.onnx.data) with mean pooling, Matryoshka truncation to %d\n"
    "dimensions and L2 normalization in the graph (eval/export_onnx.py). The weights are the source bf16 values in\n"
    "float32 (exact conversion).\n" % (hm_embed.MODEL_REVISION, hm_embed.DIM))
info = json.loads((src / "hatirmail-model.json").read_text())
info["format"] = "onnx"
info["graph"] = "input_ids, attention_mask -> mean pool -> MRL %d -> L2 norm (eval/export_onnx.py)" % hm_embed.DIM
info["exported_with"] = {"torch": torch.__version__, "transformers": __import__("transformers").__version__,
                         "onnx": __import__("onnx").__version__, "onnxscript": __import__("onnxscript").__version__}
info["files"] = {p.name: {"size": p.stat().st_size, "sha256": hashlib.sha256(p.read_bytes()).hexdigest()}
                 for p in sorted(dst.iterdir()) if p.is_file() and p.name != "hatirmail-model.json"}
(dst / "hatirmail-model.json").write_text(json.dumps(info, indent=1) + "\n")
print(json.dumps(info["files"], indent=1))
