"""EmbeddingGemma 2 text-only encoder for CPU (ONNX Runtime).

The model is the text component of google/embeddinggemma-2 (vision_config=None, audio_config=None, fp32), exported once
to ONNX with mean pooling over all tokens, Matryoshka truncation to DIM and L2 normalization inside the graph
(eval/export_onnx.py). The worker needs only onnxruntime + tokenizers + numpy: no torch, no transformers (their import
took most of the load time). TorchEncoder is the reference path (eval/parity_onnx.py checks that both give the same
vectors); the worker never uses it.
The tokenizer is the official tokenizer.json, used through the `tokenizers` library directly.
"""
import os
from pathlib import Path

MODEL_REPO = "google/embeddinggemma-2"
MODEL_REVISION = "914f7f89142e33e77833254d9c9b90c3cef7303b"
DIM = 512  # MRL: the card reports MTEB multilingual 61.17 at 512d vs 61.36 at 768d.
QUERY_PREFIX = "task: search result | query: "
MAX_DOC_TOKENS = 512
MAX_QUERY_TOKENS = 128
PAD_ID = 0


def doc_text(title, text):
    """Document format from the model card: 'title: {title} | text: {content}', 'none' without a title."""
    title = " ".join((title or "").split()) or "none"
    return f"title: {title} | text: {text}"


def query_text(query):
    return QUERY_PREFIX + query


class FakeEncoder:
    """Test-only encoder (worker flag --fake-encoder): hashed bag of folded words. No model."""

    params = 0

    def __init__(self, *_):
        import numpy as np
        self.np = np

    def encode(self, texts, max_tokens=MAX_DOC_TOKENS):
        import hashlib
        import re
        np = self.np
        out = np.zeros((len(texts), DIM), np.float32)
        for i, t in enumerate(texts):
            for w in re.findall(r"\w{3,}", t.lower())[:max_tokens]:
                out[i, int(hashlib.md5(w.encode()).hexdigest()[:6], 16) % DIM] += 1.0
        out += 1e-3
        return out / np.linalg.norm(out, axis=1, keepdims=True)


class _Tokens:
    def __init__(self, model_dir):
        from tokenizers import Tokenizer
        self.tok = Tokenizer.from_file(str(Path(model_dir) / "tokenizer.json"))

    def batch(self, texts, max_tokens):
        import numpy as np
        self.tok.enable_truncation(max_tokens)
        self.tok.enable_padding(pad_id=PAD_ID, pad_token="<pad>", direction="right")
        encs = self.tok.encode_batch(list(texts))
        return (np.array([e.ids for e in encs], dtype=np.int64), np.array([e.attention_mask for e in encs], dtype=np.int64))


def session(path, threads):
    """ONNX Runtime CPU session with telemetry off. Official builds (1.29+) send usage events and keep a device id under
    the user profile by default; ORT_DISABLE_TELEMETRY=1 is read once, when the runtime starts (the add-in also sets it
    in the worker environment)."""
    os.environ["ORT_DISABLE_TELEMETRY"] = "1"
    import onnxruntime as ort
    ort.disable_telemetry_events()
    so = ort.SessionOptions()
    so.intra_op_num_threads = max(1, int(threads))
    so.inter_op_num_threads = 1
    so.log_severity_level = 3  # errors only (stderr; never mail text)
    so.graph_optimization_level = ort.GraphOptimizationLevel.ORT_ENABLE_ALL
    return ort.InferenceSession(str(path), so, providers=["CPUExecutionProvider"])


class Encoder(_Tokens):
    """The worker's encoder: ONNX Runtime on the CPU."""

    params = None

    def __init__(self, model_dir, threads=4):
        super().__init__(model_dir)
        self.sess = session(Path(model_dir) / "model.onnx", threads)

    def encode(self, texts, max_tokens=MAX_DOC_TOKENS):
        ids, mask = self.batch(texts, max_tokens)
        return self.sess.run(["sentence_embedding"], {"input_ids": ids, "attention_mask": mask})[0]


class TorchEncoder(_Tokens):
    """Reference path (torch + transformers on the safetensors model). For eval/parity_onnx.py only."""

    def __init__(self, model_dir, threads=4):
        for k in ("HF_HUB_OFFLINE", "TRANSFORMERS_OFFLINE", "HF_HUB_DISABLE_TELEMETRY"):
            os.environ[k] = "1"  # never contact the network
        import torch
        from transformers import AutoModel

        torch.set_num_threads(max(1, int(threads)))
        self.torch = torch
        super().__init__(model_dir)
        self.model = AutoModel.from_pretrained(
            model_dir, local_files_only=True, vision_config=None, audio_config=None, dtype=torch.float32)
        self.model.eval()
        if getattr(self.model, "vision_tower", None) is not None or getattr(self.model, "audio_tower", None) is not None:
            raise RuntimeError("vision/audio encoder loaded; expected the text-only model")
        self.params = sum(p.numel() for p in self.model.parameters())

    def encode(self, texts, max_tokens=MAX_DOC_TOKENS):
        torch = self.torch
        ids, mask = (torch.from_numpy(a) for a in self.batch(texts, max_tokens))
        with torch.inference_mode():
            hidden = self.model(input_ids=ids, attention_mask=mask).last_hidden_state
            m = mask.unsqueeze(-1).to(hidden.dtype)
            pooled = (hidden * m).sum(1) / m.sum(1).clamp(min=1)
            vec = torch.nn.functional.normalize(pooled[:, :DIM], dim=-1)
        return vec.float().numpy()
