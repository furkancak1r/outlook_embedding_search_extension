"""Parity check: HatırMail encoder vs the official sentence-transformers path (full model, prompt names).

Run in a separate venv with sentence-transformers==6.1.0 (dev check only, not a runtime dependency).
Usage: python parity_st.py <official_model_dir> <text_only_model_dir>
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "worker"))
import numpy as np  # noqa: E402
from sentence_transformers import SentenceTransformer  # noqa: E402

import hm_embed  # noqa: E402

queries = ["geçen ay gelen fatura iadesi", "server maintenance next weekend"]
docs = ["Cumartesi gecesi sunucularda planlı bakım yapılacaktır.", "The invoice refund was approved last month."]
import torch  # noqa: E402

# The model card: use float32 on CPU (the checkpoint dtype is bfloat16, which transformers v5 loads by default).
st = SentenceTransformer(sys.argv[1], device="cpu", local_files_only=True,
                         config_kwargs={"vision_config": None, "audio_config": None},
                         model_kwargs={"dtype": torch.float32})
q_st = st.encode(queries, prompt_name="SearchQuery", truncate_dim=hm_embed.DIM, normalize_embeddings=True)
d_st = st.encode(docs, prompt_name="Document", truncate_dim=hm_embed.DIM, normalize_embeddings=True)
enc = hm_embed.Encoder(sys.argv[2], 4)
q_hm = enc.encode([hm_embed.query_text(q) for q in queries], hm_embed.MAX_QUERY_TOKENS)
d_hm = enc.encode([hm_embed.doc_text(None, d) for d in docs])
res = {"st_prompts": {"SearchQuery": st.prompts.get("SearchQuery"), "Document": st.prompts.get("Document")},
       "st_dtype": str(next(st.parameters()).dtype), "st_norms": [float(x) for x in np.linalg.norm(q_st, axis=1)],
       "max_abs_diff_query": float(np.abs(q_st - q_hm).max()), "max_abs_diff_doc": float(np.abs(d_st - d_hm).max()),
       "cos_query": [float(x) for x in (q_st * q_hm).sum(1)], "cos_doc": [float(x) for x in (d_st * d_hm).sum(1)]}
res["pass"] = min(res["cos_query"] + res["cos_doc"]) > 0.9999
print(json.dumps(res, ensure_ascii=False))
sys.exit(0 if res["pass"] else 1)
