"""Check that the text-only model gives the same embeddings as the official files (text-only load)."""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "worker"))
import numpy as np  # noqa: E402

import hm_embed  # noqa: E402

texts = [hm_embed.query_text("geçen ay gelen fatura iadesi"),
         hm_embed.doc_text("Sunucu bakım planı", "Cumartesi gecesi sunucularda planlı bakım yapılacaktır."),
         hm_embed.doc_text("none", "Meeting notes: budget and timeline were discussed.")]
a = hm_embed.Encoder(str(ROOT / "models" / "embeddinggemma-2"), 4).encode(texts)
b = hm_embed.Encoder(str(ROOT / "models" / "text-only"), 4).encode(texts)
diff = float(np.abs(a - b).max())
print(json.dumps({"max_abs_diff": diff, "identical": diff == 0.0}))
sys.exit(0 if diff == 0.0 else 1)
