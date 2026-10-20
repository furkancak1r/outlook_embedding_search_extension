"""P1 probe: load the text-only model, report parameters, timings, peak RSS and a TR/EN similarity sanity check."""
import json, os, resource, sys, time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "worker"))
t0 = time.perf_counter()
import hm_embed  # noqa: E402

model_dir = sys.argv[1] if len(sys.argv) > 1 else str(ROOT / "models" / "embeddinggemma-2")
threads = int(sys.argv[2]) if len(sys.argv) > 2 else 4
enc = hm_embed.Encoder(model_dir, threads)
t_load = time.perf_counter() - t0

queries = ["geçen ay gelen fatura iadesi", "server maintenance plan", "toplanti notlari", "sozlesme taslagi"]
docs = [
    hm_embed.doc_text("Fatura iadesi hk.", "Merhaba, geçen ayki faturanın iadesi için onay bekliyoruz."),
    hm_embed.doc_text("Sunucu bakım planı", "Cumartesi gecesi sunucularda planlı bakım yapılacaktır."),
    hm_embed.doc_text("Toplantı notları", "Bugünkü toplantıda bütçe ve takvim konuşuldu."),
    hm_embed.doc_text("Sözleşme taslağı", "Ekte yeni tedarikçi sözleşmesinin taslağını bulabilirsiniz."),
]
t1 = time.perf_counter()
q = enc.encode([hm_embed.query_text(x) for x in queries], hm_embed.MAX_QUERY_TOKENS)
d = enc.encode(docs)
t_enc = time.perf_counter() - t1
sim = (q @ d.T).round(3).tolist()

long_doc = hm_embed.doc_text("Uzun metin", " ".join(["Proje toplantısında bütçe, takvim ve riskler konuşuldu."] * 30))
t2 = time.perf_counter()
enc.encode([long_doc] * 16)
t_batch = time.perf_counter() - t2
ntok = len(enc.tok(long_doc)["input_ids"])

rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
rss_mb = rss / 2**20 if sys.platform == "darwin" else rss / 1024
print(json.dumps({
    "params": enc.params, "dim": q.shape[1], "threads": threads, "load_s": round(t_load, 2),
    "encode_8_s": round(t_enc, 3), "batch16_tokens_each": ntok, "batch16_s": round(t_batch, 2),
    "peak_rss_mb": round(rss_mb), "sim_query_x_doc": sim,
    "diag_is_max": all(row.index(max(row)) == i for i, row in enumerate(sim)),
}, ensure_ascii=False))
