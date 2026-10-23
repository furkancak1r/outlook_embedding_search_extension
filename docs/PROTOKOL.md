# HatırMail eklenti ↔ worker protokolü (sürüm 1: ek okuma ve çalışma modları)

The attachment messages (8a–8c) are optional: an add-in without `attachments` in `begin` gets the same sequence as before.
`mode` is optional: without it the worker runs a search without a time limit (the behavior of the first version).
`keep_alive_s` and `preload` are optional: without them the process runs one session and exits.
`exclude` is optional: without it no folder is deleted from the index.
`filters` is optional: without it the search is not narrowed. Modes `update` and `similar` are new; an add-in that does not
send them gets the same behavior as before. `done.facets` and `results.similar_to` are additive fields. The `results`
stage `semantic` (step 8s) and `done.timing.scan_work_s`, `done.timing.semantic_s` are additive too.

The add-in starts `python.exe -I -B -X utf8 hm_worker.py --data <dir> --model <dir> --threads N` with the
environment variable `HM_SESSION=<48 hex chars>`. Both sides send one UTF-8 JSON object per line on the
child's stdin (add-in → worker, key `op`) and stdout (worker → add-in, key `ev`). There is no socket.

## Modes (`begin.mode`)

| Mode | Started by | What it does |
|---|---|---|
| `search` | "Ara" | Scan, read new mails, keyword results, attachments + embedding (time-limited by `index_budget_s`), final results |
| `index` | "İndekslemeyi başlat" (İndeks tab) | Scan, read, attachments, embedding without a time limit. No query, no results. When the read phase ends without `items_done.stopped`, every listed folder gets `folders.synced_gen` = this session |
| `update` | "Güncelle" (İndeks tab) | As `index`, but only the changes after the last index run: `need` = changed mails, and new entry ids only in folders with `synced_gen` (others: `stats.unsynced_skipped`). Attachments and embedding only for the delta (mails inserted or changed after `synced_gen` of their folder, or in this session). Deleted and moved-out mails leave the index (the scan, as in every mode). Never moves `synced_gen`. Extra stats: `delta`, `unsynced_skipped` |
| `status` | Window open, store change, "Durumu yenile", attachment setting changed, excluded folders saved | `hello` with the coverage, then `done`. No scan, no mail reading, no model. An unknown store gets no index file; an index of an older version is upgraded first (as in every mode). The scope can be empty (every folder excluded) |
| `similar` | "Benzerlerini bul" (Outlook shortcut menu of a mail) | `hello`, one `results` (stage `final`), `done`. `begin.similar_eid` (required, else `done/error` "similar_eid missing"): the mails nearest to this mail by its stored vectors. No scan, no mail reading, no model; an unknown store gets no index file |

## Filters (`begin.filters`, modes `search` and `similar`)

`{"sender": key or null, "recipient": key or null, "date_from": iso or null, "date_to": iso or null, "attach": kind or
null}`. Each key is optional. These values give `done/error` "bad filters": not an object, a value that is not a string
or null, a date that `datetime.fromisoformat` cannot read, an `attach` value that is not in the list below. Keys are the
first values of `done.facets` rows. The worker keeps only the mails of the scope that match every filter:

- `sender`: `fold(sender_addr or sender)` is the key.
- `recipient`: the key is one of the recipient names (To and CC split at `;`, quotes removed, folded).
- `date_from`, `date_to`: `date_from <= sent_at < date_to`. Both are local time; an offset is dropped (the add-in writes
  `sent_at` in local time). A mail without a date does not match a date filter.
- `attach`: `yes` (at least one attachment name), `no` (none), or a type: `excel` (.xlsx .xlsm .xls .csv), `pdf`, `word`
  (.docx .doc .docm .rtf), `powerpoint` (.pptx .ppt .pptm), `image` (.png .jpg .jpeg .gif .bmp .tif .tiff .heic),
  `archive` (.zip .rar .7z). The worker reads `items.attachments`: inline logos and pictures (image001.png) are not in it.

Keyword candidates are the best 400 BM25 rows of the file inside the allowed set; the semantic z statistics use the
unfiltered scope, so the confidence labels do not change. The folder filter of the Ara tab is the scope (`folders`).
Every result has `sender_key` (the `sender` filter key of its sender): the add-in uses it for "Yalnız bu gönderenle
yeniden ara".

`done.facets` (all modes, when the index file is open): `{"senders": [[key, name, address, mails], …], "recipients": [[key,
name, mails], …]}` of the whole index file, most mails first, at most 1000 each.

## Similar mails (`similar`)

Source vector: the normalized mean of the body chunk vectors of `similar_eid` (all its chunks when it has no body chunk).
A candidate scores with its best chunk (body or attachment). The source mail is never a result. `confidence`: `yüksek`
(cos ≥ 0.85), `orta` (≥ 0.80), `zayıf`. `why`: `["benzerlik 0.91"]` (+ `"ekte"` when the best chunk is in an attachment).
`hl`, `hl_subject` are empty. `results.similar_to`: `null` (the mail is not in the index), `{"eid","subject","embedded":
false}` (no vectors yet: no results), `{"eid","subject","embedded": true}`.

## Excluded folders (`begin.exclude`)

`begin.exclude` (all modes): folder entry ids that the user excluded from the index (Ayarlar "Klasörleri hariç tut…",
subfolders included by the add-in). At the session start, before `hello`, the worker deletes the mails of these folders
(keyword rows, chunks, attachments), their folder rows, the attachment text and vectors that no other mail uses, then
runs FTS5 `optimize`, `VACUUM` and `wal_checkpoint(TRUNCATE)`, so the text also leaves the file. `done.excluded_removed`
= the number of deleted mails (0 when there was nothing to delete). A folder in both `exclude` and `folders`:
`done/error` "excluded folder in scope". An empty `folders` list is allowed only in `status` mode.

## Keep-alive (one process, many sessions)

`begin.keep_alive_s` (all modes; missing or `0` = the process exits after `done`, the first behavior): after `done`
with status `ok` the worker waits up to `keep_alive_s` seconds for the next `begin` on the same stdin. The loaded model
stays in memory, so the next session does not load it (`hello.model_ready=true`, `done.timing.model_warm=true`). The
add-in sends its setting + 600 s and closes the worker itself at the end of its own keep-alive time (stdin closes). A
cancel or an error still ends the process. No `begin` in time or stdin closed: exit 0.

`begin.preload` (status mode with `keep_alive_s` > 0): after `done`, when the scope has vectors, the worker loads the
model between sessions and sends `{"ev":"model","state":"loading"}`, then `{"ev":"model","state":"ready","load_s",
"peak_rss_mb"}` or `{"ev":"model","state":"error","error":<type>}`. A `begin` that arrives during the load waits until
the load ends. Without vectors (no semantic index) nothing is loaded.

The session token (`HM_SESSION`) is the same for all sessions of one process.

`index_budget_s` (search only): seconds of indexing work after the read phase starts. The add-in stops reading new
mails after half of it and sends `items_done.stopped=true`; the worker stops attachment reading and embedding at the
limit. The model load of a search is not counted (the search needs the model anyway). `0` = no indexing work. Missing =
no limit. When the scope has no semantic index yet, a time-limited search loads no model and returns keyword results.

## Sequence (search and index)

| # | Direction | Message | Notes |
|---|---|---|---|
| 1 | → | `{"op":"begin","session","mode","query","now","top_k","index_budget_s","idle_timeout_s","keep_alive_s","preload","attachments","att_dir","store":{"id","name"},"folders":[{"id","path"}],"exclude":[folder id…],"filters":{"sender","recipient","date_from","date_to","attach"},"similar_eid"}` | `session` must equal `HM_SESSION`, else `done/error` and no file is written. Unknown `mode`: `done/error`. `attachments`: true when the add-in can save attachments into `att_dir` (an existing folder) |
| 2 | ← | `{"ev":"hello","pid","proto","index_version","model_rev","mode","coverage","model_ready"}` | `status` mode: `done` follows at once. `model_ready`: the model of this process is loaded already (keep-alive) |
| 3 | → | `{"op":"manifest","folder","items":[[eid, last_mod, class, date], …],"final":bool}` | pages of ≤ 500 rows; `final` closes a folder. Changed mails are marked "to read" at once |
| 4 | → | `{"op":"scan_done"}` | |
| 5 | ← | `{"ev":"need","eids":[…]}` | new or changed items, newest first; deleted/moved-out items are removed now (items of the listed folders that no manifest of this scan holds). A known item in the manifest of another folder moves there and is not read again. `done.timing.scan_s` = from the first manifest wait to `need` (the add-in time included); `scan_work_s` = the worker part of it |
| 6 | → | `{"op":"items","items":[{eid,folder,cls,mod,subject,sender,sender_addr,to,cc,date,attachments[],att_meta[[index,name,size,type]],body}]}` | pages of 25. `index` is 1-based; `type` is the Outlook `OlAttachmentType` (1 = file, 5 = attached item) |
| 7 | → | `{"op":"items_done","failed":[eid…],"stopped":bool}` | always sent, also after an empty `need`. `stopped`: the read time limit ended the reading; the rest stays "to read" and the next scan asks for it again |
| 8 | ← | `{"ev":"results","stage":"keyword",…}` | search only: keyword results, no model needed |
| 8s | ← | `{"ev":"results","stage":"semantic",…}` | search only, optional: hybrid results with the vectors that the index has now. Sent only when the search still has work (a chunk without a vector, or a pending attachment with `attachments`), the query is not empty, `index_budget_s` is not `0` and the scope has vectors. The `final` results (step 10) replace them. The deadline grows by the time of this step |
| 8a | ← | `{"ev":"need_att","atts":[[eid,index,name],…]}` | only with `attachments`; at most 4 per request, newest mails first |
| 8b | → | `{"op":"att_files","files":[[eid,index,path or null],…]}` | `path` must be a file directly in `att_dir` (else `done/error`). The worker reads and deletes each file. `null`: the add-in could not save it |
| 8c | ← | `{"ev":"progress","phase":"att",…}` | 8a–8c repeat until all pending attachments are read, or cancel/finish/time limit |
| 9 | ← | `{"ev":"progress","phase":"model"|"embed",…}` | the model loads here (ONNX Runtime, in this process only), unless it is loaded already. `embed` counts model work: distinct chunks without a vector |
| 10 | ← | `{"ev":"results","stage":"final","results":[…],"coverage","date_windows","terms","semantic","sem_stats"}` | search only: top 10 |
| 11 | ← | `{"ev":"done","status":"ok"|"cancelled"|"error","mode","coverage","facets","timing","stats","peak_rss_mb","model_loaded","excluded_removed"}` | `coverage` = state after the work (also after cancel and error, when the index was open). Then the process exits, or waits for the next `begin` (keep-alive) |

Progress events: `{"ev":"progress","phase":"scan"|"read"|"att"|"model"|"embed","done","total","rate","eta_s"[,"coverage"]}`.
`rate` = items per second in this phase (embed: chunks per second, model load not counted). `eta_s` = seconds left in
this phase at this rate (`null` before the first measurement, `0` when done). `coverage` comes at most every 2 s.

At any time the add-in can send `{"op":"cancel"}` (stop, exit code 3) or `{"op":"finish"}` (stop indexing, return
results for the indexed part). If stdin closes, the worker exits (code 3). Vectors are saved after every model batch
(8 chunks): a killed worker loses at most one batch.

## Result object

`eid, folder, subject, sender, sender_key, recipients, date, snippet, hl, hl_subject, match_attachment, match_att_index,
match_att_member, attachments, score, confidence ("yüksek"|"orta"|"zayıf"|"kelime"), why[], cos, z, coverage`

`sender_key`: the `begin.filters.sender` key of the sender of this mail (see "Filters").

`hl`, `hl_subject`: `[[start, length], …]` of the query words in `snippet` and `subject`, in UTF-16 code units (the add-in
marks them in the preview). `match_attachment`: the attachment name when the snippet comes from an attachment, else `null`.
`match_att_index`: the 1-based Outlook index of that attachment (`att_meta`), else `null`. `match_att_member`: for a zip
attachment, the path of the member that holds the snippet, as Python's `zipfile` names it (`/` separators; names without
the UTF-8 flag in code page 437), else `null`. The add-in uses these two for "Eki aç"; an add-in that does not know them
ignores them.

## Coverage object

`folders[{id,path,total,indexed,read,embedded,att_total,att_done,scanned_at}], total, indexed, read, embedded, att_total,
att_done, complete`

- `total`: searchable mails in the last scan of the folder (`null`: never scanned). `indexed`: mails in the index.
  `read`: of them, read with their current version (a changed mail that is not read again yet is searchable with its
  old text). `embedded`: mails with all their chunks in the semantic index.
- `att_total`/`att_done`: readable attachments of the indexed mails / read ones (0 without `attachments` in `begin`).
- `complete`: every folder scanned, all mails read and embedded, all readable attachments read.
- `folders` keeps the order of `begin.folders`.

## Exit codes

0 = ok (also: keep-alive ended, stdin closed between sessions), 2 = error (only the exception type is reported, never a
message with data), 3 = cancelled / parent gone during a session.
