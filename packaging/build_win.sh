#!/bin/sh
# Build the Windows package dist/HatirMail-win-x64-<version>.zip on macOS (or any POSIX shell).
# Inputs (all pinned and hash-checked): embedded CPython 3.13.15 amd64, win_amd64 wheels (requirements-win.lock,
# packaging/make_locks.py), pip wheel, the ONNX text model exported from google/embeddinggemma-2@914f7f8
# (models/onnx, eval/export_onnx.py; parity with torch: eval/parity_onnx.py), the add-in DLL.
set -eu
cd "$(dirname "$0")/.."
VER=0.1.0
PYVER=3.13.15
PYZIP_SHA=d1f04d990aee1253d8569e8e5104e30fa9f5fa830899f14843448872d936a2cf
STAGE=build/pkg/HatirMail
OUT=dist/HatirMail-win-x64-$VER.zip
PY=${PYTHON:-.venv/bin/python}

rm -rf build/pkg && mkdir -p "$STAGE/runtime" "$STAGE/wheels" "$STAGE/app" "$STAGE/addin" dist

# 1. embedded Python (official python.org build), hash-checked
Z=build/python-$PYVER-embed-amd64.zip
[ -f "$Z" ] || curl -sSfL -o "$Z" "https://www.python.org/ftp/python/$PYVER/python-$PYVER-embed-amd64.zip"
echo "$PYZIP_SHA  $Z" | shasum -a 256 -c -
cp "$Z" "$STAGE/runtime/python-embed-amd64.zip"

# 2. wheels: exactly the files listed in the lock, re-verified
cp packaging/requirements-win.lock "$STAGE/runtime/requirements-win.lock"
"$PY" - "$STAGE" <<'EOF'
import hashlib, re, shutil, sys
from pathlib import Path
stage = Path(sys.argv[1])
lock = Path("packaging/requirements-win.lock").read_text()
want = dict(re.findall(r"^([A-Za-z0-9_.-]+)==([^\s\\]+)", lock, re.M))
hashes = set(re.findall(r"sha256:([0-9a-f]{64})", lock))
n = 0
for w in Path("build/wheels-win").glob("*.whl"):
    h = hashlib.sha256(w.read_bytes()).hexdigest()
    if h in hashes:
        shutil.copy2(w, stage / "wheels" / w.name); n += 1
assert n == len(want) == len(hashes), (n, len(want), len(hashes))
print("wheels:", n)
EOF
PIPW=build/pip-26.2.1-py3-none-any.whl
[ -f "$PIPW" ] || "$PY" -m pip download -q --no-deps --only-binary=:all: pip==26.2.1 -d build/
cp "$PIPW" "$STAGE/wheels/"
printf 'pip==26.2.1 \\\n    --hash=sha256:%s\n' "$(shasum -a 256 "$PIPW" | cut -d' ' -f1)" > "$STAGE/runtime/pip.lock"

# 3. worker code (no tests, no caches)
mkdir -p "$STAGE/app/worker"
cp worker/hm_*.py "$STAGE/app/worker/"

# 4. model (text component only, ONNX) + license notice
cp -R models/onnx "$STAGE/model"

# 5. add-in
# the same deterministic binary that the tests use (build/addin)
sh addin/build.sh >/dev/null
cp build/addin/HatirMail.dll "$STAGE/addin/"

# 6. scripts, docs, licenses
cp packaging/win/install.ps1 packaging/win/uninstall.ps1 packaging/win/settings.json "$STAGE/"
cp docs/KURULUM.md "$STAGE/KURULUM.md"
cp packaging/THIRD_PARTY_NOTICES.md "$STAGE/THIRD_PARTY_NOTICES.md"

# 7. manifest with sha256 of every file, then zip
( cd build/pkg/HatirMail && find . -type f ! -name MANIFEST.sha256 | LC_ALL=C sort | sed 's|^\./||' | \
  while read -r f; do printf '%s  %s\n' "$(shasum -a 256 "$f" | cut -d' ' -f1)" "$f"; done > MANIFEST.sha256 )
rm -f "$OUT"
( cd build/pkg && zip -qrX "../../$OUT" HatirMail )
shasum -a 256 "$OUT" | tee "$OUT.sha256"
du -h "$OUT"
