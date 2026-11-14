#!/usr/bin/env bash
# ============================================================================
# DRAFT ONLY - NOT EXECUTED IN THIS ROUND. Requires Gate B' / Una authorization.
# Excel Privacy brand raster assets - GPT Review R3.2 revision.
#
# R3.2 P1-B (ROLLBACK IDENTITY PROTECTION) -- the ONLY change in this round:
#   R3.1 guarded rollback with the current SHA-256 alone. A different process
#   can REPLACE a published file with an inode of IDENTICAL content; the hash
#   still matches and R3.1 would delete a file it no longer owns. R3.2 freezes
#   a 3-part identity at publish time (SHA-256 + st_dev + st_ino) and deletes
#   a published file ONLY if ALL THREE still match. temp-file cleanup keeps
#   the R3.1 "this run registered this path" rule (plus a plain-file check).
#   PREPARE / GENERATE / GENERATE-README / atomic `ln` COMMIT / PREFLIGHT are
#   INHERITED UNCHANGED from R3.1 -- no capability is redeveloped here.
#
# INHERITED from R3 (unchanged capabilities, NOT redeveloped):
#   - PREPARE: 4 registered files copied from the package, SHA-256 verified.
#   - GENERATE: Master check + verbatim copy + 1024 PNG + 10-slot ICNS +
#     7-size ICO built by the STAGED build_ico.py (R3 PNG decoder).
#   - 1024 PNG dimension/colour check; ICNS reverse-decode != visual check;
#     4 image + 4 registered SHA-256 output; no-overwrite PREFLIGHT.
#
# R3.1 P1-B (COMMIT atomicity / rollback):
#   1) Each file is first copied to a RUN-UNIQUE TEMP FILE IN THE SAME
#      DIRECTORY ("<rel>.__r31_tmp__.<pid>"). Its SHA-256 is verified BEFORE
#      publishing. A `cp` that writes partial bytes and then fails can only
#      leave that temp file - NEVER a half-written final asset.
#   2) Publishing uses `ln` (hard link, same filesystem = atomic) which by
#      definition does NOT overwrite an existing target. The temp link is
#      unlinked afterwards; the final path keeps the inode.
#   3) rollback removes only (a) temp files THIS run registered, and
#      (b) final files THIS run published AND whose CURRENT SHA-256 still
#      equals what this run wrote. A file changed by another process/user is
#      NOT deleted (WARN instead). Pre-existing files are never touched.
#   4) A later run in the same directory is safe after rollback.
#
# R3.1 P1-C (README auto-generation, replaces PENDING):
#   5) After the 4 image assets exist, GENERATE-README reads their REAL
#      SHA-256 values, substitutes the 4 "@@SHA256:...@@" placeholders in the
#      STAGED README template, computes README_FINAL_SHA256, asserts
#      TEMPLATE_SHA != FINAL_SHA and no placeholder remains, then verifies
#      ALL 8 BASE_FILES have a final SHA-256. NO manual repo edit is needed.
#
# Constraints: no redesign, no recolor, no crop, no matting, no background
# removal, no tracing, no auto-vectorization. Only equal-ratio sips resize
# and lossless container conversions (ICNS / ICO).
# Toolchain: macOS sips + iconutil + Python standard library. No Pillow/IM.
#
# Usage:
#   build_icons.sh <master.png> <repo-root-of-isolated-branch> <r3.1-package-root>
#   <r3.1-package-root> must contain:
#     drafts/BRAND_ASSETS.md
#     drafts/assets-brand-README.md          (TEMPLATE with 4 placeholders)
#     scripts/build_icons.sh
#     scripts/build_ico.py
# ============================================================================
set -Eeuo pipefail

MASTER="${1:?usage: build_icons.sh <master.png> <repo-root> <r3.1-package-root>}"
REPO="${2:?usage: build_icons.sh <master.png> <repo-root> <r3.1-package-root>}"
PKG="${3:?usage: build_icons.sh <master.png> <repo-root> <r3.1-package-root>}"

EXPECTED_SHA="0be01b184f5298ba73b5012fff09c18a76a9d149f0e6833baa12ff4be139b287"
EXPECTED_DIM=1254

REGISTERED_SRC=(
  "$PKG/drafts/BRAND_ASSETS.md"
  "$PKG/drafts/assets-brand-README.md"
  "$PKG/scripts/build_icons.sh"
  "$PKG/scripts/build_ico.py"
)
REGISTERED_DST=(
  "BRAND_ASSETS.md"
  "assets/brand/README.md"
  "assets/brand/scripts/build_icons.sh"
  "assets/brand/scripts/build_ico.py"
)
IMAGE_DST=(
  "assets/brand/excel-privacy-logo-master.png"
  "assets/brand/excel-privacy-logo-1024.png"
  "assets/brand/icons/mac/ExcelPrivacy.icns"
  "assets/brand/icons/win/ExcelPrivacy.ico"
)
BASE_FILES=(
  "BRAND_ASSETS.md"
  "assets/brand/README.md"
  "assets/brand/excel-privacy-logo-master.png"
  "assets/brand/excel-privacy-logo-1024.png"
  "assets/brand/icons/mac/ExcelPrivacy.icns"
  "assets/brand/icons/win/ExcelPrivacy.ico"
  "assets/brand/scripts/build_icons.sh"
  "assets/brand/scripts/build_ico.py"
)

sha256_of() {
  if command -v shasum >/dev/null 2>&1; then shasum -a 256 "$1" | awk '{print $1}'
  else sha256sum "$1" | awk '{print $1}'; fi
}

# R3.2: device + inode of a path, printed as "dev:ino". os.stat is used so the
# script stays portable across BSD `stat -f` and GNU `stat -c`; python3 is
# already a required dependency of this script (build_ico.py / autofill).
identity_of() {
  python3 - "$1" <<'PY'
import os, sys
s = os.stat(sys.argv[1])
print("%d:%d" % (s.st_dev, s.st_ino))
PY
}

STAGE=""
CREATED=()          # relative paths successfully published THIS run
CREATED_SHA=()      # their expected SHA-256 at publish time
CREATED_DEV=()      # R3.2: st_dev of the inode THIS run published
CREATED_INO=()      # R3.2: st_ino of the inode THIS run published
N_CREATED=0
TMPFILES=()         # run-unique temp files created THIS run (incl. partial writes)
N_TMP=0
TMP_SUFFIX="__r31_tmp__.$$"
ROLLED_BACK=0
README_TEMPLATE_SHA256=""
README_FINAL_SHA256=""

# R3.1 B (inherited) + R3.2 B (identity): rollback is deliberately
# conservative. It NEVER deletes a path that this run did not create, and it
# deletes a PUBLISHED file ONLY when SHA-256 + st_dev + st_ino ALL still match
# what this run froze at publish time.
rollback() {
  local rc=$?
  trap - ERR EXIT
  if [ "$rc" -ne 0 ] && [ "$ROLLED_BACK" -eq 0 ]; then
    ROLLED_BACK=1
    echo "BUILD FAILED (exit $rc) - rolling back this run only." >&2
    local i t rel want dst now want_id now_id
    # (a) temp files this run registered (partial writes included). The path
    #     is run-unique (PID-marked); only plain files, never symlinks/dirs.
    for ((i = 0; i < N_TMP; i++)); do
      t="${TMPFILES[$i]}"
      if [ -n "$t" ] && [ -f "$t" ] && [ ! -L "$t" ]; then
        rm -f -- "$t" || true
      fi
    done
    echo "  cleaned $N_TMP run-unique temp file(s) (incl. any partial write)." >&2
    # (b) final files this run published: ALL THREE identity parts must match.
    for ((i = 0; i < N_CREATED; i++)); do
      rel="${CREATED[$i]}"
      want="${CREATED_SHA[$i]}"
      want_id="${CREATED_DEV[$i]}:${CREATED_INO[$i]}"
      dst="$REPO/$rel"
      if [ -L "$dst" ]; then
        echo "  WARN: NOT removing '$rel' - path was replaced by a symlink." >&2
        continue
      fi
      if [ -f "$dst" ]; then
        now="$(sha256_of "$dst")" || now=""
        now_id="$(identity_of "$dst")" || now_id=""
        if [ "$now" != "$want" ]; then
          echo "  WARN: NOT removing '$rel' - content changed by another process/user." >&2
          echo "        (expected $want, found ${now:-<unreadable>}). Left intact." >&2
        elif [ "$now_id" != "$want_id" ]; then
          echo "  WARN: NOT removing '$rel' - inode identity changed by another process/user." >&2
          echo "        (published $want_id, found ${now_id:-<unreadable>}). Left intact." >&2
          echo "        SHA-256 still matches; inode check is what protects this file." >&2
        else
          rm -f -- "$dst" || true
        fi
      fi
    done
    echo "  rolled back this run's published files (SHA-256 + st_dev + st_ino verified)." >&2
    if [ -n "$STAGE" ] && [ -d "$STAGE" ]; then
      rm -rf -- "$STAGE" || true
    fi
    echo "Rollback done. No partial output remains; a clean re-run is safe." >&2
  fi
}
trap rollback ERR EXIT

# ---------------------------------------------------------------------------
# 0) PREFLIGHT: sources exist, Master identity, and NO existing target asset.
# ---------------------------------------------------------------------------
[ -f "$MASTER" ] || { echo "FATAL: master not found: $MASTER" >&2; exit 1; }
[ -d "$REPO" ] || { echo "FATAL: repo-root is not a directory: $REPO" >&2; exit 1; }
[ -d "$PKG" ] || { echo "FATAL: r3.1-package-root is not a directory: $PKG" >&2; exit 1; }
for s in "${REGISTERED_SRC[@]}"; do
  [ -s "$s" ] || { echo "FATAL: R3.1 package source missing/empty: $s" >&2; exit 1; }
done
MASTER_SHA="$(sha256_of "$MASTER")"
[ "$MASTER_SHA" = "$EXPECTED_SHA" ] || {
  echo "FATAL: Master SHA-256 mismatch (expected $EXPECTED_SHA, got $MASTER_SHA)" >&2; exit 1; }
python3 - "$MASTER" "$EXPECTED_DIM" <<'PY'
import struct, sys
path, dim = sys.argv[1], int(sys.argv[2])
b = open(path, "rb").read(33)
if b[:8] != b"\x89PNG\r\n\x1a\n":
    raise SystemExit("FATAL: master is not a PNG")
w, h, depth, ctype, comp, filt, interlace = struct.unpack(">IIBBBBB", b[16:29])
if (w, h) != (dim, dim):
    raise SystemExit("FATAL: master is %dx%d, expected %dx%d" % (w, h, dim, dim))
if depth != 8 or ctype != 2:
    raise SystemExit("FATAL: master must be 8-bit RGB (color type 2); got depth=%d type=%d" % (depth, ctype))
if interlace != 0 or comp != 0 or filt != 0:
    raise SystemExit("FATAL: master must be non-interlaced, deflate, filter method 0")
print("Master OK: %dx%d 8-bit RGB (opaque, no alpha), SHA-256 verified" % (w, h))
PY
for rel in "${BASE_FILES[@]}"; do
  if [ -e "$REPO/$rel" ] || [ -L "$REPO/$rel" ]; then
    echo "FATAL: target already exists, refusing to overwrite: $rel" >&2; exit 1
  fi
done
[ "$(cd "$(dirname "$MASTER")" && pwd)/$(basename "$MASTER")" = \
  "$(cd "$REPO" 2>/dev/null && pwd)/assets/brand/excel-privacy-logo-master.png" ] && {
  echo "FATAL: master and destination are the same file" >&2; exit 1; }

# ---------------------------------------------------------------------------
# 1) STAGE: everything is produced in a temp tree. Nothing reaches the repo
#    until COMMIT. The ICO is built by the STAGED copy of build_ico.py.
# ---------------------------------------------------------------------------
STAGE="$(mktemp -d)"
OUT="$STAGE/out"
WORK="$STAGE/work"
mkdir -p "$OUT/assets/brand/icons/mac" "$OUT/assets/brand/icons/win" \
         "$OUT/assets/brand/scripts" "$WORK/ico-pngs"
ICONSET="$WORK/ExcelPrivacy.iconset"
mkdir -p "$ICONSET"

echo "--- PREPARE: 4 registered files (draft -> stage, SHA-256 verified) ---"
for i in 0 1 2 3; do
  src="${REGISTERED_SRC[$i]}"
  rel="${REGISTERED_DST[$i]}"
  dst="$OUT/$rel"
  mkdir -p "$(dirname "$dst")"
  cp -p "$src" "$dst"
  a="$(sha256_of "$src")"
  b="$(sha256_of "$dst")"
  [ "$a" = "$b" ] || { echo "FATAL: draft/stage SHA-256 mismatch: $rel" >&2; exit 1; }
  if [ "$rel" = "assets/brand/README.md" ]; then
    README_TEMPLATE_SHA256="$b"
    echo "  README_TEMPLATE_SHA256 = $README_TEMPLATE_SHA256"
  fi
  printf '  %-48s  %s  <- %s\n' "$rel" "$b" "$(basename "$src")"
done
ICO_BUILDER="$OUT/assets/brand/scripts/build_ico.py"

# ---------------------------------------------------------------------------
# 2) GENERATE (4 image assets), all inside STAGE.
# ---------------------------------------------------------------------------
echo "--- GENERATE: 4 brand image assets ---"
cp -p "$MASTER" "$OUT/assets/brand/excel-privacy-logo-master.png"
[ "$(sha256_of "$OUT/assets/brand/excel-privacy-logo-master.png")" = "$EXPECTED_SHA" ] || {
  echo "FATAL: copied Master SHA-256 changed" >&2; exit 1; }
echo "  master copied verbatim (no matting / recolor / crop)."

sips -z 1024 1024 "$MASTER" --out "$OUT/assets/brand/excel-privacy-logo-1024.png" >/dev/null
python3 - "$OUT/assets/brand/excel-privacy-logo-1024.png" <<'PY'
import struct, sys
p = sys.argv[1]
b = open(p, "rb").read(29)
if b[:8] != b"\x89PNG\r\n\x1a\n":
    raise SystemExit("FATAL: 1024 output is not a PNG")
w, h, depth, ctype, comp, filt, interlace = struct.unpack(">IIBBBBB", b[16:29])
if (w, h) != (1024, 1024):
    raise SystemExit("FATAL: 1024 output is %dx%d, expected 1024x1024" % (w, h))
if depth != 8 or ctype not in (2, 6):
    raise SystemExit("FATAL: 1024 output must be 8-bit RGB/RGBA; got depth=%d type=%d" % (depth, ctype))
if interlace != 0 or comp != 0 or filt != 0:
    raise SystemExit("FATAL: 1024 output must be non-interlaced, deflate, filter method 0")
print("  1024 PNG OK: 1024x1024, 8-bit, color type %d, non-interlaced" % ctype)
PY

gen() { sips -z "$1" "$1" "$MASTER" --out "$2" >/dev/null; }
gen 16   "$ICONSET/icon_16x16.png"
gen 32   "$ICONSET/icon_16x16@2x.png"
gen 32   "$ICONSET/icon_32x32.png"
gen 64   "$ICONSET/icon_32x32@2x.png"
gen 128  "$ICONSET/icon_128x128.png"
gen 256  "$ICONSET/icon_128x128@2x.png"
gen 256  "$ICONSET/icon_256x256.png"
gen 512  "$ICONSET/icon_256x256@2x.png"
gen 512  "$ICONSET/icon_512x512.png"
gen 1024 "$ICONSET/icon_512x512@2x.png"
iconutil -c icns "$ICONSET" -o "$OUT/assets/brand/icons/mac/ExcelPrivacy.icns"
iconutil -c iconset "$OUT/assets/brand/icons/mac/ExcelPrivacy.icns" -o "$WORK/verify.iconset"
SLOTS="$(ls "$WORK/verify.iconset" | wc -l | tr -d ' ')"
[ "$SLOTS" = "10" ] || { echo "FATAL: ICNS decode returned $SLOTS slots, expected 10" >&2; exit 1; }
echo "  ICNS container OK: reverse-decoded 10 slots."
echo "  NOTE: reverse-decode proves CONTAINER/slot integrity ONLY."
echo "        It does NOT prove colour, corner-rounding or legibility."
echo "        Actual visual check (Finder / Dock / Get Info / Quick Look,"
echo "        About window) is a SEPARATE Gate B' human QA item with screenshots."
echo "        This script must NOT be reported as a visual PASS."

for s in 16 24 32 48 64 128 256; do
  sips -z "$s" "$s" "$MASTER" --out "$WORK/ico-pngs/icon_${s}.png" >/dev/null
done
python3 "$ICO_BUILDER" "$WORK/ico-pngs" "$OUT/assets/brand/icons/win/ExcelPrivacy.ico"
python3 "$ICO_BUILDER" --verify-only "$OUT/assets/brand/icons/win/ExcelPrivacy.ico" "$WORK/ico-pngs"
python3 "$ICO_BUILDER" --verify-only "$OUT/assets/brand/icons/win/ExcelPrivacy.ico"
echo "  ICO OK: 7 entries, full decode + verify-only passed (R3 decoder)."

# ---------------------------------------------------------------------------
# 2b) GENERATE-README (R3.1 C): substitute the 4 REAL image SHA-256 values
#     into the STAGED README template. No PENDING, no manual repo edit.
# ---------------------------------------------------------------------------
echo "--- GENERATE-README: autofill 4 real SHA-256 into staged README ---"
python3 - "$OUT" "${BASE_FILES[@]}" <<'PY'
import hashlib, pathlib, re, sys
out = pathlib.Path(sys.argv[1])
base_files = sys.argv[2:]
readme = out / "assets/brand/README.md"

def sha256(p):
    return hashlib.sha256(pathlib.Path(p).read_bytes()).hexdigest()

template_sha = sha256(readme)
mapping = {
    "MASTER":  out / "assets/brand/excel-privacy-logo-master.png",
    "PNG1024": out / "assets/brand/excel-privacy-logo-1024.png",
    "ICNS":    out / "assets/brand/icons/mac/ExcelPrivacy.icns",
    "ICO":     out / "assets/brand/icons/win/ExcelPrivacy.ico",
}
text = readme.read_text(encoding="utf-8")
values = {}
for key, p in mapping.items():
    if not p.is_file() or p.stat().st_size == 0:
        raise SystemExit("FATAL: asset missing/empty for autofill: %s" % p)
    h = sha256(p)
    if not re.fullmatch(r"[0-9a-f]{64}", h):
        raise SystemExit("FATAL: bad sha for %s" % key)
    token = "@@SHA256:%s@@" % key
    n = text.count(token)
    if n != 1:
        raise SystemExit("FATAL: placeholder %s appears %d times (want 1)" % (token, n))
    text = text.replace(token, h)
    values[key] = h
    print("  %-8s %s  <- %s" % (key, h, p.name))
if "@@SHA256:" in text:
    raise SystemExit("FATAL: unresolved placeholder remains in README")
readme.write_text(text, encoding="utf-8")
final_sha = sha256(readme)
if template_sha == final_sha:
    raise SystemExit("FATAL: README_TEMPLATE_SHA == README_FINAL_SHA (autofill did not change content)")
print("  README_TEMPLATE_SHA256 = %s" % template_sha)
print("  README_FINAL_SHA256    = %s" % final_sha)
# verify the 4 values are the REAL file hashes, not recomputed guesses
for key, p in mapping.items():
    if values[key] != sha256(p):
        raise SystemExit("FATAL: autofilled value != real file hash for %s" % key)
# confirm ALL 8 BASE_FILES have a final SHA-256
final = {}
for rel in base_files:
    f = out / rel
    if not f.is_file() or f.stat().st_size == 0:
        raise SystemExit("FATAL: BASE_FILES entry missing/empty: %s" % rel)
    h = sha256(f)
    if not re.fullmatch(r"[0-9a-f]{64}", h):
        raise SystemExit("FATAL: bad final sha for %s" % rel)
    final[rel] = h
    print("  FINAL %s  %s" % (h, rel))
if len(final) != 8:
    raise SystemExit("FATAL: expected 8 BASE_FILES, got %d" % len(final))
print("  All 8 BASE_FILES have a FINAL SHA-256. README autofill OK.")
PY
README_FINAL_SHA256="$(sha256_of "$OUT/assets/brand/README.md")"
[ "$README_TEMPLATE_SHA256" != "$README_FINAL_SHA256" ] || {
  echo "FATAL: README template/final SHA-256 identical" >&2; exit 1; }
echo "  Confirmed: TEMPLATE_SHA != FINAL_SHA (README was generated, not copied)."

# ---------------------------------------------------------------------------
# 3) VERIFY STAGE: all 8 BASE_FILES present/non-empty before committing.
# ---------------------------------------------------------------------------
echo "--- VERIFY STAGE: 8 explicit BASE_FILES (no glob) ---"
for rel in "${BASE_FILES[@]}"; do
  [ -s "$OUT/$rel" ] || { echo "FATAL: staged file missing/empty: $rel" >&2; exit 1; }
  case "$rel" in
    BRAND_ASSETS.md)                 note="registered: brand ownership notice (draft copy)" ;;
    assets/brand/README.md)          note="generated: template + 4 real SHA-256 autofill" ;;
    *master.png)                     note="generated: byte-identical Master copy" ;;
    *1024.png)                       note="generated: sips equal-ratio resize" ;;
    *.icns)                          note="generated: iconutil -c icns (10 slots)" ;;
    *.ico)                           note="generated: staged build_ico.py (R3)" ;;
    *build_icons.sh|*build_ico.py)   note="registered: script copy (SHA-256 linked to generator)" ;;
    *)                               note="UNKNOWN" ;;
  esac
  printf '  %-55s %s\n' "$rel" "$note"
done
[ "$(sha256_of "$MASTER")" = "$EXPECTED_SHA" ] || {
  echo "FATAL: original Master was modified" >&2; exit 1; }

# ---------------------------------------------------------------------------
# 4) COMMIT (R3.1 B): atomic, non-overwriting publication.
#    Each file: copy to a run-unique temp file in the SAME directory ->
#    SHA-256 check -> `ln` hard link (atomic, no clobber) -> unlink temp ->
#    final SHA-256 check. When `ln` finds an existing target it FAILS rather
#    than overwriting, so a file appearing mid-build is never clobbered.
# ---------------------------------------------------------------------------
echo "--- COMMIT (stage -> repo, atomic + no overwrite) ---"
for rel in "${BASE_FILES[@]}"; do
  if [ -e "$REPO/$rel" ] || [ -L "$REPO/$rel" ]; then
    echo "FATAL: target appeared during build: $rel" >&2; exit 1
  fi
done
mkdir -p "$REPO/assets/brand/icons/mac" "$REPO/assets/brand/icons/win" \
         "$REPO/assets/brand/scripts"

publish_one() {
  local rel="$1"
  local src="$OUT/$rel"
  local dst="$REPO/$rel"
  local tmp="$REPO/${rel}.${TMP_SUFFIX}"
  local want got now ident now_ident
  # never reuse/collide with a temp file from another run
  if [ -e "$tmp" ] || [ -L "$tmp" ]; then
    echo "FATAL: run-unique temp path already exists, refusing to touch it: $tmp" >&2
    return 1
  fi
  TMPFILES[$N_TMP]="$tmp"; N_TMP=$((N_TMP + 1))
  mkdir -p "$(dirname "$tmp")"
  # If cp writes partial bytes then fails, only "$tmp" is polluted.
  if ! cp -p "$src" "$tmp"; then
    echo "FATAL: temp copy failed (possibly partial write): $rel" >&2
    return 1
  fi
  want="$(sha256_of "$src")"
  got="$(sha256_of "$tmp")"
  if [ "$want" != "$got" ]; then
    echo "FATAL: temp SHA-256 mismatch: $rel" >&2
    return 1
  fi
  # R3.2: freeze the inode identity BEFORE publishing. `ln` is a hard link,
  # so the published path MUST carry exactly this (st_dev, st_ino).
  ident="$(identity_of "$tmp")" || {
    echo "FATAL: cannot stat temp file: $rel" >&2; return 1; }
  # atomic, NON-overwriting publish: ln fails if dst exists
  if [ -e "$dst" ] || [ -L "$dst" ]; then
    echo "FATAL: target appeared before atomic link: $rel" >&2
    return 1
  fi
  if ! ln "$tmp" "$dst" 2>/dev/null; then
    echo "FATAL: atomic link failed (target present or FS error): $rel" >&2
    return 1
  fi
  # record IMMEDIATELY so rollback covers this path even if later checks fail
  CREATED[$N_CREATED]="$rel"; CREATED_SHA[$N_CREATED]="$want"
  CREATED_DEV[$N_CREATED]="${ident%%:*}"
  CREATED_INO[$N_CREATED]="${ident##*:}"
  N_CREATED=$((N_CREATED + 1))
  rm -f -- "$tmp" || true
  now_ident="$(identity_of "$dst")" || now_ident=""
  if [ "$now_ident" != "$ident" ]; then
    echo "FATAL: published inode identity mismatch: $rel" >&2
    return 1
  fi
  now="$(sha256_of "$dst")"
  if [ "$want" != "$now" ]; then
    echo "FATAL: published SHA-256 mismatch: $rel" >&2
    return 1
  fi
  printf '  published  %-55s  %s  dev:ino=%s\n' "$rel" "$want" "$ident"
}

for rel in "${BASE_FILES[@]}"; do
  publish_one "$rel"
done
echo "  COMMIT OK: $N_CREATED files published atomically (no overwrite)."

echo "--- SHA-256 of the 4 ACTUAL brand IMAGE assets ---"
for rel in "${IMAGE_DST[@]}"; do
  printf '  %s  %s\n' "$(sha256_of "$REPO/$rel")" "$rel"
done
echo "--- SHA-256 of the 4 registered text/script assets (incl. FINAL README) ---"
for rel in "${REGISTERED_DST[@]}"; do
  printf '  %s  %s\n' "$(sha256_of "$REPO/$rel")" "$rel"
done
echo "--- README template vs final (R3.1 C) ---"
echo "  README_TEMPLATE_SHA256 = $README_TEMPLATE_SHA256"
echo "  README_FINAL_SHA256    = $README_FINAL_SHA256"
echo "Count: 4 generated image assets + 1 generated README ledger + 3 registered text/script = 8 BASE_FILES paths."
echo "README §4 was filled with REAL values by this script; NO manual repo edit required."
echo "NEXT (not done here): register the 8 explicit paths in tools/stage_release.py BASE_FILES;"
echo "      then Windows PE resource readback + dual-platform visual QA remain."
[ "$(sha256_of "$MASTER")" = "$EXPECTED_SHA" ] || {
  echo "FATAL: original Master was modified" >&2; exit 1; }
echo "Original Master intact (SHA-256 unchanged)."
# Success path: discard the temp stage explicitly; EXIT trap sees rc=0.
rm -rf -- "$STAGE"
STAGE=""
exit 0
