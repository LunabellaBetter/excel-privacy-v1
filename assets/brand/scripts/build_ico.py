#!/usr/bin/env python3
"""DRAFT ONLY - NOT EXECUTED AGAINST THE REAL LOGO. Requires Una authorization.

Excel Privacy Windows ICO builder - GPT Review R3 revision.

INHERITED from R2 (GPT independently reproduced PASS; NOT redeveloped):
  A. DIB biSizeImage = 0; verifier asserts 0.
  B. 256 PNG entry: full-pixel alpha == 255, enforced also in --verify-only.
  C. Exact size set {16,24,32,48,64,128,256}: no duplicate, missing or extra.
  D. PNG integrity: IHDR/IDAT/IEND + CRC + dimensions + full filter 0-4 decode.
  E. DIB 32bpp BGRA, biHeight=2h, DWORD-aligned all-zero AND mask, opaque.

R3 P1-B point fixes to the PNG DECODER only:
  1) tRNS is REJECTED (project requires fully opaque; never interpret a
     transparent-colour chunk).
  2) Unknown CRITICAL chunk (uppercase first letter) -> REJECT.
  3) Unknown ANCILLARY chunk (lowercase first letter) -> allowed/ignored per
     the PNG spec; CRC still verified; recorded in meta.
  4) PLTE: duplicate rejected; must appear after IHDR and before the first
     IDAT; length divisible by 3 and <= 3*2^bitdepth entries.
  5) IDAT uses a STRICT zlib decompressobj: checks .eof, .unused_data,
     .unconsumed_tail; rejects truncated streams, trailing garbage and a
     second/extra zlib stream.
  6) Pre-decompression size guard: dimension + expected scanline byte cap,
     and a max_length bound on the decompressor (zip-bomb protection).

Still zero third-party dependencies: standard-library zlib + struct only.
No matting, no background removal, no recolor, no redesign, no vectorization.

Usage:
  build_ico.py <png-dir> <output.ico>
  build_ico.py --verify-only <existing.ico> [png-dir]
"""
import struct
import sys
import zlib
from pathlib import Path

PNG_MAGIC = b"\x89PNG\r\n\x1a\n"
DIB_SIZES = (16, 24, 32, 48, 64, 128)
PNG_SIZES = (256,)
SIZES = DIB_SIZES + PNG_SIZES
EXPECTED_SIZES = frozenset(SIZES)
BITMAPINFOHEADER = "<IiiHHIIiiII"  # 40 bytes

# R3-B: decoder limits and chunk classification.
MAX_DIM = 8192
MAX_SCANLINE_BYTES = 256 * 1024 * 1024  # 256 MiB pre-decompression guard
CRITICAL_KNOWN = {b"IHDR", b"PLTE", b"IDAT", b"IEND"}
# Ancillary chunks with ordering constraints we understand; everything else
# ancillary is ignored per spec (but CRC + placement are still checked).
ANCILLARY_BEFORE_IDAT = {b"cHRM", b"gAMA", b"iCCP", b"sBIT", b"sRGB",
                         b"bKGD", b"hIST", b"pHYs", b"sPLT", b"tRNS"}


def check_size_set():
    """R2-C: the declared set itself must be exact and duplicate-free."""
    if len(SIZES) != len(EXPECTED_SIZES):
        raise ValueError("size set contains duplicates")
    if sorted(SIZES) != [16, 24, 32, 48, 64, 128, 256]:
        raise ValueError("size set must be exactly 16/24/32/48/64/128/256")


def _paeth(a, b, c):
    p = a + b - c
    pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
    return a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)


def _chunk_kind(tag):
    """Return 'critical' / 'ancillary', validating the 4-letter chunk code."""
    if len(tag) != 4 or not all(65 <= c <= 90 or 97 <= c <= 122 for c in tag):
        raise ValueError("invalid PNG chunk type %r" % (tag,))
    # Reserved bit (3rd letter) must be uppercase per the PNG spec.
    if not (65 <= tag[2] <= 90):
        raise ValueError("PNG chunk %r has invalid reserved bit" % (tag,))
    return "critical" if 65 <= tag[0] <= 90 else "ancillary"


def _png_chunks(data):
    """Yield (tag, payload, offset, length) with CRC / bounds / IEND checks."""
    if not data.startswith(PNG_MAGIC):
        raise ValueError("not a PNG (bad magic)")
    pos = 8
    ended = False
    while pos < len(data):
        if pos + 8 > len(data):
            raise ValueError("truncated PNG chunk header")
        length = struct.unpack_from(">I", data, pos)[0]
        tag = data[pos + 4:pos + 8]
        if pos + 12 + length > len(data):
            raise ValueError("truncated PNG chunk payload")
        payload = data[pos + 8:pos + 8 + length]
        crc = struct.unpack_from(">I", data, pos + 8 + length)[0]
        if zlib.crc32(tag + payload) & 0xFFFFFFFF != crc:
            raise ValueError("PNG CRC mismatch in %r" % (tag,))
        _chunk_kind(tag)  # R3: validate chunk code even for ignored chunks
        yield tag, payload, pos, length
        pos += 12 + length
        if tag == b"IEND":
            ended = True
            break
    if not ended:
        raise ValueError("PNG missing IEND")
    if pos != len(data):
        raise ValueError("PNG has trailing bytes after IEND")


def decode_png(data, with_meta=False):
    """Return (width, height, pixels); optionally (..., meta).

    pixels are top-down (r, g, b, a). RGB sources are promoted to alpha=255;
    the ivory background is preserved. Only 8-bit RGB/RGBA, non-interlaced.
    """
    chunks = list(_png_chunks(data))
    tags = [c[0] for c in chunks]
    if not tags or tags[0] != b"IHDR" or tags.count(b"IHDR") != 1:
        raise ValueError("PNG IHDR must be first and appear exactly once")
    if len(chunks[0][1]) != 13:
        raise ValueError("PNG IHDR payload must be 13 bytes")
    if tags[-1] != b"IEND" or tags.count(b"IEND") != 1:
        raise ValueError("PNG IEND must be last and appear exactly once")
    if chunks[-1][3] != 0:
        raise ValueError("PNG IEND payload must be empty")
    idat_idx = [i for i, t in enumerate(tags) if t == b"IDAT"]
    if not idat_idx:
        raise ValueError("PNG has no IDAT")
    if idat_idx != list(range(idat_idx[0], idat_idx[0] + len(idat_idx))):
        raise ValueError("PNG IDAT chunks must be consecutive")

    # --- R3-B 1/2/3/4: chunk-level policy -------------------------------
    if b"tRNS" in tags:
        raise ValueError("tRNS rejected: this project requires fully opaque PNGs")
    for tag, _, _, _ in chunks:
        kind = _chunk_kind(tag)
        if kind == "critical" and tag not in CRITICAL_KNOWN:
            raise ValueError("unknown CRITICAL chunk %r rejected" % (tag,))
    plte_idx = [i for i, t in enumerate(tags) if t == b"PLTE"]
    if len(plte_idx) > 1:
        raise ValueError("duplicate PLTE chunk")
    if plte_idx:
        i = plte_idx[0]
        if i <= 0 or i >= idat_idx[0]:
            raise ValueError("PLTE must appear after IHDR and before the first IDAT")
        n = len(chunks[i][1])
        if n == 0 or n % 3:
            raise ValueError("PLTE length must be a non-zero multiple of 3")
    ignored_ancillary = []
    for i, (tag, _, _, _) in enumerate(chunks):
        if _chunk_kind(tag) == "ancillary":
            if tag in ANCILLARY_BEFORE_IDAT and i > idat_idx[0]:
                raise ValueError("ancillary %r must precede IDAT" % (tag,))
            ignored_ancillary.append(tag.decode("latin-1"))
    # ---------------------------------------------------------------------

    width, height, depth, color_type, comp, filt, interlace = struct.unpack(
        ">IIBBBBB", chunks[0][1]
    )
    if depth != 8 or color_type not in (2, 6) or interlace != 0:
        raise ValueError("unsupported PNG: need 8-bit RGB/RGBA, non-interlaced")
    if comp != 0 or filt != 0:
        raise ValueError("unsupported PNG compression/filter method")
    if width == 0 or height == 0:
        raise ValueError("PNG has a zero dimension")
    if plte_idx and len(chunks[plte_idx[0]][1]) // 3 > (1 << depth):
        raise ValueError("PLTE has more entries than the bit depth allows")

    channels = 3 if color_type == 2 else 4
    stride = width * channels
    expected = height * (stride + 1)
    # R3-B 6: pre-decompression size guard (no decompression before this).
    if width > MAX_DIM or height > MAX_DIM or expected > MAX_SCANLINE_BYTES:
        raise ValueError("PNG dimensions exceed the decoder size guard")

    idat = b"".join(c[1] for c in chunks if c[0] == b"IDAT")
    # R3-B 5: strict zlib stream. max_length bounds output; then check eof,
    # unused_data (trailing garbage / a 2nd zlib stream) and unconsumed_tail.
    dec = zlib.decompressobj()
    raw = dec.decompress(idat, expected + 1)
    if len(raw) > expected:
        raise ValueError("PNG inflated data exceeds expected scanline size")
    if not dec.eof:
        raise ValueError("PNG zlib stream is truncated / incomplete")
    if dec.unconsumed_tail:
        raise ValueError("PNG zlib stream left unconsumed input")
    if dec.unused_data:
        raise ValueError("PNG IDAT has trailing garbage or an extra zlib stream")
    if len(raw) != expected:
        raise ValueError("PNG scanline length mismatch")

    decoded = bytearray()
    prev = bytearray(stride)
    for y in range(height):
        start = y * (stride + 1)
        f = raw[start]
        line = bytearray(raw[start + 1:start + 1 + stride])
        if f == 0:
            pass
        elif f == 1:
            for i in range(channels, stride):
                line[i] = (line[i] + line[i - channels]) & 0xFF
        elif f == 2:
            for i in range(stride):
                line[i] = (line[i] + prev[i]) & 0xFF
        elif f == 3:
            for i in range(stride):
                left = line[i - channels] if i >= channels else 0
                line[i] = (line[i] + ((left + prev[i]) >> 1)) & 0xFF
        elif f == 4:
            for i in range(stride):
                a = line[i - channels] if i >= channels else 0
                b = prev[i]
                c = prev[i - channels] if i >= channels else 0
                line[i] = (line[i] + _paeth(a, b, c)) & 0xFF
        else:
            raise ValueError("bad PNG filter %d" % f)
        decoded += line
        prev = line

    pixels = []
    for i in range(0, len(decoded), channels):
        r, g, b = decoded[i], decoded[i + 1], decoded[i + 2]
        a = decoded[i + 3] if channels == 4 else 255
        pixels.append((r, g, b, a))
    if len(pixels) != width * height:
        raise ValueError("PNG decoded pixel count mismatch")

    meta = {
        "color_type": color_type,
        "bit_depth": depth,
        "channels": channels,
        "bpp_effective": 8 * channels,  # RGB -> 24, RGBA -> 32
        "interlace": interlace,
        "idat_chunks": len(idat_idx),
        "chunk_tags": [t.decode("latin-1") for t in tags],
        "ignored_ancillary": ignored_ancillary,
        "plte_present": bool(plte_idx),
        "zlib_eof": dec.eof,
        "zlib_unused_data": len(dec.unused_data),
        "zlib_unconsumed_tail": len(dec.unconsumed_tail),
    }
    return (width, height, pixels, meta) if with_meta else (width, height, pixels)


def encode_dib(size, pixels):
    """BITMAPINFOHEADER + bottom-up BGRA XOR + all-zero AND mask."""
    xor = bytearray()
    for y in range(size - 1, -1, -1):  # ICO DIB is bottom-up
        for x in range(size):
            r, g, b, a = pixels[y * size + x]
            if a != 255:
                raise ValueError("alpha must be fully opaque (ivory background kept)")
            xor += bytes((b, g, r, 255))
    mask_row = ((size + 31) // 32) * 4  # 1bpp rows padded to DWORD
    and_mask = bytes(mask_row * size)    # all zero -> no transparency mask
    header = struct.pack(
        BITMAPINFOHEADER,
        40, size, 2 * size, 1, 32, 0, 0, 0, 0, 0, 0,
    )
    return bytes(header) + bytes(xor) + and_mask


def build(png_dir, out):
    check_size_set()
    entries = []
    payload = bytearray()
    offset = 6 + 16 * len(SIZES)
    for size in SIZES:
        src = png_dir / ("icon_%d.png" % size)
        data = src.read_bytes()
        w, h, pixels, meta = decode_png(data, with_meta=True)
        if (w, h) != (size, size):
            raise ValueError("%s: expected %dx%d, got %dx%d" % (src, size, size, w, h))
        if any(a != 255 for _, _, _, a in pixels):
            raise ValueError("%s: non-opaque alpha; do not remove the background" % src)
        if size in DIB_SIZES:
            blob = encode_dib(size, pixels)
            bits = 32
        else:
            blob = data
            bits = meta["bpp_effective"]
        dim = 0 if size == 256 else size  # 0 encodes 256
        entries.append(struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, bits, len(blob), offset))
        payload += blob
        offset += len(blob)
    header = struct.pack("<HHH", 0, 1, len(SIZES))
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(header + b"".join(entries) + bytes(payload))


def verify(path, sources=None):
    check_size_set()
    raw = path.read_bytes()
    reserved, kind, count = struct.unpack_from("<HHH", raw, 0)
    if reserved != 0 or kind != 1 or count != len(SIZES):
        raise ValueError("bad ICONDIR (type/count)")
    dir_end = 6 + 16 * count
    spans = []
    seen = []
    for i in range(count):
        w, h, colors, res, planes, bits, length, offset = struct.unpack_from(
            "<BBBBHHII", raw, 6 + 16 * i
        )
        size = 256 if (w == 0 and h == 0) else w
        if size == 256:
            if (w, h) != (0, 0):
                raise ValueError("entry %d: 256 entry must encode 0/0" % i)
        elif (w, h) != (size, size):
            raise ValueError("entry %d: directory width/height mismatch" % i)
        if size not in EXPECTED_SIZES:
            raise ValueError("entry %d: unexpected size %d" % (i, size))
        if size in seen:
            raise ValueError("entry %d: duplicate size %d" % (i, size))
        seen.append(size)
        if colors != 0 or res != 0 or planes != 1:
            raise ValueError("entry %d: bad directory metadata" % i)
        if offset < dir_end or offset + length > len(raw) or length == 0:
            raise ValueError("entry %d: offset/length out of bounds" % i)
        spans.append((offset, offset + length))
        blob = raw[offset:offset + length]

        if size in DIB_SIZES:
            if bits != 32:
                raise ValueError("entry %d: DIB directory bit depth must be 32" % i)
            (bi_size, bi_w, bi_h, bi_planes, bi_bits, bi_comp,
             bi_size_image, _, _, _, _) = struct.unpack_from(BITMAPINFOHEADER, blob, 0)
            if bi_size != 40 or bi_w != size or bi_h != 2 * size:
                raise ValueError("entry %d: bad BITMAPINFOHEADER geometry" % i)
            if bi_planes != 1 or bi_bits != 32 or bi_comp != 0 or bi_bits != bits:
                raise ValueError("entry %d: bit depth / metadata mismatch" % i)
            if bi_size_image != 0:
                raise ValueError("entry %d: biSizeImage must be 0 for BI_RGB" % i)
            xor_len = size * size * 4
            and_len = ((size + 31) // 32) * 4 * size
            if length != 40 + xor_len + and_len:
                raise ValueError("entry %d: image length mismatch" % i)
            xor = blob[40:40 + xor_len]
            if any(blob[40 + xor_len:]):
                raise ValueError("entry %d: AND mask must be all zero (opaque)" % i)
            decoded = []
            for y in range(size):
                row = xor[(size - 1 - y) * size * 4:(size - y) * size * 4]
                for x in range(size):
                    b, g, r, a = row[x * 4:x * 4 + 4]
                    if a != 255:
                        raise ValueError("entry %d: non-opaque DIB pixel" % i)
                    decoded.append((r, g, b, a))
            encoding = "DIB/BMP"
        else:
            pw, ph, decoded, meta = decode_png(blob, with_meta=True)
            if (pw, ph) != (256, 256):
                raise ValueError("entry %d: PNG entry is not 256x256" % i)
            if bits != meta["bpp_effective"]:
                raise ValueError("entry %d: directory bit depth != PNG pixel depth" % i)
            # R2-B: full alpha check, including --verify-only (sources=None).
            if any(a != 255 for _, _, _, a in decoded):
                raise ValueError("entry %d: 256 PNG has transparent pixels" % i)
            encoding = "PNG"

        if sources is not None:
            sw, sh, expected = decode_png((sources / ("icon_%d.png" % size)).read_bytes())
            if (sw, sh) != (size, size) or decoded != expected:
                raise ValueError("entry %d: decoded pixels differ from source PNG" % i)

        print(
            "entry %d: %dx%d %-8s offset=%d length=%d bpp=%d "
            "decoded_pixels=%d alpha=opaque%s" % (
                i, size, size, encoding, offset, length, bits, len(decoded),
                " biSizeImage=0" if size in DIB_SIZES else "",
            )
        )
    # R2-C: exact set, no missing / extra / duplicate sizes.
    if sorted(seen) != sorted(SIZES):
        raise ValueError("icon sizes are missing or extra: %r" % (seen,))
    spans.sort()
    for (_, e1), (s2, _) in zip(spans, spans[1:]):
        if e1 > s2:
            raise ValueError("overlapping icon entries")
    print("ICO OK: %d entries %r, %d bytes, no overlap, full decode passed"
          % (count, sorted(seen), len(raw)))


if __name__ == "__main__":
    args = sys.argv[1:]
    if len(args) == 2 and args[0] == "--verify-only":
        verify(Path(args[1]))
    elif len(args) == 3 and args[0] == "--verify-only":
        verify(Path(args[1]), Path(args[2]))
    elif len(args) == 2:
        png_dir, out = Path(args[0]), Path(args[1])
        build(png_dir, out)
        verify(out, png_dir)
    else:
        raise SystemExit(
            "usage: build_ico.py <png-dir> <output.ico> | "
            "build_ico.py --verify-only <existing.ico> [png-dir]"
        )
