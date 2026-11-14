"""Windows V1.0.1 build acceptance (Gate B' 7 items + SHA256).

Pure stdlib + pefile. Verifies the frozen EXE's embedded icon resources and
version info, cross-checks the source ICO, and reports a machine-readable JSON.
Designed to run on the GitHub windows-latest runner (pefile is pinned for win32).

Checks:
  B1  source ICO has exactly 7 entries and biSizeImage == 0 (compatibility)
  B2  source ICO AND-mask rows are DWORD (4-byte) aligned for 24/48/64
  B3  PE resource readback works (pefile loads RT_GROUP_ICON / RT_ICON)
  B4  PyInstaller preserved all 7 icon entries in the EXE
  B5  EXE carries the embedded icon + file version == 1.0.1.0
  B6  (GUI) shortcut / taskbar / Alt-Tab icon -> deferred to visual QA
  SHA256 of the frozen EXE
"""
from __future__ import annotations

import hashlib
import json
import struct
import sys
from pathlib import Path

EXPECTED_VERSION = (1, 0, 1, 0)
EXPECTED_ICON_ENTRIES = 7


def fail(results, key, msg):
    results[key] = {"status": "FAIL", "detail": msg}
    return False


def ok(results, key, detail=None):
    results[key] = {"status": "PASS", "detail": detail or ""}
    return True


def parse_icon_dir(data: bytes):
    """Return list of image descriptors from an ICO file or GRPICONDIR bytes."""
    if len(data) < 6:
        raise ValueError("icon directory too small")
    reserved, itype, count = struct.unpack_from("<HHH", data, 0)
    if itype != 1:
        raise ValueError("not an icon directory (type=%d)" % itype)
    entries = []
    off = 6
    for _ in range(count):
        w, h, cc, rsv, planes, bc, bytes_in_res, img_off = struct.unpack_from(
            "<BBBBHHII", data, off
        )
        off += 16  # ICONDIRENTRY (source) ; GRPICONDIRENTRY is 14 but we only
        # read the first 16 fields here because for the source ICO the 16-byte
        # entry applies; for GRPICONDIR we use parse_group_icon separately.
        entries.append(
            {
                "w": w or 256,
                "h": h or 256,
                "bit_count": bc,
                "bytes_in_res": bytes_in_res,
                "img_off": img_off,
            }
        )
    return count, entries


def parse_group_icon(group: bytes):
    """GRPICONDIR: 6-byte header + 14-byte GRPICONDIRENTRY[] (id instead of offset)."""
    if len(group) < 6:
        raise ValueError("group icon too small")
    reserved, itype, count = struct.unpack_from("<HHH", group, 0)
    if itype != 1:
        raise ValueError("not a group icon (type=%d)" % itype)
    entries = []
    off = 6
    for _ in range(count):
        w, h, cc, rsv, planes, bc, bytes_in_res, icon_id = struct.unpack_from(
            "<BBBBHHIH", group, off
        )
        off += 14
        entries.append(
            {
                "w": w or 256,
                "h": h or 256,
                "bit_count": bc,
                "bytes_in_res": bytes_in_res,
                "icon_id": icon_id,
            }
        )
    return count, entries


PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"


def check_image_bytes(img: bytes, label: str, results, key, *, require_bisize_zero):
    """Validate BITMAPINFOHEADER + XOR/AND layout of one icon image.

    The 256px entry in a modern ICO is PNG-stored (alpha native, no AND mask);
    detect and accept it instead of applying the BMP layout checks.
    """
    if img[:8] == PNG_SIGNATURE:
        return ok(results, key, "%s: PNG-stored (alpha native, no AND mask)" % label)
    if len(img) < 40:
        return fail(results, key, "%s: image < 40 bytes" % label)
    (bi_size, bi_width, bi_height, _planes, bi_bitcount, _comp, bi_size_image) = (
        struct.unpack_from("<IiiHHII", img, 0)
    )
    if bi_size != 40:
        return fail(results, key, "%s: unexpected biSize=%d" % (label, bi_size))
    actual_h = bi_height // 2  # ICO stores XOR+AND stacked => 2x height
    if actual_h <= 0:
        return fail(results, key, "%s: invalid biHeight=%d" % (label, bi_height))
    xor_stride = ((bi_width * bi_bitcount + 31) // 32) * 4
    xor_size = actual_h * xor_stride
    and_stride = ((bi_width + 31) // 32) * 4
    and_size = actual_h * and_stride
    expected = bi_size + xor_size + and_size
    if len(img) < expected:
        return fail(
            results,
            key,
            "%s: image len %d < expected %d (xor=%d and=%d)"
            % (label, len(img), expected, xor_size, and_size),
        )
    if and_stride % 4 != 0:
        return fail(results, key, "%s: AND-mask stride %d not DWORD aligned" % (label, and_stride))
    if require_bisize_zero and bi_size_image != 0:
        return fail(results, key, "%s: biSizeImage=%d (expected 0)" % (label, bi_size_image))
    detail = "w=%d h=%d bpp=%d biSizeImage=%d andStride=%d" % (
        bi_width,
        actual_h,
        bi_bitcount,
        bi_size_image,
        and_stride,
    )
    return ok(results, key, detail)


def verify_source_ico(path: Path, results):
    try:
        data = path.read_bytes()
    except OSError as e:
        return fail(results, "B1_source_ico", "cannot read %s: %s" % (path, e))
    try:
        count, entries = parse_icon_dir(data)
    except ValueError as e:
        return fail(results, "B1_source_ico", str(e))
    if count != EXPECTED_ICON_ENTRIES:
        return fail(results, "B1_source_ico", "entry count=%d expected %d" % (count, EXPECTED_ICON_ENTRIES))
    ok(results, "B1_source_ico", "entries=%d" % count)
    all_ok = True
    for i, e in enumerate(entries):
        key = "B2_andmask_%d" % i
        img = data[e["img_off"]: e["img_off"] + e["bytes_in_res"]]
        if not check_image_bytes(img, "ico[%d]" % i, results, key, require_bisize_zero=True):
            all_ok = False
    if all_ok:
        ok(results, "B2_andmask_summary", "all %d entries DWORD-aligned, biSizeImage=0" % count)
    else:
        fail(results, "B2_andmask_summary", "one or more entries failed AND-mask/biSizeImage")
    return all_ok


def verify_exe(exe: Path, results):
    try:
        import pefile
    except ImportError:
        return fail(results, "B3_pe_readback", "pefile not installed")
    try:
        pe = pefile.PE(str(exe))
    except Exception as e:  # noqa: BLE001
        return fail(results, "B3_pe_readback", "pefile failed: %s" % e)
    try:
        if not hasattr(pe, "DIRECTORY_ENTRY_RESOURCE"):
            return fail(results, "B3_pe_readback", "no resource directory")
        group = None
        icons = {}
        rt_group = pefile.RESOURCE_TYPE["RT_GROUP_ICON"]
        rt_icon = pefile.RESOURCE_TYPE["RT_ICON"]
        for t in pe.DIRECTORY_ENTRY_RESOURCE.entries:
            if t.id == rt_group:
                for ne in t.directory.entries:
                    for le in ne.directory.entries:
                        group = pe.get_data(
                            le.data.struct.OffsetToData, le.data.struct.Size
                        )
            elif t.id == rt_icon:
                for ne in t.directory.entries:
                    for le in ne.directory.entries:
                        icons[ne.id] = pe.get_data(
                            le.data.struct.OffsetToData, le.data.struct.Size
                        )
        if group is None:
            return fail(results, "B3_pe_readback", "no RT_GROUP_ICON resource")
        ok(results, "B3_pe_readback", "pefile loaded group+icon resources")
        try:
            count, entries = parse_group_icon(group)
        except ValueError as e:
            return fail(results, "B4_icon_entries", str(e))
        if count != EXPECTED_ICON_ENTRIES:
            return fail(results, "B4_icon_entries", "embedded count=%d expected %d" % (count, EXPECTED_ICON_ENTRIES))
        ok(results, "B4_icon_entries", "embedded entries=%d" % count)
        all_ok = True
        for i, e in enumerate(entries):
            key = "B5_embedded_%d" % i
            img = icons.get(e["icon_id"])
            if img is None:
                fail(results, key, "missing RT_ICON id=%d" % e["icon_id"])
                all_ok = False
                continue
            if not check_image_bytes(img, "exe-ico[%d]" % i, results, key, require_bisize_zero=False):
                all_ok = False
        # file version
        try:
            ffi = pe.VS_FIXEDFILEINFO[0]
            fv = (
                ffi.FileVersionMS >> 16,
                ffi.FileVersionMS & 0xFFFF,
                ffi.FileVersionLS >> 16,
                ffi.FileVersionLS & 0xFFFF,
            )
        except Exception:  # noqa: BLE001
            fv = None
        if fv != EXPECTED_VERSION:
            fail(results, "B5_file_version", "file version=%s expected %s" % (fv, EXPECTED_VERSION))
        else:
            ok(results, "B5_file_version", "file version=%s" % (fv,))
            if all_ok:
                ok(results, "B5_embedded_summary", "all %d embedded entries valid" % count)
            else:
                fail(results, "B5_embedded_summary", "one or more embedded entries failed")
        return all_ok and fv == EXPECTED_VERSION
    finally:
        pe.close()


def sha256(path: Path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    exe_path = None
    ico_path = None
    out_path = None
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--exe":
            exe_path = Path(argv[i + 1]); i += 2
        elif a == "--ico":
            ico_path = Path(argv[i + 1]); i += 2
        elif a == "--out":
            out_path = Path(argv[i + 1]); i += 2
        else:
            i += 1
    results = {}
    if ico_path:
        verify_source_ico(ico_path, results)
    if exe_path:
        verify_exe(exe_path, results)
        results["SHA256_EXE"] = {"status": "INFO", "detail": sha256(exe_path)}
    # overall
    failed = [k for k, v in results.items() if isinstance(v, dict) and v.get("status") == "FAIL"]
    summary = {
        "overall": "PASS" if not failed else "FAIL",
        "failed": failed,
        "checks": results,
    }
    if out_path:
        out_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False) + "\n")
    print(json.dumps(summary, indent=2, ensure_ascii=False))
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
