"""Rebuild an MS-OVBA project from source; never retain stale compiled code.

Preserves original project references, document modules and stream names. The
small CFB writer implements both FAT and MiniFAT (MS-CFB version 3).
"""
from __future__ import annotations

import io
import math
import re
import struct
from pathlib import Path

import olefile
from oletools.olevba import VBA_Project, decompress_stream

FREE, END, FAT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD


def compress_vba(data: bytes) -> bytes:
    result = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        payload = bytearray()
        pos = 0
        while pos < len(chunk):
            flags_at = len(payload)
            payload.append(0)
            for bit in range(8):
                if pos == len(chunk):
                    break
                bits = max(4, (pos - 1).bit_length())
                maximum = min((0xFFFF >> bits) + 3, len(chunk) - pos)
                best_length = 0
                best_offset = 0
                if maximum >= 3:
                    # Check matching prefixes, including overlapping copies.
                    previous = chunk.rfind(chunk[pos:pos + 3], max(0, pos - (1 << bits)), pos + 2)
                    while 0 <= previous < pos:
                        length = 3
                        while length < maximum and chunk[previous + length] == chunk[pos + length]:
                            length += 1
                        if length > best_length:
                            best_length, best_offset = length, pos - previous
                            if length == maximum:
                                break
                        previous = chunk.rfind(chunk[pos:pos + 3], max(0, pos - (1 << bits)), previous + 2)
                if best_length >= 3:
                    payload[flags_at] |= 1 << bit
                    token = ((best_offset - 1) << (16 - bits)) | (best_length - 3)
                    payload.extend(struct.pack("<H", token))
                    pos += best_length
                else:
                    payload.append(chunk[pos])
                    pos += 1
        if len(payload) >= 4096 and len(chunk) == 4096:
            result.extend(struct.pack("<H", 0x3FFF))
            result.extend(chunk)
        else:
            if len(payload) > 4096:
                raise ValueError("Incompressible final VBA chunk; pad the source with whitespace")
            result.extend(struct.pack("<H", 0xB000 | (len(payload) - 1)))
            result.extend(payload)
    assert decompress_stream(bytes(result)) == data
    return bytes(result)


def cfb_write(streams: dict[str, bytes]) -> bytes:
    names = ["Root Entry", "VBA"] + sorted(streams)
    entries = [{"name": n.rsplit("/", 1)[-1], "type": 5 if i == 0 else 1 if i == 1 else 2,
                "left": FREE, "right": FREE, "child": FREE, "color": 1,
                "start": END, "size": 0} for i, n in enumerate(names)]

    def tree(indices: list[int]) -> int:
        # CLRS red-black insertion; CFB compares length then uppercase UTF-16.
        parent = {FREE: FREE}
        root = FREE

        def color(i):
            return 1 if i == FREE else entries[i]["color"]

        def key(i):
            n = entries[i]["name"]
            return len(n.encode("utf-16le")), n.upper()

        def rotate(x, left):
            nonlocal root
            side, other = ("right", "left") if left else ("left", "right")
            y = entries[x][side]
            entries[x][side] = entries[y][other]
            if entries[y][other] != FREE:
                parent[entries[y][other]] = x
            parent[y] = parent[x]
            if parent[x] == FREE:
                root = y
            elif x == entries[parent[x]]["left"]:
                entries[parent[x]]["left"] = y
            else:
                entries[parent[x]]["right"] = y
            entries[y][other] = x
            parent[x] = y

        for z in indices:
            p, x = FREE, root
            while x != FREE:
                p = x
                x = entries[x]["left" if key(z) < key(x) else "right"]
            parent[z] = p
            if p == FREE:
                root = z
            else:
                entries[p]["left" if key(z) < key(p) else "right"] = z
            entries[z]["color"] = 0
            while color(parent[z]) == 0:
                p = parent[z]
                g = parent[p]
                side = "left" if p == entries[g]["left"] else "right"
                other = "right" if side == "left" else "left"
                uncle = entries[g][other]
                if color(uncle) == 0:
                    entries[p]["color"] = entries[uncle]["color"] = 1
                    entries[g]["color"] = 0
                    z = g
                else:
                    if z == entries[p][other]:
                        z = p
                        rotate(z, side == "left")
                    p = parent[z]
                    g = parent[p]
                    entries[p]["color"] = 1
                    entries[g]["color"] = 0
                    rotate(g, side != "left")
            entries[root]["color"] = 1
        return root

    entries[0]["child"] = tree([1] + [i for i, n in enumerate(names) if i > 1 and "/" not in n])
    entries[1]["child"] = tree([i for i, n in enumerate(names) if i > 1 and n.startswith("VBA/")])
    sectors: list[bytes] = []
    fat: list[int] = []
    mini_data = bytearray()
    mini_fat: list[int] = []

    def allocate(data: bytes) -> int:
        if not data:
            return END
        first = len(sectors)
        for pos in range(0, len(data), 512):
            sectors.append(data[pos:pos + 512].ljust(512, b"\0"))
            fat.append(len(sectors) if pos + 512 < len(data) else END)
        return first

    for i, name in enumerate(names[2:], 2):
        data = streams[name]
        entries[i]["size"] = len(data)
        if 0 < len(data) < 4096:
            entries[i]["start"] = len(mini_fat)
            for pos in range(0, len(data), 64):
                mini_data.extend(data[pos:pos + 64].ljust(64, b"\0"))
                mini_fat.append(len(mini_fat) + 1 if pos + 64 < len(data) else END)
        else:
            entries[i]["start"] = allocate(data)
    entries[0]["start"] = allocate(bytes(mini_data))
    entries[0]["size"] = len(mini_data)
    mini_count = math.ceil(len(mini_fat) * 4 / 512)
    mini_start = allocate(b"".join(struct.pack("<I", n) for n in mini_fat).ljust(mini_count * 512, b"\xff"))
    directory = bytearray()
    for e in entries:
        name = (e["name"] + "\0").encode("utf-16le")
        assert len(name) <= 64
        d = bytearray(128)
        d[:len(name)] = name
        struct.pack_into("<HBBIII", d, 64, len(name), e["type"], e["color"], e["left"], e["right"], e["child"])
        struct.pack_into("<IQ", d, 116, e["start"], e["size"])
        directory.extend(d)
    directory_start = allocate(directory)
    fat_count = 1
    while fat_count * 128 < len(sectors) + fat_count:
        fat_count += 1
    assert fat_count <= 109, "DIFAT extension is not needed for this small workbook"
    fat_ids = list(range(len(sectors), len(sectors) + fat_count))
    fat.extend([FAT] * fat_count)
    fat_bytes = b"".join(struct.pack("<I", n) for n in fat).ljust(fat_count * 512, b"\xff")
    sectors.extend(fat_bytes[p:p + 512] for p in range(0, len(fat_bytes), 512))
    header = bytearray(512)
    header[:8] = bytes.fromhex("d0cf11e0a1b11ae1")
    struct.pack_into("<HHHHH", header, 24, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIII", header, 40, 0, fat_count, directory_start, 0, 4096, mini_start, mini_count, END, 0)
    for i in range(109):
        struct.pack_into("<I", header, 76 + 4 * i, fat_ids[i] if i < fat_count else FREE)
    return bytes(header) + b"".join(sectors)


def record(kind: int, value: bytes) -> bytes:
    return struct.pack("<HI", kind, len(value)) + value


def module_record(name: str, document: bool = False) -> bytes:
    ansi, wide = name.encode("cp936"), name.encode("utf-16le")
    return (record(0x19, ansi) + record(0x47, wide) + record(0x1A, ansi) + record(0x32, wide)
            + record(0x1C, b"") + record(0x48, b"") + record(0x31, struct.pack("<I", 0))
            + record(0x1E, struct.pack("<I", 0)) + record(0x2C, struct.pack("<H", 0xFFFF))
            + record(0x22 if document else 0x21, b"") + record(0x2B, b""))


def rebuild_vba(original: bytes, new_source: Path, replacements: dict[str, Path] | None = None) -> bytes:
    ole = olefile.OleFileIO(io.BytesIO(original))
    project = VBA_Project(ole, "", "PROJECT", "VBA/dir")
    project.parse_project_stream()
    list(project.parse_modules())
    directory = bytearray(decompress_stream(ole.openstream("VBA/dir").read()))
    # Keep only the workbook, WBS document, model, and the two v5 modules.
    # The original template also contains Data/Resource document modules and a
    # recorded-formatting standard module; retaining those makes Excel expose
    # deleted sheets and leaves stale references in the VBA project.
    wanted = {"ThisWorkbook", "Sheet1", "Model_1"}
    module_map = {module.name: module for module in project.modules}
    streams = {}
    for name in sorted(wanted):
        module = module_map[name]
        source_path = (replacements or {}).get(name)
        source = (source_path.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\n", "\r\n").encode("cp936")
                  if source_path is not None else module.code_raw)
        streams["VBA/" + module.streamname] = compress_vba(source)
    # The module records are the final records in the directory stream. Keep
    # the references and project metadata before them, then write fresh records
    # with zero text offsets matching the rebuilt source streams.
    first_module = min(directory.index(record(0x19, module.name.encode("cp936")))
                         for module in project.modules)
    prefix = bytearray(directory[:first_module])
    old_count = record(0x0F, struct.pack("<H", len(project.modules)))
    assert prefix.count(old_count) == 1
    count_offset = prefix.index(old_count) + 6
    struct.pack_into("<H", prefix, count_offset, 5)
    directory = prefix + module_record("ThisWorkbook", document=True) + module_record("Sheet1", document=True)
    directory += module_record("Model_1") + module_record("ProgressChart") + module_record("Sheet4", document=True)
    directory += record(0x10, b"")
    source = new_source.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\n", "\r\n")
    streams["VBA/ProgressChart"] = compress_vba(source.encode("cp936"))
    sheet_source = (replacements or {}).get("Sheet4", new_source.with_name("Sheet4.cls"))
    source = sheet_source.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\n", "\r\n")
    streams["VBA/Sheet4"] = compress_vba(source.encode("cp936"))
    streams["VBA/dir"] = compress_vba(bytes(directory))
    # MS-OVBA 2.3.4.1: 0xFFFF denotes no implementation-specific version/cache.
    streams["VBA/_VBA_PROJECT"] = bytes.fromhex("cc61ffff000000")
    project_text = ole.openstream("PROJECT").read().decode("cp936")
    project_text = re.sub(r"Document=Sheet2/[^\r\n]*\r?\n", "", project_text)
    project_text = re.sub(r"Document=Sheet3/[^\r\n]*\r?\n", "", project_text)
    project_text = re.sub(r"Module=模块2\r?\n", "", project_text)
    project_text = re.sub(r"(?m)^(?:Sheet[23]|模块2)=[^\r\n]*\r?\n", "", project_text)
    project_text = project_text.replace('Module=Model_1\r\n', 'Module=Model_1\r\nModule=ProgressChart\r\nDocument=Sheet4/&H00000000\r\n')
    streams["PROJECT"] = project_text.encode("cp936")
    wm = ole.openstream("PROJECTwm").read()
    assert wm.endswith(b"\0\0")
    # PROJECTwm is the paired ANSI/UTF-16 module-name table. Rebuild it from
    # the exact five streams exposed by the v5 workbook.
    names = ["ThisWorkbook", "Sheet1", "Model_1", "ProgressChart", "Sheet4"]
    wm_data = bytearray()
    for name in names:
        wm_data += name.encode("cp936") + b"\0"
        wm_data += name.encode("utf-16le") + b"\0\0"
    streams["PROJECTwm"] = bytes(wm_data) + b"\0\0"
    result = cfb_write(streams)
    check = olefile.OleFileIO(io.BytesIO(result), raise_defects=olefile.DEFECT_INCORRECT)
    for path, expected in streams.items():
        assert check.openstream(path).read() == expected, path
    return result
