#!/usr/bin/env python3
from __future__ import annotations

import struct
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import unicodedata

import pandas as pd


ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC


@dataclass
class DirEntry:
    name: str
    obj_type: int
    start_sector: int
    stream_size: int


def _read_uint16(buf: bytes, off: int) -> int:
    return struct.unpack_from("<H", buf, off)[0]


def _read_uint32(buf: bytes, off: int) -> int:
    return struct.unpack_from("<I", buf, off)[0]


def _decode_xls_unicode(data: bytes, off: int, cch: int, flags: int) -> Tuple[str, int]:
    is_16bit = flags & 0x01
    has_rich = flags & 0x08
    has_asian = flags & 0x04
    start = off
    if has_rich:
        _ = _read_uint16(data, off)
        off += 2
    if has_asian:
        _ = _read_uint32(data, off)
        off += 4

    if is_16bit:
        raw = data[off : off + cch * 2]
        text = raw.decode("utf-16le", errors="replace")
        off += cch * 2
    else:
        raw = data[off : off + cch]
        text = raw.decode("latin1", errors="replace")
        off += cch

    # Skip rich-text formatting runs and FarEast extension bytes.
    if has_rich:
        rt_count = _read_uint16(data, start)
        off += rt_count * 4
    if has_asian:
        ext_len = _read_uint32(data, start + (2 if has_rich else 0))
        off += ext_len

    return text, off


def _build_cfb_streams(path: Path) -> Dict[str, bytes]:
    b = path.read_bytes()
    if b[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError(f"{path.name}: not a CFB/OLE file")

    sector_shift = _read_uint16(b, 30)
    mini_sector_shift = _read_uint16(b, 32)
    sector_size = 1 << sector_shift
    mini_sector_size = 1 << mini_sector_shift
    first_dir_sector = _read_uint32(b, 48)
    mini_stream_cutoff = _read_uint32(b, 56)
    first_mini_fat_sector = _read_uint32(b, 60)
    num_mini_fat_sectors = _read_uint32(b, 64)
    first_difat_sector = _read_uint32(b, 68)
    num_difat_sectors = _read_uint32(b, 72)

    difat = list(struct.unpack_from("<109I", b, 76))

    def sector_off(sec_id: int) -> int:
        return 512 + sec_id * sector_size

    next_difat = first_difat_sector
    for _ in range(num_difat_sectors):
        if next_difat in (ENDOFCHAIN, FREESECT):
            break
        off = sector_off(next_difat)
        entries = struct.unpack_from(f"<{sector_size // 4}I", b, off)
        difat.extend(entries[:-1])
        next_difat = entries[-1]

    fat_sectors = [s for s in difat if s not in (FREESECT, ENDOFCHAIN, DIFSECT) and s != FATSECT]
    fat: List[int] = []
    for sec in fat_sectors:
        if sec in (FREESECT, ENDOFCHAIN):
            continue
        off = sector_off(sec)
        fat.extend(struct.unpack_from(f"<{sector_size // 4}I", b, off))

    def read_chain(start_sec: int) -> bytes:
        if start_sec in (ENDOFCHAIN, FREESECT):
            return b""
        out = bytearray()
        sec = start_sec
        visited = set()
        while sec not in (ENDOFCHAIN, FREESECT):
            if sec in visited:
                break
            visited.add(sec)
            off = sector_off(sec)
            out.extend(b[off : off + sector_size])
            if sec >= len(fat):
                break
            sec = fat[sec]
        return bytes(out)

    dir_stream = read_chain(first_dir_sector)
    entries: List[DirEntry] = []
    for i in range(0, len(dir_stream), 128):
        ent = dir_stream[i : i + 128]
        if len(ent) < 128:
            continue
        name_len = _read_uint16(ent, 64)
        if name_len < 2:
            name = ""
        else:
            raw_name = ent[: name_len - 2]
            name = raw_name.decode("utf-16le", errors="replace")
        obj_type = ent[66]
        start_sector = _read_uint32(ent, 116)
        stream_size = struct.unpack_from("<Q", ent, 120)[0]
        entries.append(DirEntry(name=name, obj_type=obj_type, start_sector=start_sector, stream_size=stream_size))

    root = next((e for e in entries if e.obj_type == 5), None)
    if root is None:
        raise ValueError(f"{path.name}: root entry not found")

    root_mini_stream = read_chain(root.start_sector)[: root.stream_size]

    mini_fat: List[int] = []
    sec = first_mini_fat_sector
    visited = set()
    for _ in range(num_mini_fat_sectors + 8):
        if sec in (ENDOFCHAIN, FREESECT) or sec in visited:
            break
        visited.add(sec)
        off = sector_off(sec)
        mini_fat.extend(struct.unpack_from(f"<{sector_size // 4}I", b, off))
        if sec >= len(fat):
            break
        sec = fat[sec]

    def read_mini_chain(start_mini_sec: int) -> bytes:
        if start_mini_sec in (ENDOFCHAIN, FREESECT):
            return b""
        out = bytearray()
        sec = start_mini_sec
        visited = set()
        while sec not in (ENDOFCHAIN, FREESECT):
            if sec in visited:
                break
            visited.add(sec)
            off = sec * mini_sector_size
            out.extend(root_mini_stream[off : off + mini_sector_size])
            if sec >= len(mini_fat):
                break
            sec = mini_fat[sec]
        return bytes(out)

    streams: Dict[str, bytes] = {}
    for e in entries:
        if e.obj_type != 2 or not e.name:
            continue
        if e.stream_size < mini_stream_cutoff:
            streams[e.name] = read_mini_chain(e.start_sector)[: e.stream_size]
        else:
            streams[e.name] = read_chain(e.start_sector)[: e.stream_size]
    return streams


def _parse_biff_xls_raw(path: Path) -> pd.DataFrame:
    streams = _build_cfb_streams(path)
    wb = streams.get("Workbook") or streams.get("Book")
    if wb is None:
        raise ValueError(f"{path.name}: Workbook stream not found")

    sheet_offsets: List[int] = []
    sheet_names: List[str] = []
    sst: List[str] = []
    date_1904 = False

    i = 0
    recs: List[Tuple[int, int, bytes]] = []
    while i + 4 <= len(wb):
        rid, rlen = struct.unpack_from("<HH", wb, i)
        payload = wb[i + 4 : i + 4 + rlen]
        recs.append((rid, rlen, payload))
        i += 4 + rlen
        if rid == 0x000A:  # EOF
            # Workbook globals EOF, keep scanning to keep absolute offsets valid.
            pass

    for idx, (rid, rlen, payload) in enumerate(recs):
        if rid == 0x0022 and rlen >= 2:  # DATEMODE
            date_1904 = bool(_read_uint16(payload, 0))
        elif rid == 0x0085 and rlen >= 8:  # BOUNDSHEET
            off = _read_uint32(payload, 0)
            sheet_offsets.append(off)
            name_len = payload[6]
            flags = payload[7]
            if flags & 0x01:
                name = payload[8 : 8 + name_len * 2].decode("utf-16le", errors="replace")
            else:
                name = payload[8 : 8 + name_len].decode("latin1", errors="replace")
            sheet_names.append(name)
        elif rid == 0x00FC:  # SST (with CONTINUE support)
            total = _read_uint32(payload, 0)
            unique = _read_uint32(payload, 4)
            data = bytearray(payload[8:])
            j = idx + 1
            while j < len(recs) and recs[j][0] == 0x003C:
                data.extend(recs[j][2])
                j += 1

            p = 0
            out: List[str] = []
            for _ in range(unique):
                if p + 3 > len(data):
                    break
                cch = _read_uint16(data, p)
                p += 2
                flags = data[p]
                p += 1
                text, p = _decode_xls_unicode(data, p, cch, flags)
                out.append(text)
            sst = out

    def parse_sheet(off: int) -> Dict[Tuple[int, int], object]:
        cells: Dict[Tuple[int, int], object] = {}
        p = off
        while p + 4 <= len(wb):
            rid, rlen = struct.unpack_from("<HH", wb, p)
            payload = wb[p + 4 : p + 4 + rlen]
            p += 4 + rlen
            if rid == 0x000A:  # EOF
                break
            if rid == 0x0203 and rlen >= 14:  # NUMBER
                r, c = struct.unpack_from("<HH", payload, 0)
                v = struct.unpack_from("<d", payload, 6)[0]
                cells[(r, c)] = v
            elif rid == 0x027E and rlen >= 10:  # RK
                r, c = struct.unpack_from("<HH", payload, 0)
                rk = _read_uint32(payload, 6)
                cells[(r, c)] = _decode_rk(rk)
            elif rid == 0x00BD and rlen >= 6:  # MULRK
                r = _read_uint16(payload, 0)
                c_first = _read_uint16(payload, 2)
                c_last = _read_uint16(payload, rlen - 2)
                count = c_last - c_first + 1
                for k in range(count):
                    base = 4 + k * 6
                    rk = _read_uint32(payload, base + 2)
                    cells[(r, c_first + k)] = _decode_rk(rk)
            elif rid == 0x00FD and rlen >= 10:  # LABELSST
                r, c = struct.unpack_from("<HH", payload, 0)
                sst_idx = _read_uint32(payload, 6)
                text = sst[sst_idx] if 0 <= sst_idx < len(sst) else ""
                cells[(r, c)] = text
            elif rid == 0x0204 and rlen >= 8:  # LABEL (BIFF2-7)
                r, c = struct.unpack_from("<HH", payload, 0)
                ln = _read_uint16(payload, 6)
                cells[(r, c)] = payload[8 : 8 + ln].decode("latin1", errors="replace")
            elif rid == 0x0006 and rlen >= 20:  # FORMULA
                r, c = struct.unpack_from("<HH", payload, 0)
                result = payload[6:14]
                if result[:2] == b"\xff\xff":
                    # String/boolean/error result; ignore unless STRING record follows.
                    continue
                cells[(r, c)] = struct.unpack("<d", result)[0]

        return cells

    if not sheet_offsets:
        raise ValueError(f"{path.name}: no sheet offsets found")

    # Prefer first visible data sheet.
    sheet_cells = parse_sheet(sheet_offsets[0])
    if not sheet_cells:
        raise ValueError(f"{path.name}: parsed sheet is empty")

    max_row = max(r for r, _ in sheet_cells.keys())
    max_col = max(c for _, c in sheet_cells.keys())
    rows: List[List[object]] = []
    for r in range(max_row + 1):
        row = []
        for c in range(max_col + 1):
            v = sheet_cells.get((r, c), "")
            row.append(v)
        rows.append(row)

    return pd.DataFrame(rows)


def _parse_biff_xls(path: Path) -> pd.DataFrame:
    df = _parse_biff_xls_raw(path)
    header_idx = _find_header_row(df)
    if header_idx is None:
        raise ValueError(f"{path.name}: header row not found")

    header = [str(x).strip() for x in df.iloc[header_idx].tolist()]
    data = df.iloc[header_idx + 1 :].copy()
    data.columns = header
    data = data.reset_index(drop=True)
    return data


def _decode_rk(rk: int) -> float:
    is_mult_100 = rk & 0x01
    is_int = rk & 0x02
    if is_int:
        v = rk >> 2
        if v & 0x20000000:  # signed 30-bit
            v -= 1 << 30
        out = float(v)
    else:
        # High 30 bits are the high bits of IEEE754 double; low 34 bits are zero.
        raw = (rk & 0xFFFFFFFC) << 32
        out = struct.unpack("<d", struct.pack("<Q", raw))[0]
    if is_mult_100:
        out /= 100.0
    return out


def _find_header_row(df: pd.DataFrame) -> Optional[int]:
    targets = {
        "날짜",
        "사용처",
        "금액",
        "카테고리",
        "세부 카테고리",
        "거래일시",
        "거래처",
        "승인일",
        "가맹점명",
        "승인금액",
    }
    for i in range(min(len(df), 60)):
        vals = {str(v).strip() for v in df.iloc[i].tolist() if str(v).strip()}
        if len(vals & targets) >= 3:
            return i
    return None


def _read_strict_xlsx(path: Path) -> pd.DataFrame:
    with zipfile.ZipFile(path) as zf:
        wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
        ns_main = {"m": wb_xml.tag.split("}")[0].strip("{")}
        rel_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
        rel_map = {}
        for rel in rel_xml.findall("r:Relationship", rel_ns):
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target", "")
            rel_map[rid] = target

        sheet_el = wb_xml.find("m:sheets/m:sheet", ns_main)
        if sheet_el is None:
            raise ValueError(f"{path.name}: no sheet in workbook.xml")
        rid = sheet_el.attrib.get("{http://purl.oclc.org/ooxml/officeDocument/relationships}id") or sheet_el.attrib.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        if not rid or rid not in rel_map:
            raise ValueError(f"{path.name}: sheet relationship not found")

        sheet_path = "xl/" + rel_map[rid]
        if sheet_path.startswith("xl//"):
            sheet_path = "xl/" + sheet_path[4:]

        shared: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst_xml = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            sst_ns = {"s": sst_xml.tag.split("}")[0].strip("{")}
            for si in sst_xml.findall("s:si", sst_ns):
                chunks = []
                for t in si.findall(".//s:t", sst_ns):
                    chunks.append(t.text or "")
                shared.append("".join(chunks))

        sheet_xml = ET.fromstring(zf.read(sheet_path))
        s_ns = {"s": sheet_xml.tag.split("}")[0].strip("{")}

        rows: Dict[int, Dict[int, object]] = {}

        def col_to_idx(col: str) -> int:
            n = 0
            for ch in col:
                if "A" <= ch <= "Z":
                    n = n * 26 + (ord(ch) - ord("A") + 1)
            return n - 1

        for row in sheet_xml.findall(".//s:sheetData/s:row", s_ns):
            r_idx = int(row.attrib.get("r", "1")) - 1
            rmap = rows.setdefault(r_idx, {})
            for c in row.findall("s:c", s_ns):
                ref = c.attrib.get("r", "")
                col = "".join(ch for ch in ref if ch.isalpha())
                c_idx = col_to_idx(col) if col else 0
                c_type = c.attrib.get("t")
                v_el = c.find("s:v", s_ns)
                if c_type == "inlineStr":
                    t_el = c.find("s:is/s:t", s_ns)
                    val = t_el.text if t_el is not None else ""
                elif c_type == "s":
                    if v_el is None or v_el.text is None:
                        val = ""
                    else:
                        sid = int(v_el.text)
                        val = shared[sid] if 0 <= sid < len(shared) else ""
                else:
                    if v_el is None or v_el.text is None:
                        val = ""
                    else:
                        txt = v_el.text
                        try:
                            if "." in txt:
                                val = float(txt)
                            else:
                                val = int(txt)
                        except ValueError:
                            val = txt
                rmap[c_idx] = val

        if not rows:
            return pd.DataFrame()

        max_row = max(rows)
        max_col = max((max(cols.keys()) for cols in rows.values() if cols), default=0)
        matrix: List[List[object]] = []
        for r in range(max_row + 1):
            row = []
            cmap = rows.get(r, {})
            for c in range(max_col + 1):
                row.append(cmap.get(c, ""))
            matrix.append(row)

        df = pd.DataFrame(matrix)
        header_idx = _find_header_row(df)
        if header_idx is None:
            raise ValueError(f"{path.name}: header row not found")
        header = [str(x).strip() for x in df.iloc[header_idx].tolist()]
        data = df.iloc[header_idx + 1 :].copy()
        data.columns = header
        data = data.reset_index(drop=True)
        return data


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    aliases = {
        "지출날짜": ["지출날짜", "날짜", "거래일시", "거래일", "사용일시", "일시", "승인일"],
        "사용처": ["사용처", "거래처", "가맹점명", "적요", "내역"],
        "지출금액": ["지출금액", "금액", "거래금액", "사용금액", "출금금액", "출금액", "결제금액", "승인금액"],
        "카테고리": ["카테고리", "분류", "대분류"],
        "세부 카테고리": ["세부 카테고리", "세부카테고리", "소분류"],
    }

    mapping: Dict[str, str] = {}
    cols = [str(c).strip() for c in df.columns]
    for target, cand_list in aliases.items():
        for c in cols:
            if c in cand_list:
                mapping[c] = target
                break

    out = df.rename(columns=mapping).copy()
    for col in ["지출날짜", "사용처", "지출금액", "카테고리", "세부 카테고리"]:
        if col not in out.columns:
            out[col] = pd.NA
    out = out[["지출날짜", "사용처", "지출금액", "카테고리", "세부 카테고리"]]

    out["카테고리"] = out["카테고리"].astype("string").str.strip()
    out = out[out["카테고리"].notna()]
    out = out[out["카테고리"] != ""]
    out = out[out["카테고리"] != "제외"]

    out["사용처"] = out["사용처"].astype("string").str.strip()

    raw_date = out["지출날짜"].copy()
    dt = pd.to_datetime(raw_date, errors="coerce")
    mask_nat = dt.isna()
    if mask_nat.any():
        num = pd.to_numeric(raw_date[mask_nat], errors="coerce")
        serial_mask = num.notna()
        if serial_mask.any():
            dt.loc[mask_nat[mask_nat].index[serial_mask]] = pd.to_datetime(
                num[serial_mask], unit="D", origin="1899-12-30", errors="coerce"
            )
    out["지출날짜"] = dt.dt.strftime("%Y-%m-%d")

    out["지출금액"] = pd.to_numeric(out["지출금액"], errors="coerce")
    out = out[out["지출금액"].notna()]
    out = out[out["지출금액"] > 0]

    out["세부 카테고리"] = out["세부 카테고리"].astype("string").str.strip()
    out = out.reset_index(drop=True)
    return out


def main() -> None:
    files = list(Path(".").glob("*.xls")) + list(Path(".").glob("*.xlsx"))
    files = [p for p in files if not p.name.startswith("~$")]
    name_map = {p.name: p for p in files}
    norm_map = {unicodedata.normalize("NFC", p.name): p for p in files}

    required = ["하나은행.xls", "신한은행.xlsx", "우리카드.xlsx", "현대카드.xlsx"]
    # Handle Unicode-normalized filenames.
    resolved: Dict[str, Path] = {}
    for req in required:
        if req in name_map:
            resolved[req] = name_map[req]
            continue
        nreq = unicodedata.normalize("NFC", req)
        if nreq in norm_map:
            resolved[req] = norm_map[nreq]
            continue
        raise FileNotFoundError(f"Required file not found: {req}")

    dfs: List[pd.DataFrame] = []

    hana = _parse_biff_xls(resolved["하나은행.xls"])
    dfs.append(_normalize_columns(hana))

    shinhan = pd.read_excel(resolved["신한은행.xlsx"], sheet_name=0)
    dfs.append(_normalize_columns(shinhan))

    woori = pd.read_excel(resolved["우리카드.xlsx"], sheet_name=0)
    dfs.append(_normalize_columns(woori))

    hyundai = _read_strict_xlsx(resolved["현대카드.xlsx"])
    dfs.append(_normalize_columns(hyundai))

    merged = pd.concat(dfs, ignore_index=True)
    merged = merged.sort_values(["지출날짜", "사용처"], na_position="last").reset_index(drop=True)

    out_path = Path("통합지출내역.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        merged.to_excel(writer, index=False, sheet_name="지출내역")

    print(f"saved: {out_path} rows={len(merged)}")
    print(merged.head(10).to_string(index=False))


if __name__ == "__main__":
    main()
