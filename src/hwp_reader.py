"""임금실태조사 hwp(HWP 5.x) → 노임 데이터. 표 컨트롤을 grid로 디코드한다.

위치/문단 인덱스에 의존하지 않는다. 표 셀의 (row,col)을 HWP 레코드에서 복원하므로
분기별 양식이 바뀌어도(직종 수·페이지 분할) 동작한다.
"""
import struct
import zlib

import olefile

# HWP BodyText 레코드 태그
_PARA_TEXT = 67
_CTRL_HEADER = 71
_LIST_HEADER = 72
_TABLE = 76


class HwpFormatError(Exception):
    """hwp 양식이 기대와 달라 안전하게 파싱할 수 없을 때."""


def _records(buf):
    """(tag, level, payload) 스트림. HWP 레코드 헤더 = tag(10)+level(10)+size(12)."""
    i = 0
    while i + 4 <= len(buf):
        h = struct.unpack("<I", buf[i:i + 4])[0]
        tag = h & 0x3FF
        level = (h >> 10) & 0x3FF
        size = (h >> 20) & 0xFFF
        i += 4
        if size == 0xFFF:  # 확장 크기
            size = struct.unpack("<I", buf[i:i + 4])[0]
            i += 4
        yield tag, level, buf[i:i + size]
        i += size


def _clean(payload):
    """PARA_TEXT(utf-16-le)에서 제어문자 제거."""
    s = payload.decode("utf-16-le", "ignore")
    return "".join(ch for ch in s if ord(ch) >= 32)


def _section_bytes(ole):
    """BodyText/Section* 스트림을 압축 해제해 이어붙인다."""
    out = []
    for entry in ole.listdir():
        if len(entry) == 2 and entry[0] == "BodyText" and entry[1].startswith("Section"):
            raw = ole.openstream(entry).read()
            try:
                out.append((entry[1], zlib.decompress(raw, -15)))
            except zlib.error:
                out.append((entry[1], raw))  # 비압축 문서
    out.sort()  # Section0, Section1 ...
    return [b for _, b in out]


def read_grids(path):
    """hwp의 모든 표를 {(row,col): 텍스트} grid 리스트로 반환한다."""
    if not olefile.isOleFile(path):
        raise HwpFormatError(f"HWP(OLE) 파일이 아님: {path}")
    ole = olefile.OleFileIO(path)
    if "FileHeader" not in ["/".join(s) for s in ole.listdir()] or \
            not ole.openstream("FileHeader").read(17) == b"HWP Document File":
        raise HwpFormatError(f"HWP 5.x 문서가 아님(FileHeader 서명 불일치): {path}")
    grids = []
    for buf in _section_bytes(ole):
        cur = None
        pos = None
        for tag, _level, p in _records(buf):
            if tag == _CTRL_HEADER and len(p) >= 4 and p[:4][::-1] == b"tbl ":
                cur = {}
                pos = None
                grids.append(cur)
            elif tag == _LIST_HEADER and cur is not None and len(p) >= 12:
                col = struct.unpack("<H", p[8:10])[0]
                row = struct.unpack("<H", p[10:12])[0]
                pos = (row, col)
            elif tag == _PARA_TEXT and cur is not None and pos is not None:
                cur[pos] = cur.get(pos, "") + _clean(p)
    return grids
