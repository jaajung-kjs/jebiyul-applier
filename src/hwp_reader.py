"""임금실태조사 hwp(HWP 5.x) → 노임 데이터. 표 컨트롤을 grid로 디코드한다.

위치/문단 인덱스에 의존하지 않는다. 표 셀의 (row,col)을 HWP 레코드에서 복원하므로
분기별 양식이 바뀌어도(직종 수·페이지 분할) 동작한다.
"""
import re
import struct
import zlib
from dataclasses import dataclass

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


BUMUN_ORDER = ("일반공사", "광전자", "국가유산", "원자력", "기타")
_BUMUN_BY_PREFIX = {"1": "일반공사", "2": "광전자", "3": "국가유산",
                    "4": "원자력", "5": "기타"}

_DATE = re.compile(r"^\d{4}\.\d{1,2}\.\d{1,2}$")
_CODE = re.compile(r"^(\*{0,2})(\d{4})$")
_WAGE = re.compile(r"^-?[\d,]+")


@dataclass(frozen=True)
class NomuRate:
    seq: int
    code: str
    name: str
    bumun: str
    wages: tuple      # 최신-우선, 미조사는 None
    marker: str       # "" | "*" | "**"


@dataclass(frozen=True)
class NomuReport:
    dates: tuple
    half: str
    order: tuple
    rates: dict
    marker_notes: dict = None   # {"*": 각주문장, "**": ...} — hwp 원문에서 추출


def _row_cells(grid, r):
    if not grid:
        return []
    maxc = max(c for (rr, c) in grid if rr == r)
    return ["".join(grid.get((r, c), "")).strip() for c in range(maxc + 1)]


def _grid_rows(grid):
    return sorted({r for (r, _c) in grid})


def _is_nomu_grid(grid):
    """헤더행에 날짜 칸이 있고 데이터행에 코드가 있으면 노임표."""
    has_date = has_code = False
    for r in _grid_rows(grid):
        cells = _row_cells(grid, r)
        if any(_DATE.match(x) for x in cells):
            has_date = True
        if cells and _CODE.match(cells[0]):
            has_code = True
    return has_date and has_code


def _parse_wage(text):
    """숫자(,) 접두부만 취한다. 표 마지막 셀에는 표 뒤 각주 단락이 이어붙는
    OLE 레코드 스트림 특성상 꼬리 텍스트(각주·제어문자)가 섞일 수 있다."""
    t = text.strip()
    if t in ("", "-"):
        return None
    m = _WAGE.match(t)
    if not m:
        return None
    return int(m.group().replace(",", ""))


def _half(newest_date):
    y, m, _d = newest_date.split(".")
    return f"{y}년도 {'상반기' if int(m) <= 6 else '하반기'}"


def read_hwp(path):
    grids = [g for g in read_grids(path) if _is_nomu_grid(g)]
    if not grids:
        raise HwpFormatError("개별직종 노임단가 표를 찾지 못함 — 양식 불일치")

    dates = None
    order = []
    rates = {}
    seq = 0
    for grid in grids:
        for r in _grid_rows(grid):
            cells = _row_cells(grid, r)
            date_cols = [x for x in cells if _DATE.match(x)]
            if date_cols:                       # 헤더행(페이지마다 반복) → 날짜만 취하고 스킵
                if dates is None:
                    dates = tuple(date_cols)    # 최신-우선(왼→오)
                continue
            if not cells or not _CODE.match(cells[0]):
                continue                        # 데이터행 아님
            m = _CODE.match(cells[0])
            marker, code = m.group(1), m.group(2)
            name = cells[1].strip()
            if not name:
                continue
            wage_cells = cells[2:2 + (len(dates) if dates else 4)]
            wages = tuple(_parse_wage(x) for x in wage_cells)
            bumun = _BUMUN_BY_PREFIX.get(code[0], "기타")
            seq += 1
            rates[name] = NomuRate(seq, code, name, bumun, wages, marker)
            order.append(name)

    if dates is None or len(rates) < 50:
        raise HwpFormatError(f"노임 데이터 파싱 실패(직종 {len(rates)}개) — 양식 불일치")
    marker_notes = _marker_legends(_read_paragraphs(path))
    return NomuReport(dates, _half(dates[0]), tuple(order), rates, marker_notes)


def _read_paragraphs(path):
    """hwp 본문의 모든 문단 텍스트(표 밖 포함)를 순서대로 반환한다."""
    ole = olefile.OleFileIO(path)
    out = []
    for buf in _section_bytes(ole):
        for tag, _lvl, p in _records(buf):
            if tag == _PARA_TEXT:
                out.append(_clean(p))
    return out


def _marker_legends(paras):
    """직종번호 마커(*,**) 설명 문장을 hwp 각주에서 그대로 읽어온다.

    기준 문구('5개 미만' 등)를 하드코딩하지 않고 원문을 쓰므로, 보고서가 기준을
    바꿔도 자동 반영된다. 전용 각주 줄만 고른다:
    - '*' 줄: 「*」는 있고 「**」는 없는 문단
    - '**' 줄: 「**」는 있고 「*」는 없는 문단(둘 다 든 통합 안내문은 제외)
    """
    notes = {}
    for t in paras:
        s = t.strip()
        has1, has2 = ("「*」" in s), ("「**」" in s)
        if has2 and not has1 and "**" not in notes:
            notes["**"] = s
        elif has1 and not has2 and "*" not in notes:
            notes["*"] = s
    return notes
