import olefile
import pytest

from src.hwp_reader import HwpFormatError, read_grids


class _FakeStream:
    def __init__(self, data):
        self._data = data

    def read(self, n=-1):
        return self._data[:n] if n is not None and n >= 0 else self._data


class _FakeNonHwpOle:
    """유효한 OLE지만 HWP 서명이 아닌 FileHeader 스트림을 가진 파일을 흉내낸다."""

    def __init__(self, path):
        pass

    def listdir(self):
        return [["FileHeader"]]

    def openstream(self, name):
        return _FakeStream(b"NOT-AN-HWP-FILE!!")


def test_read_grids_rejects_non_hwp_ole(monkeypatch):
    """유효한 OLE 컴파운드 파일이라도 FileHeader 서명이 HWP 5.x와 다르면
    조용히 []가 아니라 HwpFormatError를 던져야 한다(오탐 방지: fail-loud)."""
    monkeypatch.setattr(olefile, "isOleFile", lambda path: True)
    monkeypatch.setattr(olefile, "OleFileIO", _FakeNonHwpOle)
    with pytest.raises(HwpFormatError):
        read_grids("x")


def _nomu_grids(grids):
    """헤더에 날짜(YYYY.M.D)와 코드행을 가진 노임표만."""
    import re
    date = re.compile(r"\d{4}\.\d{1,2}\.\d{1,2}")
    out = []
    for g in grids:
        joined = " ".join(g.values())
        if date.search(joined) and "통신설비공" in joined:
            out.append(g)
    return out


def test_read_grids_recovers_tongsin_row(hwp_path):
    grids = read_grids(hwp_path)
    assert len(grids) > 10  # 문서에 표가 여럿
    hit = _nomu_grids(grids)
    assert hit, "통신설비공이 든 노임표를 찾지 못함"
    g = hit[0]
    # 통신설비공 행 전체 복원
    pos = next((r, c) for (r, c), v in g.items() if v.strip() == "통신설비공")
    r, _ = pos
    maxc = max(c for (_r, c) in g if _r == r)
    row = ["".join(g.get((r, c), "")).strip() for c in range(maxc + 1)]
    assert row[0] == "1087"
    assert row[1] == "통신설비공"
    assert row[2].replace(",", "") == "315528"
    assert row[5].replace(",", "") == "305050"
