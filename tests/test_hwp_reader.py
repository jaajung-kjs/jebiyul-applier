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


from src.hwp_reader import read_hwp, HwpFormatError


def test_read_hwp_golden_values(hwp_path):
    rep = read_hwp(hwp_path)
    assert rep.dates[0] == "2026.1.1"      # 최신이 첫 열
    assert rep.half == "2026년도 상반기"
    bo = rep.rates["보통인부"]
    assert bo.wages[0] == 172068           # 현재(2026.1.1)
    assert bo.wages[1] == 171037           # 직전(2025.9.1)
    assert bo.bumun == "일반공사"
    ts = rep.rates["통신설비공"]
    assert ts.code == "1087"
    assert ts.wages == (315528, 314787, 308930, 305050)


def test_read_hwp_bumun_by_code_prefix(hwp_path):
    rep = read_hwp(hwp_path)
    # 5xxx = 기타
    assert rep.rates["전기공사기사"].bumun == "기타"
    # 화물차운전사(1049)는 코드 prefix상 일반공사(보고서 분류)
    assert rep.rates["화물차운전사"].bumun == "일반공사"


def test_read_hwp_handles_markers_and_missing(hwp_path):
    rep = read_hwp(hwp_path)
    # 미조사(**) 직종 중에는 wages가 모두 None인 경우가 있어야 함(값 파싱 시 '-' → None 확인).
    # 문서 수록 순서상 첫 **직종(연마공 1032)은 2025.9.1 한 시점만 실측치가 남아있으므로
    # 특정 인덱스가 아니라 "존재 여부"로 검증한다(실 데이터 기준 — .superpowers/sdd/task-2-report.md 참조).
    misugjo = [r for r in rep.rates.values() if r.marker == "**"]
    assert misugjo, "** 마커 직종이 있어야 함"
    assert any(all(w is None for w in r.wages) for r in misugjo)


def test_read_hwp_fail_loud_on_non_hwp(tmp_path):
    bad = tmp_path / "x.txt"
    bad.write_text("not hwp")
    with pytest.raises(HwpFormatError):
        read_hwp(str(bad))


# ── 각주 문구 다듬기 ──────────────────────────────────────────────────────
# 기준 값(5개 미만 등)은 hwp 원문에서 읽되, 산출물에 남길 필요 없는
# 교차참조 꼬리('…이므로 그 적용은 6페이지…참고')와 '직종임' 말투는 정리한다.

def test_marker_notes_are_tidied(hwp_path):
    notes = read_hwp(hwp_path).marker_notes
    assert notes["*"].endswith("직종")          # '직종임' 아님
    assert "5개 미만" in notes["*"]             # 기준 값은 원문 유지(띄어쓰기만 정리)
    assert notes["**"].endswith("조사되지 않은 직종")
    for v in notes.values():                    # 교차참조 꼬리 제거
        assert "참고하시기" not in v
        assert "페이지" not in v


def test_tidy_legend_keeps_source_number():
    """기준 숫자는 원문 그대로 — 보고서가 바뀌면 그대로 따라간다."""
    from src.hwp_reader import _tidy_legend
    assert _tidy_legend("주)「*」표시 직종은 조사현장수가 3개미만 직종임") \
        == "주)「*」표시 직종은 조사현장수가 3개 미만 직종"
    assert _tidy_legend("「**」표시 직종은 조사되지 않은 직종이므로 그 적용은 "
                        "'6페이지 4.참고사항 라.'를 참고하시기 바람") \
        == "「**」표시 직종은 조사되지 않은 직종"
