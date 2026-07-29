import pytest
from src.hwp_reader import read_hwp
from src import nomu_model as N


def test_standard_set_names_exist_in_hwp(hwp_path):
    """표준세트 직종명이 실제 hwp에 전부 존재해야(오탈자 방지)."""
    rep = read_hwp(hwp_path)
    missing = [n for n in N.standard_set() if n not in rep.rates]
    assert missing == [], f"hwp에 없는 표준세트 직종: {missing}"


def test_build_groups_and_delta(hwp_path):
    rep = read_hwp(hwp_path)
    blocks = N.build(rep, ["보통인부", "통신설비공", "전기공사기사"])
    kinds = [type(b).__name__ for b in blocks]
    assert kinds[0] == "NomuTitle"
    assert "NomuHeader" in kinds
    groups = [b for b in blocks if isinstance(b, N.NomuGroup)]
    # 보통인부·통신설비공=일반공사(Ⅰ), 전기공사기사=기타(Ⅱ, 앞선 부문만 카운트)
    assert [g.name for g in groups] == ["일반공사직종", "기타직종"]
    assert groups[0].roman == "Ⅰ" and groups[1].roman == "Ⅱ"
    row = next(b for b in blocks if isinstance(b, N.NomuRow) and b.name == "보통인부")
    assert row.current_wage == 172068
    assert abs(row.delta - (172068 - 171037) / 171037) < 1e-9
    assert row.past_wages == [167081, 169804, 171037]  # 오래된→최신


def test_build_avg_is_group_mean(hwp_path):
    rep = read_hwp(hwp_path)
    blocks = N.build(rep, ["보통인부", "특별인부"])
    rows = [b for b in blocks if isinstance(b, N.NomuRow)]
    avg = next(b for b in blocks if isinstance(b, N.NomuAvg))
    assert abs(avg.value - sum(r.delta for r in rows) / len(rows)) < 1e-9


def test_build_header_display_order(hwp_path):
    rep = read_hwp(hwp_path)
    blocks = N.build(rep, ["보통인부"])
    hdr = next(b for b in blocks if isinstance(b, N.NomuHeader))
    assert hdr.current_col == "2026.1.1"
    assert hdr.past_cols == ["2024.9.1", "2025.1.1", "2025.9.1"]  # 오래된→최신


def test_build_sets_past_count_from_dates(hwp_path):
    rep = read_hwp(hwp_path)
    blocks = N.build(rep, ["보통인부", "특별인부"])
    expected_p = len(rep.dates) - 1
    groups = [b for b in blocks if isinstance(b, N.NomuGroup)]
    avgs = [b for b in blocks if isinstance(b, N.NomuAvg)]
    assert groups and all(g.past_count == expected_p for g in groups)
    assert avgs and all(a.past_count == expected_p for a in avgs)


def test_preselect_marks_standard(hwp_path):
    rep = read_hwp(hwp_path)
    pairs = N.preselect(list(rep.order))
    checked = {n for n, on in pairs if on}
    assert "통신설비공" in checked
    assert "보통인부" in checked
    # 비표준(예: 도편수 등 국가유산)은 기본 미체크
    unchecked = {n for n, on in pairs if not on}
    assert unchecked, "표준세트 밖 직종은 미체크로 남아야"
    # 순서 보존
    assert [n for n, _ in pairs] == list(rep.order)
