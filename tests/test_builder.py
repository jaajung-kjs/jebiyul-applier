"""src/builder.py — 마스터 양식에 값 채우기 검증."""
import os

import openpyxl
import pytest

from src.builder import build_output, HEADLINE_CELL

SAMPLE_RATES = {
    "간접노무비": 0.195,
    "공구손료": 0.03,
    "산재보험료": 0.03656,
    "고용보험료": 0.0101,
    "건강보험료": 0.03595,
    "연금보험료": 0.0475,
    "퇴직공제부금비": 0.023,
    "노인장기요양보험료": 0.1314,
    "산업안전보건관리비": 0.0315,
    "기타경비": 0.057,
    "일반관리비": 0.08,
    "이윤": 0.15,
}


def test_build_returns_path(tmp_path):
    out_path = str(tmp_path / "out.xlsx")
    assert build_output(SAMPLE_RATES, out_path) == out_path
    assert os.path.isfile(out_path)


def test_headline_cells_written(tmp_path):
    """헤드라인 ☞ 적용율 셀(I열)에 각 항목 율이 들어간다."""
    out = build_output(SAMPLE_RATES, str(tmp_path / "out.xlsx"))
    ws = openpyxl.load_workbook(out)["적용근거"]
    for item, cell in HEADLINE_CELL.items():
        assert ws[cell].value == pytest.approx(SAMPLE_RATES[item]), (
            f"{item} → {cell} 불일치: {ws[cell].value}"
        )
    # 산안비 '사급 포함시'(I55)도 동일값
    assert ws["I55"].value == pytest.approx(SAMPLE_RATES["산업안전보건관리비"])


def test_inline_rate_text_rewritten(tmp_path):
    """설명 문구에 박힌 보험요율이 갱신된 율로 다시 써진다."""
    out = build_output(SAMPLE_RATES, str(tmp_path / "out.xlsx"))
    ws = openpyxl.load_workbook(out)["적용근거"]
    assert "3.656%" in ws["A23"].value      # 산재
    assert "3.595%" in ws["A32"].value      # 건강
    assert "13.14%" in ws["A44"].value      # 노인장기요양


def test_structure_preserved(tmp_path):
    """값만 채우고 양식(섹션 헤더·병합)은 그대로 유지된다."""
    out = build_output(SAMPLE_RATES, str(tmp_path / "out.xlsx"))
    ws = openpyxl.load_workbook(out)["적용근거"]
    assert ws["A2"].value == "8. 공사비 산출 적용근거"
    assert "산업안전보건관리비" in ws["A45"].value
    assert len(ws.merged_cells.ranges) > 50


def test_tables_filled_from_file(tomok_path, tmp_path):
    """params+제비율 경로가 주어지면 간접노무비/산안비/이윤 구간표가 파일값으로 채워진다."""
    params = dict(jikjeop_cost=500_000_000, days=200, kind="토목", contract="경쟁",
                  sanjae_basis="한전", sanan_target=400_000_000, jebiyul_path=tomok_path)
    out = build_output(SAMPLE_RATES, str(tmp_path / "out.xlsx"),
                       params=params, jebiyul_path=tomok_path)
    ws = openpyxl.load_workbook(out)["적용근거"]
    # 간접노무비표: 이번 규모(10억미만) 라벨 + 183일 값
    assert ws["B7"].value == "10억 미만"
    assert ws["E7"].value == pytest.approx(0.191)
    # 이윤표 경쟁 50억미만 = 0.15, 수의 1000억이상 = 0.09
    assert ws["F87"].value == pytest.approx(0.15)
    assert ws["F92"].value == pytest.approx(0.09)
    # 산안비 5억미만 구간
    assert ws["D48"].value == pytest.approx(0.0315)
