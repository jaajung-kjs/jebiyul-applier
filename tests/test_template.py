"""마스터 적용근거 템플릿(실제 설계서 시트 추출본)의 구조 검증."""
import os

import openpyxl

ASSET = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "assets", "template_적용근거.xlsx",
)


def _ws():
    return openpyxl.load_workbook(ASSET)["적용근거"]


def test_template_exists():
    assert os.path.exists(ASSET)


def test_sheet_name():
    assert "적용근거" in openpyxl.load_workbook(ASSET).sheetnames


def test_section_headers_present():
    ws = _ws()
    assert ws["A2"].value == "8. 공사비 산출 적용근거"
    assert ws["A4"].value == "1. 간접노무비"
    assert ws["A16"].value.strip().startswith("2.")
    assert "산업안전보건관리비" in ws["A45"].value
    assert ws["A76"].value == "3. 일반관리비"
    assert ws["A85"].value.strip().startswith("4.")


def test_rich_formatting_preserved():
    """추출 시 병합 셀(서식)이 유지되어야 한다(골격이 아니라 실제 양식)."""
    ws = _ws()
    assert len(ws.merged_cells.ranges) > 50


def test_headline_rate_label_rows():
    """☞ 적 용 율 라벨이 헤드라인 셀 행에 존재한다."""
    ws = _ws()
    assert ws["A14"].value.strip().startswith("☞ 적 용 율")   # 간접노무비
    assert ws["A19"].value.strip().startswith("☞ 적 용 율")   # 공구손료
    assert ws["A83"].value.strip().startswith("☞ 적 용 율")   # 일반관리비
