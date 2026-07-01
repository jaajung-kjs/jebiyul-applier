import openpyxl
from src import sheet_model as M
from src import renderer


def _blocks():
    return [
        M.Title("8. 공사비 산출 적용근거"),
        M.SectionHeader("1. 간접노무비", M.NOTE),
        M.NoteLines(["    ☞ 계상금액 : 직접노무비 × 적용율"]),
        M.AppliedRate("    ☞ 적 용 율 :  ", 0.126, fmt="0.0%"),
    ]


def test_render_writes_title_merged(tmp_path):
    out = str(tmp_path / "o.xlsx")
    renderer.render(_blocks(), out)
    ws = openpyxl.load_workbook(out)["적용근거"]
    assert ws["A2"].value == "8. 공사비 산출 적용근거"
    assert "A2:J2" in [str(m) for m in ws.merged_cells.ranges]


def test_section_header_blue_bold(tmp_path):
    out = str(tmp_path / "o.xlsx")
    renderer.render(_blocks(), out)
    ws = openpyxl.load_workbook(out)["적용근거"]
    found = None
    for row in ws.iter_rows():
        for c in row:
            if c.value == "1. 간접노무비":
                found = c
    assert found is not None
    assert found.font.bold is True
    assert found.font.color.rgb == "FF0000FF"


def test_applied_rate_value_in_I_with_percent_format(tmp_path):
    out = str(tmp_path / "o.xlsx")
    renderer.render(_blocks(), out)
    ws = openpyxl.load_workbook(out)["적용근거"]
    hit = None
    for c in ws["I"]:
        if isinstance(c.value, (int, float)) and abs(c.value - 0.126) < 1e-9:
            hit = c
    assert hit is not None
    assert hit.number_format == "0.0%"
    assert hit.font.color.rgb == "FFFF0000"


def test_applied_rate_annotation_in_J(tmp_path):
    blocks = [M.AppliedRate("    ☞ 적 용 율 :  (사급재료비 제외시)", 0.0315,
                            fmt="0.000%", annotation="× 1.2")]
    out = str(tmp_path / "a.xlsx")
    renderer.render(blocks, out)
    ws = openpyxl.load_workbook(out)["적용근거"]
    assert any(c.value == "× 1.2" for c in ws["J"])


def test_column_widths_applied(tmp_path):
    out = str(tmp_path / "o.xlsx")
    renderer.render(_blocks(), out)
    ws = openpyxl.load_workbook(out)["적용근거"]
    assert round(ws.column_dimensions["A"].width, 1) == 4.2
    assert round(ws.column_dimensions["K"].width, 1) == 2.5


def test_table_header_filled_and_highlight(tmp_path):
    t = M.BandTable(
        kind="ilban",
        headers=[("공 사 규 모", 2, 3), ("적용율", 4, 6), ("적  용  기  준", 7, 10)],
        rows=[
            M.BandRow([("5억원 미만", 2, 3, None), (0.06, 4, 6, "0.0%")], highlight=True),
            M.BandRow([("100억원 이상", 2, 3, None), (0.045, 4, 6, "0.0%")]),
        ],
        criterion="[지입재료비+노무비+도급분경비]×적용율",
    )
    out = str(tmp_path / "t.xlsx")
    renderer.render([t], out)
    ws = openpyxl.load_workbook(out)["적용근거"]
    hdr = None
    hirow = None
    for row in ws.iter_rows():
        for c in row:
            if c.value == "공 사 규 모":
                hdr = c
            if c.value == "5억원 미만":
                hirow = c
    assert hdr.fill.fgColor.rgb == "FFDBE5F1"
    assert hirow.fill.fgColor.rgb == "FFFFF2CC"
