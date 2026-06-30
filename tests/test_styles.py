from openpyxl.styles import Font, PatternFill, Alignment
from src import styles


def test_section_font_is_bold_blue():
    f = styles.section_font()
    assert isinstance(f, Font)
    assert f.bold is True
    assert f.name == styles.BODY_FONT_NAME
    assert f.color.rgb == styles.SECTION_COLOR


def test_rate_font_is_bold_red():
    f = styles.rate_font()
    assert f.bold is True
    assert f.color.rgb == styles.RATE_COLOR


def test_header_fill_is_light_blue():
    fill = styles.header_fill()
    assert isinstance(fill, PatternFill)
    assert fill.patternType == "solid"
    assert fill.fgColor.rgb == styles.HEADER_FILL_RGB


def test_col_widths_cover_key_columns():
    assert styles.COL_WIDTHS["A"] == 4.2
    assert styles.COL_WIDTHS["K"] == 2.5


def test_center_alignment_vertical_center():
    al = styles.center()
    assert isinstance(al, Alignment)
    assert al.vertical == "center"
    assert al.horizontal == "center"
