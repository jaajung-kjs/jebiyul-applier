from src import sheet_model as M


RATES = {
    "간접노무비": 0.126, "공구손료": 0.03, "산재보험료": 0.03626,
    "고용보험료": 0.0101, "건강보험료": 0.03545, "연금보험료": 0.045,
    "퇴직공제부금비": 0.023, "노인장기요양보험료": 0.1295,
    "산업안전보건관리비": 0.0185, "기타경비": 0.052, "일반관리비": 0.06, "이윤": 0.15,
}
PARAMS = dict(kind="토목", jikjeop_cost=500_000_000, days=120,
              contract="경쟁", sanjae_basis="조달청", sanan_target=300_000_000)


def test_first_block_is_title():
    blocks = M.build(PARAMS, RATES)
    assert isinstance(blocks[0], M.Title)
    assert "적용근거" in blocks[0].text


def test_has_section_headers_in_order():
    blocks = M.build(PARAMS, RATES)
    sections = [b.text for b in blocks if isinstance(b, M.SectionHeader)]
    assert sections[0].startswith("1. 간접노무비")
    assert any(s.startswith("2. 경") for s in sections)
    assert any(s.startswith("3. 일반관리비") for s in sections)
    assert any(s.startswith("4. 이") for s in sections)


def test_sanjae_note_embeds_rate():
    blocks = M.build(PARAMS, RATES)
    notes = [ln for b in blocks if isinstance(b, M.NoteLines) for ln in b.lines]
    assert any("3.626%" in ln for ln in notes)


def test_applied_rate_blocks_carry_decimal_value():
    blocks = M.build(PARAMS, RATES)
    applied = [b for b in blocks if isinstance(b, M.AppliedRate)]
    assert any(abs(b.value - 0.03626) < 1e-9 for b in applied)


def test_pct_helper():
    assert M.pct(0.03626) == "3.626"
    assert M.pct(0.03) == "3"
