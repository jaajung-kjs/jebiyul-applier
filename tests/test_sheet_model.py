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


def _stub_lookup(monkeypatch):
    """build()가 부르는 모든 제비율 조회를 가짜로 대체(파일 불필요)."""
    import src.sheet_model as SM
    monkeypatch.setattr(SM.lookup, "sanan_rate",
                        lambda p, band: {"rate": 0.0315, "기초액": None})
    monkeypatch.setattr(SM.lookup, "table_rate", lambda *a, **k: 0.12)


def test_sanan_table_four_bands_highlight_applicable(monkeypatch):
    _stub_lookup(monkeypatch)
    params = dict(PARAMS, sanan_target=300_000_000)  # 5억 미만
    blocks = M.build(params, RATES, jebiyul_path="DUMMY")
    t = [b for b in blocks if isinstance(b, M.BandTable) and b.kind == "sanan"][0]
    assert len(t.rows) == 4
    hi = [r for r in t.rows if r.highlight]
    assert len(hi) == 1
    assert any("5억원 미만" in str(c[0]) for c in hi[0].cells)


def test_sanan_criterion_is_single_merged_block(monkeypatch):
    """산안비 적용기준은 다른 표처럼 세로 병합 한 칸으로 설명한다(행별 분할 금지)."""
    _stub_lookup(monkeypatch)
    blocks = M.build(dict(PARAMS, sanan_target=300_000_000), RATES, jebiyul_path="DUMMY")
    t = [b for b in blocks if isinstance(b, M.BandTable) and b.kind == "sanan"][0]
    assert t.criterion  # 세로 병합 블록으로 설명
    # 데이터 행에는 적용기준(G열=7) 셀이 없어야 한다(규모/율 2칸만)
    for r in t.rows:
        assert all(c[1] != 7 for c in r.cells), "행별 G열 적용기준 셀이 남아있음"


def test_all_tables_use_criterion_block(monkeypatch):
    """모든 구간표가 적용기준을 세로 병합 한 칸(criterion)으로 통일해 그린다."""
    _stub_lookup(monkeypatch)
    blocks = M.build(dict(PARAMS, jikjeop_cost=3_000_000_000), RATES, jebiyul_path="DUMMY")
    for t in [b for b in blocks if isinstance(b, M.BandTable)]:
        assert t.criterion, f"{t.kind} 표에 criterion 블록이 없음"
        for r in t.rows:
            assert all(c[1] != 7 for c in r.cells), f"{t.kind} 표 행에 G열 셀 잔존"


def test_ilban_table_highlights_applicable_size(monkeypatch):
    _stub_lookup(monkeypatch)
    params = dict(PARAMS, jikjeop_cost=7_000_000_000)  # 30~100억
    blocks = M.build(params, RATES, jebiyul_path="DUMMY")
    t = [b for b in blocks if isinstance(b, M.BandTable) and b.kind == "ilban"][0]
    assert len(t.rows) == 4
    hi = [r for r in t.rows if r.highlight]
    assert len(hi) == 1
    assert any("30억원 이상" in str(c[0]) for c in hi[0].cells)


def test_iyun_table_highlights_contract_size_row(monkeypatch):
    _stub_lookup(monkeypatch)
    params = dict(PARAMS, contract="경쟁", jikjeop_cost=7_000_000_000)  # 50~300억
    blocks = M.build(params, RATES, jebiyul_path="DUMMY")
    t = [b for b in blocks if isinstance(b, M.BandTable) and b.kind == "iyun"][0]
    rows = t.rows
    hi = [r for r in rows if r.highlight]
    assert len(hi) == 1
    assert any("50억원 이상" in str(c[0]) for c in hi[0].cells)


def _gibon(monkeypatch, params):
    _stub_lookup(monkeypatch)
    blocks = M.build(params, RATES, jebiyul_path="DUMMY")
    return [b for b in blocks if isinstance(b, M.BandTable) and b.kind == "gibon"]


def test_gibon_under_50_compact_with_reference_row(monkeypatch):
    params = dict(PARAMS, jikjeop_cost=3_000_000_000, days=120)  # 30억 <50억, 183일
    tables = _gibon(monkeypatch, params)
    assert len(tables) == 2  # 간접노무비 + 기타경비
    t = tables[0]
    labels = [str(r.cells[0][0]) for r in t.rows]
    assert "50억 미만" in labels[0]
    assert any("50억 이상" in l for l in labels)        # 참조행 존재
    hi = [r for r in t.rows if r.highlight]
    assert len(hi) == 1


def test_gibon_over_50_expands_actual_band(monkeypatch):
    params = dict(PARAMS, jikjeop_cost=100_000_000_000, days=400)  # 1000억, 365일
    t = _gibon(monkeypatch, params)[0]
    labels = [str(r.cells[0][0]) for r in t.rows]
    assert any("1000억" in l for l in labels)
    assert all("조달청 발표자료 참조" not in str(r.cells) for r in t.rows)
