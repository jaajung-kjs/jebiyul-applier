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


def test_sanan_2천만미만_shows_적용제외(monkeypatch):
    """2천만원 미만 구간은 0.00% 대신 '적용제외'로 표기한다."""
    import src.sheet_model as SM

    def fake_sanan(path, band):
        return {"rate": 0.0, "기초액": None} if band == "2천만미만" \
            else {"rate": 0.0315, "기초액": None}

    monkeypatch.setattr(SM.lookup, "sanan_rate", fake_sanan)
    monkeypatch.setattr(SM.lookup, "table_rate", lambda *a, **k: 0.12)
    blocks = M.build(dict(PARAMS, sanan_target=300_000_000), RATES, jebiyul_path="DUMMY")
    t = [b for b in blocks if isinstance(b, M.BandTable) and b.kind == "sanan"][0]
    first = t.rows[0]  # 2천만원 미만
    assert any("적용제외" in str(c[0]) for c in first.cells)
    assert all(c[0] != 0 for c in first.cells)   # 0 값이 남지 않음


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
    # 적용기준 컬럼을 갖는 구간표만 검사(etc_detail 구성비표는 적용기준 컬럼이 없음)
    for t in [b for b in blocks if isinstance(b, M.BandTable) and b.kind != "etc_detail"]:
        assert t.criterion, f"{t.kind} 표에 criterion 블록이 없음"
        for r in t.rows:
            assert all(c[1] != 7 for c in r.cells), f"{t.kind} 표 행에 G열 셀 잔존"


def test_ilban_table_is_file_driven_and_highlights_size(monkeypatch):
    """일반관리비 표는 파일값(5구간)으로 채우고 해당 규모를 강조한다(하드코딩 아님)."""
    import src.sheet_model as SM
    seen = {}

    def fake_table_rate(path, item, kind, size, dur, contract=None):
        seen[(item, size)] = 0.065 if size == "50-300억" else 0.08
        return seen[(item, size)]

    monkeypatch.setattr(SM.lookup, "table_rate", fake_table_rate)
    monkeypatch.setattr(SM.lookup, "sanan_rate",
                        lambda p, band: {"rate": 0.0315, "기초액": None})
    params = dict(PARAMS, jikjeop_cost=7_000_000_000)  # 70억 → 50-300억
    blocks = M.build(params, RATES, jebiyul_path="DUMMY")
    t = [b for b in blocks if isinstance(b, M.BandTable) and b.kind == "ilban"][0]
    assert len(t.rows) == 5                      # 5개 규모구간
    assert ("일반관리비", "50-300억") in seen     # 파일에서 읽었다
    hi = [r for r in t.rows if r.highlight]
    assert len(hi) == 1
    assert any("50억" in str(c[0]) and "300억" in str(c[0]) for c in hi[0].cells)
    # 강조 행의 적용율이 파일값 0.065
    assert any(abs(c[0] - 0.065) < 1e-9 for c in hi[0].cells if isinstance(c[0], float))


def test_iyun_table_highlights_contract_size_row(monkeypatch):
    _stub_lookup(monkeypatch)
    params = dict(PARAMS, contract="경쟁", jikjeop_cost=7_000_000_000)  # 50~300억
    blocks = M.build(params, RATES, jebiyul_path="DUMMY")
    t = [b for b in blocks if isinstance(b, M.BandTable) and b.kind == "iyun"][0]
    rows = t.rows
    hi = [r for r in rows if r.highlight]
    assert len(hi) == 1
    assert any("50억원 이상" in str(c[0]) for c in hi[0].cells)


def test_etc_composition_table_present(monkeypatch):
    """기타경비 경비 구성비율 세부표(수도광열비·도서인쇄비·합계)가 포함된다."""
    _stub_lookup(monkeypatch)
    blocks = M.build(dict(PARAMS, jikjeop_cost=3_000_000_000), RATES, jebiyul_path="DUMMY")
    detail = [b for b in blocks if isinstance(b, M.BandTable) and b.kind == "etc_detail"]
    assert detail, "경비 구성비율 세부표가 없음"
    texts = [str(c[0]) for r in detail[0].rows for c in r.cells]
    assert "수도광열비" in texts
    assert "도서인쇄비" in texts
    assert "합계" in texts


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
