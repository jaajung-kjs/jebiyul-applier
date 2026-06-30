"""율-테이블 골든 테스트. 실제 xlsx 파일의 셀 값으로 검증.

파일 기반 골든 값 (probe 확인):
  토목 파일 ('26. 4. 13.' 시트):
    AC15=19.1, AC17=19.5, AF15=18.4   (간접노무비)
    AL15=5.5, AO15=4.9               (기타경비)
    BD15=8,   BD43=4.5               (일반관리비, col56)
    BS15=15,  BS43=9                 (이윤, col71)
    BD15=8 (10-50억 merged → row 15  (일반관리비 10-50억, probe=8)
    BS43=9  (이윤 1000억이상 경쟁 파일값, probe=9)
  건축 파일 ('건축제비율(4.13.)' 시트):
    AC15=17.5                        (간접노무비)
    AL15=5.0                         (기타경비)
    BE15=8                           (일반관리비, col57)
    BR15=15                          (이윤, col70)
"""
import pytest
from src.lookup import table_rate


# ── 간접노무비 (golden: 브리핑 제공 검증값) ──────────────────────────────

def test_tomok_ganjeop_10eok_183(tomok_path):
    assert table_rate(tomok_path, "간접노무비", "토목", "10억미만", "183") == pytest.approx(0.191)

def test_tomok_ganjeop_10eok_365(tomok_path):
    assert table_rate(tomok_path, "간접노무비", "토목", "10억미만", "365") == pytest.approx(0.195)

def test_tomok_ganjeop_jogyeong(tomok_path):
    assert table_rate(tomok_path, "간접노무비", "조경", "10억미만", "183") == pytest.approx(0.184)

def test_geonchuk_ganjeop_10eok_183(geonchuk_path):
    assert table_rate(geonchuk_path, "간접노무비", "건축", "10억미만", "183") == pytest.approx(0.175)


# ── 기타경비 (probe 확인: AL15=5.5, AO15=4.9) ──────────────────────────

def test_tomok_gita_10eok_183(tomok_path):
    assert table_rate(tomok_path, "기타경비", "토목", "10억미만", "183") == pytest.approx(0.055)

def test_tomok_gita_jogyeong_10eok_183(tomok_path):
    assert table_rate(tomok_path, "기타경비", "조경", "10억미만", "183") == pytest.approx(0.049)

def test_geonchuk_gita_10eok_183(geonchuk_path):
    # 건축 파일 AL15=5.0
    assert table_rate(geonchuk_path, "기타경비", "건축", "10억미만", "183") == pytest.approx(0.05)


# ── 일반관리비 (기간 무관; probe 확인: BD15=8, BD43=4.5) ────────────────

def test_tomok_ilban_10eok(tomok_path):
    # 공사원가 5억미만 행(BD15) — duration은 무시됨
    assert table_rate(tomok_path, "일반관리비", "토목", "10억미만", "183") == pytest.approx(0.08)

def test_tomok_ilban_1000eok(tomok_path):
    # 공사원가 1000억이상 행(BD43)
    assert table_rate(tomok_path, "일반관리비", "토목", "1000억이상", "183") == pytest.approx(0.045)

def test_geonchuk_ilban_10eok(geonchuk_path):
    # 건축 파일 BE15=8
    assert table_rate(geonchuk_path, "일반관리비", "건축", "10억미만", "183") == pytest.approx(0.08)


# ── 이윤 (기간 무관; probe 확인: BS15=15, BS43=9) ───────────────────────

def test_tomok_iyun_10eok(tomok_path):
    assert table_rate(tomok_path, "이윤", "토목", "10억미만", "183") == pytest.approx(0.15)

def test_tomok_iyun_1000eok(tomok_path):
    assert table_rate(tomok_path, "이윤", "토목", "1000억이상", "183") == pytest.approx(0.09)

def test_geonchuk_iyun_10eok(geonchuk_path):
    # 건축 파일 BR15=15
    assert table_rate(geonchuk_path, "이윤", "건축", "10억미만", "183") == pytest.approx(0.15)


# ── Fix 1: LookupError 경로 테스트 ──────────────────────────────────────────

def test_bad_label_raises(tomok_path):
    import openpyxl
    from src.lookup import cell_text_search
    ws = openpyxl.load_workbook(tomok_path, data_only=True).worksheets[0]
    with pytest.raises(LookupError):
        cell_text_search(ws, "[절대없는라벨XYZ]")

def test_no_data_sheet_raises(tmp_path):
    import openpyxl
    from src.lookup import find_data_sheet
    wb = openpyxl.Workbook(); wb.active.title = "빈시트"
    with pytest.raises(LookupError):
        find_data_sheet(wb)


# ── Fix 2: 수의 이윤율 하드코딩 ──────────────────────────────────────────────

def test_iyun_gyeongjaeng_uses_file(tomok_path):
    # 경쟁(또는 None) → 파일값. 1000억이상 = 0.09 (probe 확인: BS43=9)
    assert table_rate(tomok_path, "이윤", "토목", "1000억이상", "183", "경쟁") == pytest.approx(0.09)

def test_iyun_suui_under_1000_hardcoded(tomok_path):
    # 수의계약 + 1000억미만 → 조달청 기준 하드코딩 0.10
    assert table_rate(tomok_path, "이윤", "토목", "50-300억", "183", "수의") == pytest.approx(0.10)

def test_iyun_suui_over_1000_hardcoded(tomok_path):
    # 수의계약 + 1000억이상 → 조달청 기준 하드코딩 0.09
    assert table_rate(tomok_path, "이윤", "토목", "1000억이상", "183", "수의") == pytest.approx(0.09)


# ── Fix 3: 병합 셀 워크어라운드 고정 테스트 ──────────────────────────────────

def test_tomok_ilban_10_50eok_merged(tomok_path):
    # 10-50억 → SIZE_BASE_ROW_ILBAN=15 (BD15:BK30 merged, probe 확인: BD15=8)
    assert table_rate(tomok_path, "일반관리비", "토목", "10-50억", "183") == pytest.approx(0.08)
