"""고정요율·산재·산안비 lookup 단위 테스트.

실제 파일에서 탐색한 값(probe 결과):
  건강보험료    3.595 %  (AU51: '(직노) x 3.595')
  연금보험료    4.75  %  (BR51: '(직노) x 4.75')
  퇴직공제부금비 2.3   %  (AU110: '(직노) x 2.3')
  노인장기요양   13.14 %  (BD51: '(건강보험료) x 13.14')
  고용보험료    1.01  %  (AH79: 1.01  — 7등급)

산안비 canonical 4구간 (토목공사 기준, col149/col157):
  "2천만미만"  → 파일에 행 없음; 2천만 미만은 적용 대상 아님 → rate 0.0
  "5억미만"    → 3.15 % 기초액 없음   (row21, col149=3.15)
  "5-50억"     → 2.53 % 기초액 3,300,000원 (row29, col149=2.53, col157=3300)
  "50억이상"   → 2.60 % 기초액 없음   (row37, col149=2.6; 800억미만 sub-band)
"""
import pytest
from src.lookup import sanjae_rate, gonggu_rate, fixed_rate, sanan_rate


# ── 산재보험료 (순수 상수) ───────────────────────────────────────
def test_sanjae_hanjeon():
    assert sanjae_rate("한전") == pytest.approx(0.03656)


def test_sanjae_jodalcheong():
    assert sanjae_rate("조달청") == pytest.approx(0.03626)


def test_sanjae_invalid():
    with pytest.raises(ValueError):
        sanjae_rate("무효")


# ── 공구손료 (순수 상수) ─────────────────────────────────────────
def test_gonggu():
    assert gonggu_rate() == pytest.approx(0.03)


# ── 4대보험·노인장기요양 고정요율 (파일 필요) ────────────────────
def test_fixed_geongang(tomok_path):
    """건강보험료 3.595% — probe: AU51 '(직노) x 3.595'"""
    assert fixed_rate(tomok_path, "건강보험료") == pytest.approx(0.03595, abs=1e-5)


def test_fixed_noin(tomok_path):
    """노인장기요양보험료 13.14% — probe: BD51 '(건강보험료) x 13.14'"""
    assert fixed_rate(tomok_path, "노인장기요양보험료") == pytest.approx(0.1314, abs=1e-4)


def test_fixed_yeonkeum(tomok_path):
    """연금보험료 4.75% — probe: BR51 '(직노) x 4.75'"""
    assert fixed_rate(tomok_path, "연금보험료") == pytest.approx(0.0475, abs=1e-5)


def test_fixed_toejik(tomok_path):
    """퇴직공제부금비 2.3% — probe: AU110 '(직노) x 2.3'"""
    assert fixed_rate(tomok_path, "퇴직공제부금비") == pytest.approx(0.023, abs=1e-5)


def test_fixed_goyong(tomok_path):
    """고용보험료 1.01% — probe: AH79=1.01 (7등급)"""
    assert fixed_rate(tomok_path, "고용보험료") == pytest.approx(0.0101, abs=1e-5)


def test_fixed_invalid_item(tomok_path):
    with pytest.raises((ValueError, KeyError)):
        fixed_rate(tomok_path, "존재하지않는항목")


# ── 산안비 구간별 요율 — canonical 4구간 (파일 필요) ────────────

def test_sanan_canonical_2cheonman_miman(tomok_path):
    """2천만미만: 파일에 행 없음 → rate 0.0, 기초액 None.

    산업안전보건관리비는 총 공사금액 2천만 원 이상 건설공사부터 적용되므로
    2천만 미만 구간은 요율 0으로 처리한다.
    """
    r = sanan_rate(tomok_path, "2천만미만")
    assert r["rate"] == pytest.approx(0.0)
    assert r["기초액"] is None


def test_sanan_canonical_5eok_miman(tomok_path):
    """5억미만 / 토목공사 3.15% — probe: row21 col149=3.15, 기초액 없음"""
    r = sanan_rate(tomok_path, "5억미만")
    assert r["rate"] == pytest.approx(0.0315, abs=1e-4)
    assert r["기초액"] is None


def test_sanan_canonical_5_50eok(tomok_path):
    """5-50억 (canonical) / 토목공사 2.53% + 기초액 3,300,000원.

    파일: "5억 ~ \\n50억 미만" (row29 col149=2.53, col157=3300).
    이 테스트가 fix 전에는 LookupError로 실패하던 케이스다.
    """
    r = sanan_rate(tomok_path, "5-50억")
    assert r["rate"] == pytest.approx(0.0253, abs=1e-4)
    assert r["기초액"] == 3_300_000


def test_sanan_canonical_50eok_isang(tomok_path):
    """50억이상 / 토목공사 2.60% — probe: row37 col149=2.6, 기초액 없음 (800억미만 sub-band)."""
    r = sanan_rate(tomok_path, "50억이상")
    assert r["rate"] == pytest.approx(0.026, abs=1e-4)
    assert r["기초액"] is None


def test_sanan_5_50eok_legacy_filetext(tomok_path):
    """5억~50억미만 (파일 정규화 텍스트 직접 전달) — 하위호환 확인."""
    r = sanan_rate(tomok_path, "5억~50억미만")
    assert r["rate"] == pytest.approx(0.0253, abs=1e-4)
    assert r["기초액"] == 3_300_000
