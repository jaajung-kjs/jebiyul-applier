"""고정요율·산재·산안비 lookup 단위 테스트.

실제 파일에서 탐색한 값(probe 결과):
  건강보험료    3.595 %  (AU51: '(직노) x 3.595')
  연금보험료    4.75  %  (BR51: '(직노) x 4.75')
  퇴직공제부금비 2.3   %  (AU110: '(직노) x 2.3')
  노인장기요양   13.14 %  (BD51: '(건강보험료) x 13.14')
  고용보험료    1.01  %  (AH79: 1.01  — 7등급)

산안비 (토목공사 기준, ES/FA 열):
  5억 미만      3.15 %  기초액 없음         (row 21, ES21=3.15)
  5억~50억 미만 2.53 % + 기초액 3,300,000원 (row 29, ES29=2.53, FA29=3300)
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


# ── 산안비 구간별 요율 (파일 필요) ──────────────────────────────
def test_sanan_5eok_miman(tomok_path):
    """5억 미만 / 토목공사 3.15% — probe: ES21=3.15, FA None"""
    r = sanan_rate(tomok_path, "5억미만")
    assert r["rate"] == pytest.approx(0.0315, abs=1e-4)
    assert r["기초액"] is None


def test_sanan_5_50eok(tomok_path):
    """5억~50억 미만 / 토목공사 2.53% + 기초액 3,300,000원 — probe: ES29=2.53, FA29=3300"""
    r = sanan_rate(tomok_path, "5억~50억미만")
    assert r["rate"] == pytest.approx(0.0253, abs=1e-4)
    assert r["기초액"] == 3_300_000
