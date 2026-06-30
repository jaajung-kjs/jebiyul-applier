"""compute_rates 통합 테스트.

검증 기준:
  - 12개 키 집합 완전성
  - 산재보험료: 한전 기준 0.03656
  - 공구손료: 고정 0.03
  - 이윤 계약방법 분기: 수의(50-300억) → 0.10, 경쟁(50-300억) → 0.12 (probe BS31=12)
"""
import pytest
from src.lookup import compute_rates

EXPECTED_KEYS = {
    "간접노무비", "공구손료", "산재보험료", "고용보험료",
    "건강보험료", "연금보험료", "퇴직공제부금비", "노인장기요양보험료",
    "산업안전보건관리비", "기타경비", "일반관리비", "이윤",
}

# 공통 파라미터 (키/직접공사비: 5억 → 10억미만 밴드, sanan 4억 → 5억미만 밴드)
_BASE_PARAMS = dict(
    jikjeop_cost=500_000_000,
    days=200,
    kind="토목",
    contract="경쟁",
    sanjae_basis="한전",
    sanan_target=400_000_000,
)

# 이윤 수의/경쟁 비교용 파라미터 (직접공사비 100억 → 50-300억 밴드)
_IYUN_PARAMS_BASE = dict(
    jikjeop_cost=10_000_000_000,
    days=200,
    kind="토목",
    sanjae_basis="조달청",
    sanan_target=400_000_000,
)


def test_compute_tomok_keys(tomok_path):
    """반환 dict가 정확히 12개 키를 포함한다."""
    rates = compute_rates(tomok_path, _BASE_PARAMS)
    assert set(rates) == EXPECTED_KEYS


def test_compute_sanjae_uses_basis(tomok_path):
    """산재보험료는 sanjae_basis='한전' 기준값 0.03656을 반환한다."""
    rates = compute_rates(tomok_path, _BASE_PARAMS)
    assert rates["산재보험료"] == pytest.approx(0.03656)


def test_compute_gonggu_fixed(tomok_path):
    """공구손료는 항상 고정값 0.03이다."""
    params = dict(_BASE_PARAMS, sanjae_basis="조달청")
    rates = compute_rates(tomok_path, params)
    assert rates["공구손료"] == pytest.approx(0.03)


def test_compute_iyun_suui(tomok_path):
    """수의계약 + 50-300억 → 이윤 조달청 정책값 0.10."""
    params = dict(_IYUN_PARAMS_BASE, contract="수의")
    rates = compute_rates(tomok_path, params)
    assert rates["이윤"] == pytest.approx(0.10)


def test_compute_iyun_gyeongjaeng(tomok_path):
    """경쟁계약 + 50-300억 → 이윤 파일값 0.12 (probe: BS31=12)."""
    params = dict(_IYUN_PARAMS_BASE, contract="경쟁")
    rates = compute_rates(tomok_path, params)
    assert rates["이윤"] == pytest.approx(0.12)
