"""엔드투엔드: generate() 파이프라인 검증.

파이프라인: params -> compute_rates -> build_output -> 결과 xlsx.
"""
import openpyxl
import pytest

from src.main import generate
from src.lookup import compute_rates

# 12개 항목 순서 (builder.py CELL_MAP 계약: I{3 + i*2})
# 간접노무비 = ITEMS[0] → I3 → row 3, col 9
_GANJEOP_NOMU_CELL = (3, 9)  # (row, col) 1-based


@pytest.fixture
def base_params(tomok_path):
    return dict(
        jikjeop_cost=500_000_000,
        days=200,
        kind="토목",
        contract="경쟁",
        sanjae_basis="한전",
        sanan_target=400_000_000,  # 5억미만 구간 — 800억 issue 무관
        jebiyul_path=tomok_path,
    )


def test_generate_tomok(base_params, tmp_path):
    """generate()가 결과 xlsx를 생성하고 산재보험료(한전) 값을 포함한다."""
    out = generate(base_params, str(tmp_path / "out.xlsx"))
    ws = openpyxl.load_workbook(out)["적용근거"]
    vals = [
        ws.cell(r, c).value
        for r in range(1, ws.max_row + 1)
        for c in range(1, ws.max_column + 1)
    ]
    assert 0.03656 in vals  # 산재 한전


def test_generate_pipeline_ganjeop_nomu_cell(base_params, tmp_path):
    """간접노무비 rate가 compute_rates와 일치하며 올바른 셀(I3)에 기입된다.

    파이프라인 검증: params → compute_rates → build_output → 셀 값 확인.
    단순히 '어떤 값이 있다'가 아니라 '올바른 값이 올바른 셀에' 있음을 검증한다.
    """
    # 1. compute_rates 결과를 직접 계산 (참값)
    expected_rate = compute_rates(base_params["jebiyul_path"], base_params)["간접노무비"]

    # 2. generate() 로 결과 파일 생성
    out = generate(base_params, str(tmp_path / "out2.xlsx"))

    # 3. 간접노무비 셀(I3 = row 3, col 9) 값 검증
    ws = openpyxl.load_workbook(out)["적용근거"]
    row, col = _GANJEOP_NOMU_CELL
    actual = ws.cell(row, col).value

    assert actual == pytest.approx(expected_rate, rel=1e-6), (
        f"간접노무비 셀 I3 값 불일치: actual={actual!r}, expected={expected_rate!r}"
    )


def test_generate_does_not_raise(base_params, tmp_path):
    """generate()가 예외 없이 완료된다."""
    out = generate(base_params, str(tmp_path / "out3.xlsx"))
    assert out.endswith(".xlsx")
