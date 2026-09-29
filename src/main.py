"""엔트리포인트: GUI ↔ 계산 ↔ 출력 배선."""
import os

from src.lookup import compute_rates
from src.builder import build_output


def generate(params: dict, out_path: str) -> str:
    """파라미터를 받아 적용근거 xlsx를 생성하고 경로를 반환한다.

    Parameters
    ----------
    params : dict
        compute_rates 가 요구하는 키-값 집합.
        jebiyul_path, jikjeop_cost, days, kind, contract,
        sanjae_basis, sanan_target.
        선택: hwp_path, selected_nomu — 둘 다 있으면 '시중노무임'
        시트가 추가된다(없으면 기존과 동일한 단일 시트).
    out_path : str
        저장할 결과 파일 경로.

    Returns
    -------
    str
        저장된 파일 경로 (out_path 와 동일).
    """
    rates = compute_rates(params["jebiyul_path"], params)
    return build_output(rates, out_path, params=params,
                        jebiyul_path=params["jebiyul_path"],
                        hwp_path=params.get("hwp_path"),
                        selected_nomu=params.get("selected_nomu"))


def _default_out() -> str:
    """기본 출력 경로: 현재 작업 디렉토리의 적용근거_결과.xlsx."""
    return os.path.join(os.getcwd(), "적용근거_결과.xlsx")


def main() -> None:
    """GUI를 실행하고 submit 시 generate를 호출한다."""
    from src.gui import run_app  # import here so module is importable without display
    run_app(lambda params: generate(params, _default_out()))


if __name__ == "__main__":
    main()
