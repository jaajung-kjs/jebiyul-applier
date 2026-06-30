"""기존 적용근거 시트 양식(마스터 템플릿)에 이번 공사의 값을 채워 결과 xlsx를 만든다.

마스터 템플릿(`assets/template_적용근거.xlsx`)은 실제 설계서의 `적용근거` 시트를 그대로
추출한 것이라 ☞ 적용기준·구간표·서식이 전부 보존돼 있다. 빌더는 그 안에서 "값이 바뀌는
칸"만 덮어쓴다:

  1. 각 항목 헤드라인 ☞ 적용율 셀 (I열)
  2. 보험요율이 박혀 있는 설명 문구 (산재/고용/건강/연금/퇴직/노인)
  3. 간접노무비·기타경비·산안비·이윤 구간표 값

일반관리비 구간표(전문 5/30/100억 스케줄)는 제비율 본표와 구간 기준이 달라 자동 갱신하지
않고 표준 참고표로 둔다. 헤드라인 일반관리비 적용율(I83)만 채운다.
"""
import os
import sys
import shutil

import openpyxl

from src import lookup
from src import params as P

# Canonical item list — single source of truth for the runtime package.
ITEMS = [
    "간접노무비",
    "공구손료",
    "산재보험료",
    "고용보험료",
    "건강보험료",
    "연금보험료",
    "퇴직공제부금비",
    "노인장기요양보험료",
    "산업안전보건관리비",
    "기타경비",
    "일반관리비",
    "이윤",
]

# 마스터 템플릿 경로. PyInstaller --onefile 번들에서는 sys._MEIPASS, 아니면 저장소 루트.
_BASE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_ASSET = os.path.join(_BASE, "assets", "template_적용근거.xlsx")

# 항목 → 헤드라인 ☞ 적용율 셀 (마스터 시트 좌표).
HEADLINE_CELL = {
    "간접노무비": "I14",
    "공구손료": "I19",
    "산재보험료": "I24",
    "고용보험료": "I29",
    "건강보험료": "I33",
    "연금보험료": "I37",
    "퇴직공제부금비": "I41",
    "노인장기요양보험료": "I44",
    "산업안전보건관리비": "I54",  # 사급재료비 제외시 (J54='*1.2')
    "기타경비": "I74",
    "일반관리비": "I83",
    "이윤": "I94",  # 일반(경쟁) — 수의는 I95에 별도
}

# 공사규모 밴드 → 사람이 읽는 라벨.
SIZE_LABEL = {
    "10억미만": "10억 미만",
    "10-50억": "10억 ~ 50억 미만",
    "50-300억": "50억 ~ 300억 미만",
    "300-1000억": "300억 ~ 1000억 미만",
    "1000억이상": "1000억 이상",
}

_DURATIONS = ["183", "365", "1095", "1096+"]


def _pct(rate: float) -> str:
    """0.03656 → '3.656' (불필요한 0 제거)."""
    return f"{rate * 100:g}"


def build_output(rates: dict, out_path: str, params: dict | None = None,
                 jebiyul_path: str | None = None, template_path: str = _ASSET) -> str:
    """마스터 템플릿을 복사해 이번 공사의 값을 채운 뒤 저장한다.

    Args:
        rates:         compute_rates 결과(12개 항목 → 적용율).
        out_path:      저장 경로.
        params:        공사 파라미터(구간표·이윤 경쟁/수의 채움용). None이면 헤드라인만 채움.
        jebiyul_path:  제비율 파일 경로(구간표 룩업용). params와 함께 있어야 표를 채운다.
        template_path: 마스터 템플릿 경로.

    Returns:
        저장된 파일 경로.
    """
    shutil.copyfile(template_path, out_path)
    wb = openpyxl.load_workbook(out_path)
    ws = wb["적용근거"]

    # 1) 헤드라인 ☞ 적용율 (수식 → 실제값으로 덮어씀)
    for item, cell in HEADLINE_CELL.items():
        if item in rates:
            ws[cell] = rates[item]
    # 산안비 '사급재료비 포함시' 율도 동일값
    if "산업안전보건관리비" in rates:
        ws["I55"] = rates["산업안전보건관리비"]

    # 2) 설명 문구에 박힌 보험요율 갱신
    if "산재보험료" in rates:
        ws["A23"] = f"☞ 계상금액 : 노무비 × {_pct(rates['산재보험료'])}%"
    if "고용보험료" in rates:
        ws["A28"] = f"☞ 계상금액 : 노무비 × 7등급 보험요율({_pct(rates['고용보험료'])}%)"
    if "건강보험료" in rates:
        ws["A32"] = f"☞ 계상금액 : 직접노무비 × 보험요율({_pct(rates['건강보험료'])}%)"
    if "연금보험료" in rates:
        ws["A36"] = f"☞ 계상금액 : 직접노무비 × 보험요율({_pct(rates['연금보험료'])}%)"
    if "퇴직공제부금비" in rates:
        ws["A40"] = f"☞ 계상금액 : 직접노무비 × 보험요율({_pct(rates['퇴직공제부금비'])}%)"
    if "노인장기요양보험료" in rates:
        ws["A44"] = f"☞ 국민건강보험의 보험료 × {_pct(rates['노인장기요양보험료'])}%"

    # 3) 구간표 — params/제비율 파일이 있을 때만
    if params is not None and jebiyul_path:
        _fill_tables(ws, params, jebiyul_path)

    wb.save(out_path)
    return out_path


def _fill_tables(ws, params: dict, path: str) -> None:
    """간접노무비·기타경비·산안비·이윤 구간표를 제비율 파일 값으로 채운다."""
    kind = params["kind"]
    size = P.size_band(params["jikjeop_cost"])

    # --- 간접노무비표 (행 7~10): 이번 공사규모 밴드의 기간별 율 ---
    ws["B7"] = SIZE_LABEL[size]
    for i, dur in enumerate(_DURATIONS):
        ws[f"E{7 + i}"] = lookup.table_rate(path, "간접노무비", kind, size, dur)
    ws["B11"] = None  # '50억 이상 → 참조' 행은 이번 밴드 표시로 대체되므로 비움
    ws["C11"] = None

    # --- 기타경비표 (행 59~62): 기준·적용율 동일하게 파일값으로 ---
    ws["B59"] = SIZE_LABEL[size]
    for i, dur in enumerate(_DURATIONS):
        v = lookup.table_rate(path, "기타경비", kind, size, dur)
        ws[f"E{59 + i}"] = v
        ws[f"F{59 + i}"] = v
    ws["B63"] = None
    ws["C63"] = None

    # --- 산업안전보건관리비표 (행 47~50): 대상액 구간별 ---
    ws["D47"] = lookup.sanan_rate(path, "2천만미만")["rate"]      # 0.0
    ws["D48"] = lookup.sanan_rate(path, "5억미만")["rate"]
    s5 = lookup.sanan_rate(path, "5-50억")
    if s5.get("기초액"):
        ws["D49"] = f"{_pct(s5['rate'])}%+{s5['기초액'] / 1000:,.0f}천원"
    else:
        ws["D49"] = s5["rate"]
    ws["D50"] = lookup.sanan_rate(path, "50억이상")["rate"]

    # --- 이윤표 (행 87~92): 경쟁(규모별) + 수의 ---
    # 경쟁: 50억미만/50-300/300-1000/1000+
    ws["F87"] = lookup.table_rate(path, "이윤", kind, "10억미만", "183", "경쟁")
    ws["F88"] = lookup.table_rate(path, "이윤", kind, "50-300억", "183", "경쟁")
    ws["F89"] = lookup.table_rate(path, "이윤", kind, "300-1000억", "183", "경쟁")
    ws["F90"] = lookup.table_rate(path, "이윤", kind, "1000억이상", "183", "경쟁")
    # 수의: 1000억 미만/이상
    ws["F91"] = lookup.table_rate(path, "이윤", kind, "50-300억", "183", "수의")
    ws["F92"] = lookup.table_rate(path, "이윤", kind, "1000억이상", "183", "수의")
    # 헤드라인 수의 이윤(I95)도 이번 규모 기준으로
    ws["I95"] = lookup.table_rate(path, "이윤", kind, size, "183", "수의")
