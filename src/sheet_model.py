"""입력+적용율 → 적용근거 시트의 순수 데이터 블록 리스트. openpyxl 의존 없음.

입력(공사종류·규모·기간·계약방법)이 구조를 결정하는 유일한 곳이다. 렌더러는 이
블록 리스트를 위→아래로 그릴 뿐 분기 판단을 하지 않는다.
"""
from dataclasses import dataclass, field

NOTE = "[조달청 공사 원가계산 제비율 변경 '24.6.27. 적용]"


@dataclass
class Title:
    text: str


@dataclass
class SectionHeader:
    text: str
    note: str | None = None


@dataclass
class SubHeader:
    text: str


@dataclass
class NoteLines:
    lines: list


@dataclass
class AppliedRate:
    label: str
    value: float
    fmt: str = "0.0%"


@dataclass
class BandRow:
    cells: list
    highlight: bool = False


@dataclass
class BandTable:
    """Task 4~5에서 렌더 규칙과 함께 채운다."""
    kind: str                 # 'gibon' | 'sanan' | 'ilban' | 'iyun'
    headers: list = field(default_factory=list)
    rows: list = field(default_factory=list)


def pct(rate: float) -> str:
    """0.03626 → '3.626' (불필요한 0 제거)."""
    return f"{rate * 100:g}"


def build(params: dict, rates: dict, jebiyul_path: str | None = None) -> list:
    """적용근거 블록 리스트를 만든다. (Task 2: 표 제외 골격)"""
    b: list = []
    b.append(Title("8. 공사비 산출 적용근거"))

    # 1. 간접노무비 — 표는 Task 5에서. 지금은 섹션/근거/적용율만.
    b.append(SectionHeader("1. 간접노무비", NOTE))
    b.append(NoteLines([
        "    ☞ 공사규모별 적용기준 : 재료비 + 직접노무비 + 경비",
        "    ☞ 계상금액 : 직접노무비 × 적용율",
    ]))
    b.append(AppliedRate("    ☞ 적 용 율 :  ", rates["간접노무비"]))

    # 2. 경비
    b.append(SectionHeader("2. 경    비"))
    # 가. 공구손료
    b.append(SubHeader(" 가. 공구손료"))
    b.append(NoteLines(["    ☞ 계상금액 : 직접노무비 × 3%"]))
    b.append(AppliedRate("    ☞ 적 용 율 :  ", rates["공구손료"], fmt="0.0%"))
    # 나. 산재
    b.append(SubHeader(" 나. 산업재해보상보험료"))
    b.append(NoteLines([
        "    ☞ 대상공사 : 모든 건설공사(다만, 총공사금액[(도급금액+관급재료)에서 부가세 제외] 2천만원 미만의",
        "                  건설공사를 건설업자가 아닌 자가 시공시 적용제외",
        f"    ☞ 계상금액 : 노무비 × {pct(rates['산재보험료'])}%",
    ]))
    b.append(AppliedRate("    ☞ 적 용 율 :  ", rates["산재보험료"], fmt="0.000%"))
    # 다. 고용
    b.append(SubHeader(" 다. 고용보험료"))
    b.append(NoteLines([
        "    ☞ 대상공사 : 모든 건설공사(다만, 총공사금액[(도급금액+관급재료)에서 부가세 제외] 2천만원 미만의",
        "                  건설공사를 건설업자가 아닌 자가 시공시 적용제외",
        f"    ☞ 계상금액 : 노무비 × 7등급 보험요율({pct(rates['고용보험료'])}%)",
    ]))
    b.append(AppliedRate("    ☞ 적 용 율 :  ", rates["고용보험료"], fmt="0.00%"))
    # 라. 건강
    b.append(SubHeader(" 라. 국민건강보험료(실적정산)"))
    b.append(NoteLines([
        "    ☞ 대상공사 : 공사기간이 1개월(30일)이상인 공사",
        f"    ☞ 계상금액 : 직접노무비 × 보험요율({pct(rates['건강보험료'])}%)",
    ]))
    b.append(AppliedRate("    ☞ 적 용 율 :  ", rates["건강보험료"], fmt="0.000%"))
    # 마. 연금
    b.append(SubHeader(" 마. 국민연금보험료(실적정산)"))
    b.append(NoteLines([
        "    ☞ 대상공사 : 공사기간이 1개월(30일)이상인 공사",
        f"    ☞ 계상금액 : 직접노무비 × 보험요율({pct(rates['연금보험료'])}%)",
    ]))
    b.append(AppliedRate("    ☞ 적 용 율 :  ", rates["연금보험료"], fmt="0.00%"))
    # 바. 퇴직
    b.append(SubHeader(" 바. 퇴직공제가입금(실적정산)"))
    b.append(NoteLines([
        "    ☞ 대상공사 : 추정금액이 1억원 이상인 공사",
        f"    ☞ 계상금액 : 직접노무비 × 보험요율({pct(rates['퇴직공제부금비'])}%)",
    ]))
    b.append(AppliedRate("    ☞ 적 용 율 :  ", rates["퇴직공제부금비"], fmt="0.0%"))
    # 사. 노인
    b.append(SubHeader(" 사. 노인장기요양보험료  "))
    b.append(NoteLines([
        "    ☞ 대상공사 : 공사기간이 1개월(30일)이상인 공사",
        f"    ☞ 국민건강보험의 보험료 × {pct(rates['노인장기요양보험료'])}%",
    ]))
    b.append(AppliedRate("", rates["노인장기요양보험료"], fmt="0.0%"))
    # 아. 산안비 — 표는 Task 4.
    b.append(SubHeader(" 아. 산업안전보건관리비"))
    b.append(NoteLines([
        "    ☞ 공사규모(대상액)별 적용기준 : 재료비(사급재료비 포함) + 직접노무비",
        "    ☞ 계상금액 : [지입재료비+직접노무비] × 적용율 × 1.2 와",
        "                    [사급재료비+지입재료비+직접노무비] × 적용율 중 적은 금액",
    ]))
    b.append(AppliedRate("    ☞ 적 용 율 :  (사급재료비 제외시)", rates["산업안전보건관리비"], fmt="0.000%"))
    b.append(AppliedRate("    ☞ 적 용 율 :  (사급재료비 포함시)", rates["산업안전보건관리비"], fmt="0.000%"))
    # 자. 기타경비 — 표는 Task 5.
    b.append(SubHeader(" 자. 기타 경비"))
    b.append(NoteLines([
        "    ☞ 공사규모별 적용기준 : 도급재료비 + 노무비 + 경비",
        "    ☞ 계상금액 : (도급재료비 + 노무비) × 적용율",
    ]))
    b.append(AppliedRate("    ☞ 적 용 율 :  ", rates["기타경비"], fmt="0.0%"))

    # 3. 일반관리비 — 표는 Task 4.
    b.append(SectionHeader("3. 일반관리비", NOTE))
    b.append(NoteLines(["    ☞ 공사규모별 적용기준 : 지입재료비 + 노무비 + 도급분경비"]))
    b.append(AppliedRate("    ☞ 적 용 율 :  ", rates["일반관리비"], fmt="0.0%"))

    # 4. 이윤 — 표는 Task 4.
    b.append(SectionHeader("4. 이   윤", NOTE))
    b.append(NoteLines(["    ☞ 공사규모별 적용기준 : 지입재료비 + 노무비 + 도급분경비"]))
    b.append(AppliedRate("    ☞ 적 용 율 :  (일반)", rates["이윤"], fmt="0.00%"))
    b.append(AppliedRate("    ☞ 적 용 율 :  (수의)", rates["이윤"], fmt="0.00%"))

    return b
