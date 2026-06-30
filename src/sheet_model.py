"""입력+적용율 → 적용근거 시트의 순수 데이터 블록 리스트. openpyxl 의존 없음.

입력(공사종류·규모·기간·계약방법)이 구조를 결정하는 유일한 곳이다. 렌더러는 이
블록 리스트를 위→아래로 그릴 뿐 분기 판단을 하지 않는다.
"""
from dataclasses import dataclass, field

from src import lookup
from src import params as P

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
    """구간표.

    headers: [(text, col_start, col_end), ...] 단일 헤더행.
    rows:    [BandRow, ...]. BandRow.cells = [(value, c0, c1, fmt, rowspan?), ...].
    criterion: 비었으면 없음. 있으면 렌더러가 G:J(7~10)를 데이터 행 전체에 세로 병합해
               이 텍스트를 '적용기준' 블록으로 그린다(PIU의 G7:J11 같은 세로 병합).
    """
    kind: str                 # 'gibon' | 'sanan' | 'ilban' | 'iyun'
    headers: list = field(default_factory=list)
    rows: list = field(default_factory=list)
    criterion: str = ""


def pct(rate: float) -> str:
    """0.03626 → '3.626' (불필요한 0 제거)."""
    return f"{rate * 100:g}"


# ── 고정 구간표(규모만으로 결정, 행 수 불변) ──────────────────────────────

_SANAN_BANDS = [("2천만원 미만", "2천만미만"), ("5억원 미만", "5억미만"),
                ("5 - 50억원 미만", "5-50억"), ("50억원 이상", "50억이상")]
_SANAN_CRIT = ["[사급재료비+지입재료비+직접노무비]×",
               "적용율과   [지입재료비+직접노무비]×",
               "적용율 × 1.2 중 적은것을 적용한다", ""]

_ILBAN = [("5억원 미만", 0.06), ("5억원 이상 ~ 30억원 미만", 0.055),
          ("30억원 이상 ~ 100억원 미만", 0.05), ("100억원 이상", 0.045)]

_IYUN_COMP = [("50억원 미만", "10억미만"), ("50억원 이상 ~ 300억원 미만", "50-300억"),
              ("300억원 이상 ~ 1000억원 미만", "300-1000억"), ("1000억원 이상", "1000억이상")]
_IYUN_SUUI = [("1000억원 미만", "50-300억"), ("1000억원 이상", "1000억이상")]


def _sanan_table(path, target):
    applied = P.sanan_band(target)
    rows = []
    for (label, band), crit in zip(_SANAN_BANDS, _SANAN_CRIT):
        info = lookup.sanan_rate(path, band)
        if info.get("기초액"):
            cell_rate = (f"{pct(info['rate'])}%+{info['기초액'] / 1000:,.0f}천원", 4, 6, None)
        else:
            cell_rate = (info["rate"], 4, 6, "0.00%")
        rows.append(BandRow([(label, 2, 3, None), cell_rate, (crit, 7, 10, None)],
                            highlight=(band == applied)))
    return BandTable(
        kind="sanan",
        headers=[("공사규모(대상액)별", 2, 3), ("적    용    율[특수 및 기타 적용]", 4, 6),
                 ("적  용  기  준", 7, 10)],
        rows=rows,
    )


def _ilban_table(jikjeop):
    applied = _ilban_band(jikjeop)
    rows = []
    for i, (label, rate) in enumerate(_ILBAN):
        rows.append(BandRow([(label, 2, 3, None), (rate, 4, 6, "0.0%")],
                            highlight=(i == applied)))
    return BandTable(
        kind="ilban",
        headers=[("공 사 규 모", 2, 3), ("적용율 (전기, 통신, 소방 전문 적용)", 4, 6),
                 ("적  용  기  준", 7, 10)],
        rows=rows,
        criterion="[지입재료비+노무비+도급분경비]×적용율",
    )


def _ilban_band(jikjeop):
    if jikjeop < 5e8:
        return 0
    if jikjeop < 3e9:
        return 1
    if jikjeop < 1e10:
        return 2
    return 3


def _iyun_table(path, kind, jikjeop, contract):
    comp_applied = P.size_band(jikjeop)
    suui_applied = "1000억이상" if jikjeop >= 1e11 else "50-300억"
    rows = []
    for i, (label, band) in enumerate(_IYUN_COMP):
        rate = lookup.table_rate(path, "이윤", kind, band, "183", "경쟁")
        cells = []
        if i == 0:
            cells.append(("경쟁계약", 2, 2, None, len(_IYUN_COMP)))  # 세로 병합
        cells += [(label, 3, 5, None), (rate, 6, 6, "0.0%")]
        rows.append(BandRow(cells, highlight=(contract == "경쟁" and band == comp_applied)))
    for i, (label, band) in enumerate(_IYUN_SUUI):
        rate = lookup.table_rate(path, "이윤", kind, band, "183", "수의")
        cells = []
        if i == 0:
            cells.append(("수의계약", 2, 2, None, len(_IYUN_SUUI)))
        cells += [(label, 3, 5, None), (rate, 6, 6, "0.0%")]
        rows.append(BandRow(cells, highlight=(contract == "수의" and band == suui_applied)))
    return BandTable(
        kind="iyun",
        headers=[("구        분", 2, 2), ("공사규모(추정가격 기준)", 3, 5),
                 ("적용율", 6, 6), ("적  용  기  준", 7, 10)],
        rows=rows,
        criterion="[노무비+경비(기술료,외주가공비 제외)+\n   일반관리비]×적용율",
    )


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
    # 아. 산안비
    b.append(SubHeader(" 아. 산업안전보건관리비"))
    if jebiyul_path:
        b.append(_sanan_table(jebiyul_path, params["sanan_target"]))
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

    # 3. 일반관리비
    b.append(SectionHeader("3. 일반관리비", NOTE))
    if jebiyul_path:
        b.append(_ilban_table(params["jikjeop_cost"]))
    b.append(NoteLines(["    ☞ 공사규모별 적용기준 : 지입재료비 + 노무비 + 도급분경비"]))
    b.append(AppliedRate("    ☞ 적 용 율 :  ", rates["일반관리비"], fmt="0.0%"))

    # 4. 이윤
    b.append(SectionHeader("4. 이   윤", NOTE))
    if jebiyul_path:
        b.append(_iyun_table(jebiyul_path, params["kind"], params["jikjeop_cost"], params["contract"]))
    b.append(NoteLines(["    ☞ 공사규모별 적용기준 : 지입재료비 + 노무비 + 도급분경비"]))
    b.append(AppliedRate("    ☞ 적 용 율 :  (일반)", rates["이윤"], fmt="0.00%"))
    b.append(AppliedRate("    ☞ 적 용 율 :  (수의)", rates["이윤"], fmt="0.00%"))

    return b
