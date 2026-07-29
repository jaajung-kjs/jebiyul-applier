"""선택 직종 + NomuReport → 노무임 블록 리스트. 좌표 없음, 순서만(렌더러가 좌표 부여).

입력(선택 직종)이 그룹 구성·행 수·번호를 결정하는 유일한 곳.
"""
from dataclasses import dataclass

from src.hwp_reader import BUMUN_ORDER

_ROMAN = ("Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ")

# 통신공사 표준세트(프리체크 기본값). 이름은 hwp 표기와 정확히 일치해야 한다
# (test_standard_set_names_exist_in_hwp가 강제).
STANDARD_SET = (
    "보통인부", "특별인부", "비계공", "배관공", "기계설비공", "인력운반공", "화물차운전사",
    "내선전공", "송전전공", "배전전공", "저압케이블전공", "계장공",
    "통신내선공", "통신외선공", "통신설비공", "무선안테나공", "광케이블설치사",
    "H/W시험사", "S/W시험사",
    "지적기사", "지적기능사", "통신관련산업기사", "전기공사기사",
)


@dataclass(frozen=True)
class NomuTitle:
    text: str
    past_count: int = 3


@dataclass(frozen=True)
class NomuHeader:
    past_cols: list
    current_col: str


@dataclass(frozen=True)
class NomuGroup:
    roman: str
    name: str
    past_count: int = 3


@dataclass(frozen=True)
class NomuRow:
    name: str
    code: str
    past_wages: list
    current_wage: object
    delta: object
    note: str


@dataclass(frozen=True)
class NomuAvg:
    value: object
    past_count: int = 3


@dataclass(frozen=True)
class NomuFootnote:
    lines: list


def standard_set():
    return list(STANDARD_SET)


def preselect(all_names, standard=None):
    """(직종명, 체크여부) 쌍을 보고서 순서대로. 표준세트만 True."""
    std = set(standard if standard is not None else STANDARD_SET)
    return [(n, n in std) for n in all_names]


def _delta(rate):
    cur, prev = rate.wages[0], rate.wages[1]
    if cur is None or prev is None or prev == 0:
        return None
    return (cur - prev) / prev


def build(report, selected):
    picked = [report.rates[n] for n in selected if n in report.rates]
    p = len(report.dates) - 1
    blocks = [
        NomuTitle("시중노무임 산출", past_count=p),
        NomuHeader(list(reversed(report.dates[1:])), report.dates[0]),
    ]
    gi = 0
    for bumun in BUMUN_ORDER:
        members = sorted((r for r in picked if r.bumun == bumun), key=lambda r: r.code)
        if not members:
            continue
        blocks.append(NomuGroup(_ROMAN[gi], f"{bumun}직종", past_count=p))
        gi += 1
        deltas = []
        for r in members:
            d = _delta(r)
            deltas.append(d)
            blocks.append(NomuRow(
                name=r.name, code=r.code,
                past_wages=list(reversed(r.wages[1:])),
                current_wage=r.wages[0], delta=d, note=r.marker,
            ))
        valid = [d for d in deltas if d is not None]
        blocks.append(NomuAvg(sum(valid) / len(valid) if valid else None, past_count=p))

    # 선택 직종에 실제로 있는 마커만, hwp 원문 각주 문장으로 하단에 표기
    notes = report.marker_notes or {}
    present = {r.marker for r in picked if r.marker}
    footnote = [notes[m] for m in ("*", "**") if m in present and m in notes]
    if footnote:
        blocks.append(NomuFootnote(footnote))
    return blocks
