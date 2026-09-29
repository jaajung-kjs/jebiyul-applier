"""선택 직종 + NomuReport → 노무임 블록 리스트. 좌표 없음, 순서만(렌더러가 좌표 부여).

입력(선택 직종)이 그룹 구성·행 수·번호를 결정하는 유일한 곳.
"""
from dataclasses import dataclass

from src.hwp_reader import BUMUN_ORDER

_ROMAN = ("Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ")


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
