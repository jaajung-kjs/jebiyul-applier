from src.hwp_reader import read_grids


def _nomu_grids(grids):
    """헤더에 날짜(YYYY.M.D)와 코드행을 가진 노임표만."""
    import re
    date = re.compile(r"\d{4}\.\d{1,2}\.\d{1,2}")
    out = []
    for g in grids:
        joined = " ".join(g.values())
        if date.search(joined) and "통신설비공" in joined:
            out.append(g)
    return out


def test_read_grids_recovers_tongsin_row(hwp_path):
    grids = read_grids(hwp_path)
    assert len(grids) > 10  # 문서에 표가 여럿
    hit = _nomu_grids(grids)
    assert hit, "통신설비공이 든 노임표를 찾지 못함"
    g = hit[0]
    # 통신설비공 행 전체 복원
    pos = next((r, c) for (r, c), v in g.items() if v.strip() == "통신설비공")
    r, _ = pos
    maxc = max(c for (_r, c) in g if _r == r)
    row = ["".join(g.get((r, c), "")).strip() for c in range(maxc + 1)]
    assert row[0] == "1087"
    assert row[1] == "통신설비공"
    assert row[2].replace(",", "") == "315528"
    assert row[5].replace(",", "") == "305050"
