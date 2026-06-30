"""제비율 xlsx에서 적용율을 앵커 기반으로 읽는다.

공개 함수:
  cell_text_search(ws, label) -> (row, col)
  find_data_sheet(wb) -> Worksheet
  table_rate(path, item, kind, size, duration, contract=None) -> float
"""
import openpyxl
from src import mapping


def cell_text_search(ws, label):
    """시트에서 label로 시작하는 셀을 찾아 (row, col)을 반환한다.

    검색 기준: 셀 값의 앞뒤 공백·개행을 정규화 후 label로 시작하는 첫 번째 셀.
    없으면 LookupError.
    """
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.strip().replace("\n", " ").startswith(label):
                return cell.row, cell.column
    raise LookupError(f"앵커 라벨을 찾지 못함: {label!r} (시트 {ws.title!r})")


def find_data_sheet(wb):
    """워크북에서 '[간접노무비]' 앵커가 있는 데이터 시트를 반환한다.

    없으면 LookupError.
    """
    for ws in wb.worksheets:
        try:
            cell_text_search(ws, "[간접노무비]")
            return ws
        except LookupError:
            continue
    raise LookupError("'[간접노무비]' 앵커가 있는 데이터 시트를 찾지 못함 — 제비율 양식 불일치")


def _col_for_item_kind(ws, item, src_kind):
    """항목·종류에 맞는 데이터 열 번호(1-base)를 반환한다.

    간접노무비·기타경비는 종류별로 열이 다르므로 DATA_COL 직접 참조.
    일반관리비·이윤(및 DATA_COL에 없는 건축 계열)은 해당 항목 앵커 열 = 데이터 열.
    """
    key = (item, src_kind)
    if key in mapping.DATA_COL:
        return mapping.DATA_COL[key]
    # 앵커 기반 열 결정:
    #   - 일반관리비/이윤: 앵커가 있는 열이 주 데이터 열(토목·건축 파일 공통)
    #   - 건축 파일의 '건축' 종류(간접노무비·기타경비): 앵커 열 = col 29/38 = 토목 열과 동일
    anchor_label = mapping.ANCHOR_LABEL[item]
    _, anc_col = cell_text_search(ws, anchor_label)
    return anc_col


def table_rate(path, item, kind, size, duration, contract=None):
    """제비율표에서 적용율(소수)을 반환한다.

    Parameters
    ----------
    path : str
        제비율 xlsx 파일 경로.
    item : str
        항목명 — '간접노무비', '기타경비', '일반관리비', '이윤' 중 하나.
    kind : str
        공사종류 — '건축', '토목', '조경', '산업설비', '전기통신소방전문' 등.
    size : str
        공사규모 밴드 — params.size_band() 반환값.
    duration : str
        공사기간 밴드 — params.duration_band() 반환값.
        일반관리비·이윤은 기간 무관이므로 이 인수는 무시됨.
    contract : str | None
        이윤 계약방법 ('경쟁' 또는 '수의'). 현재 조달청 파일은 단일 요율 열만
        제공하므로 이 인수는 무시된다. Task 6 인터페이스 호환용으로 포함.

    Returns
    -------
    float
        소수 적용율 (예: 19.1% → 0.191).

    Raises
    ------
    LookupError
        앵커 미발견 또는 해당 셀이 비어 있을 때.
    """
    src_kind = mapping.rate_source_kind(item, kind)

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = find_data_sheet(wb)

    col = _col_for_item_kind(ws, item, src_kind)

    if item in ("일반관리비", "이윤"):
        # 기간 무관; 공사원가 기준 규모 행 (SIZE_BASE_ROW_ILBAN) 사용
        row = mapping.SIZE_BASE_ROW_ILBAN[size]
    else:
        row = mapping.SIZE_BASE_ROW[size] + mapping.DURATION_OFFSET[duration]

    raw = ws.cell(row=row, column=col).value
    if raw is None:
        raise LookupError(
            f"{item}/{kind}/{size}/{duration} 위치(row={row}, col={col}) 값이 비어 있음"
        )
    return float(raw) / 100.0
