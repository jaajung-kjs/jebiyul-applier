"""제비율 xlsx에서 적용율을 앵커 기반으로 읽는다.

공개 함수:
  cell_text_search(ws, label) -> (row, col)
  find_data_sheet(wb) -> Worksheet
  table_rate(path, item, kind, size, duration, contract=None) -> float
  sanjae_rate(basis) -> float
  gonggu_rate() -> float
  fixed_rate(path, item) -> float
  sanan_rate(path, target_band) -> dict
  compute_rates(path, params) -> dict[str, float]
"""
import re
import openpyxl
from src import mapping

# 조달청 수의계약 이윤율 기준 (파일에서 읽지 않고 정책값으로 고정).
# 출처: 조달청 수의계약 기준 (간접공사비 적용기준 고시 외 별도 기준).
SUUI_IYUN_UNDER_1000 = 0.10  # 공사규모 1000억 미만 수의계약 이윤율
SUUI_IYUN_OVER_1000  = 0.09  # 공사규모 1000억 이상 수의계약 이윤율


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


def _merged_value(ws, row, col):
    """셀 값. 병합 영역 내부(좌상단이 아닌) 칸이면 좌상단 값으로 해석한다.

    제비율표는 같은 요율이 걸친 구간을 병합해 두므로, 대표 행이 병합 내부에
    떨어져도 값을 읽을 수 있어야 한다.
    """
    v = ws.cell(row=row, column=col).value
    if v is not None:
        return v
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return ws.cell(row=m.min_row, column=m.min_col).value
    return None


def _ilban_jeonmun_col(ws):
    """일반관리비 '전문·전기·통신·소방·기타' 전용 열을 헤더 텍스트로 찾는다.

    열 번호는 파일마다 다르므로(토목 col64 / 건축 col63) 하드코딩하지 않는다.
    """
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if not isinstance(v, str):
                continue
            t = v.replace(" ", "").replace("\n", "")
            if (t.startswith("전문") and len(t) <= 20
                    and "전기" in t and "통신" in t and "소방" in t):
                return c.column
    raise LookupError(
        "일반관리비 '전문·전기·통신·소방' 열 헤더를 찾지 못함 — 제비율 양식 불일치")


# 공사종류 → 헤더 텍스트. '산업설비(토목)'/'산업설비(건축)'처럼 괄호가 붙는다.
_KIND_HEADER_PREFIX = {
    "토목": "토목", "건축": "건축", "조경": "조경", "산업설비": "산업설비",
}


def _item_block_end(ws, anc_col):
    """항목 블록의 끝 열 — 오른쪽에 있는 다음 항목 앵커 열."""
    right = []
    for label in mapping.ANCHOR_LABEL.values():
        try:
            _, c = cell_text_search(ws, label)
        except LookupError:
            continue
        if c > anc_col:
            right.append(c)
    return min(right) if right else anc_col + 25


def _kind_col_by_header(ws, item, kind, anc_row, anc_col):
    """항목 블록의 헤더행에서 공사종류 열을 찾는다. 없으면 None.

    열 번호는 파일마다 다르므로(건축 파일은 산업설비가 34, 토목 파일은 35)
    하드코딩하지 않고 헤더 텍스트로 해석한다.
    """
    prefix = _KIND_HEADER_PREFIX.get(kind)
    if prefix is None:
        return None
    end_col = _item_block_end(ws, anc_col)
    for r in range(anc_row, anc_row + 12):
        for c in range(anc_col, end_col):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                t = v.replace(" ", "").replace("\n", "")
                if t == prefix or t.startswith(prefix + "("):
                    return c
    return None


def _col_for_item_kind(ws, item, src_kind, kind=None):
    """항목·종류에 맞는 데이터 열 번호(1-base)를 반환한다.

    - 간접노무비·기타경비: 블록 헤더에서 공사종류 열을 텍스트로 찾는다.
    - 일반관리비·이윤: 앵커 열이 주 데이터 열. 단 전기·통신·소방·전문의
      일반관리비는 전용 열을 쓴다(파일 주석: "일반관리비요율을 제외한 각종
      요율은 토목, 건축 등 관련 공사업종에 따라 적용").
    - 전기·통신·소방·전문의 그 외 항목: 파일 기본 업종(앵커 열)을 따른다.
    """
    if item == "일반관리비" and src_kind == "전기통신소방전문":
        return _ilban_jeonmun_col(ws)

    anc_row, anc_col = cell_text_search(ws, mapping.ANCHOR_LABEL[item])

    if item in ("간접노무비", "기타경비"):
        col = _kind_col_by_header(ws, item, src_kind, anc_row, anc_col)
        if col is not None:
            return col
        if (kind or src_kind) == "전기통신소방전문":
            return anc_col      # 파일 기본 업종 열
        raise LookupError(
            f"'{src_kind}' 공사종류의 {item} 열을 이 제비율 파일에서 찾지 못함 — "
            f"해당 공사종류에 맞는 제비율 파일인지 확인하세요")

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
        계약방법 ('경쟁' 또는 '수의').
        이윤(이윤)에 한해 반영됨:
          - '수의' → 조달청 수의계약 기준값 (SUUI_IYUN_UNDER_1000 / SUUI_IYUN_OVER_1000)
            을 하드코딩으로 반환한다. 파일값을 사용하지 않는다.
          - None 또는 '경쟁' → 파일에서 읽은 경쟁계약 이윤율을 반환한다.
        다른 항목(간접노무비·기타경비·일반관리비)에서는 무시된다.

    Returns
    -------
    float
        소수 적용율 (예: 19.1% → 0.191).

    Raises
    ------
    LookupError
        앵커 미발견 또는 해당 셀이 비어 있을 때.
    """
    # 수의계약 이윤율: 파일 대신 조달청 정책값을 바로 반환.
    if item == "이윤" and contract == "수의":
        if size == "1000억이상":
            return SUUI_IYUN_OVER_1000
        return SUUI_IYUN_UNDER_1000

    src_kind = mapping.rate_source_kind(item, kind)

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = find_data_sheet(wb)

    col = _col_for_item_kind(ws, item, src_kind, kind)

    if item in ("일반관리비", "이윤"):
        # 기간 무관; 공사원가 기준 규모 행 (SIZE_BASE_ROW_ILBAN) 사용
        row = mapping.SIZE_BASE_ROW_ILBAN[size]
    else:
        row = mapping.SIZE_BASE_ROW[size] + mapping.DURATION_OFFSET[duration]

    raw = _merged_value(ws, row, col)
    if raw is None:
        raise LookupError(
            f"{item}/{kind}/{size}/{duration} 위치(row={row}, col={col}) 값이 비어 있음"
        )
    return float(raw) / 100.0


# ─────────────────────────────────────────────────────────────────────────────
# 산재보험료 / 공구손료 (순수 상수)
# ─────────────────────────────────────────────────────────────────────────────

# 경로 없이 호출되는 레거시 경로용 폴백값.
# 한전은 제비율 표의 기준이 아니라 주석에만 있는 값이라 정책 상수로 유지한다.
_SANJAE = {"한전": 0.03656, "조달청": 0.03626}


def sanjae_rate(basis: str, path: str | None = None) -> float:
    """산재보험료율 (소수) — basis ∈ {'한전', '조달청'}.

    조달청 기준은 제비율 파일의 [산재보험료] 앵커 주변 '(노) x <요율>' 셀에서
    읽는다(고용·건강·연금 등 다른 고정요율과 동일한 방식). 파일 경로가 없거나
    파싱에 실패하면 폴백 상수를 쓴다.
    한전 기준은 표 기준이 아니라 파일 주석의 안내값이므로 정책 상수를 유지한다.
    """
    if basis not in _SANJAE:
        raise ValueError(f"알 수 없는 산재 기준: {basis!r}. 유효값: {list(_SANJAE)}")
    if basis == "조달청" and path:
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = find_data_sheet(wb)
            anc_row, anc_col = cell_text_search(ws, "[산재보험료]")
            return _formula_rate_near(ws, anc_row, anc_col) / 100.0
        except LookupError:
            pass          # 앵커/패턴 없으면 폴백 상수
    return _SANJAE[basis]


def gonggu_rate() -> float:
    """공구손료 고정율 3 % → 0.03."""
    return 0.03


# ─────────────────────────────────────────────────────────────────────────────
# 4대보험·노인장기요양 고정요율 (제비율 파일 참조)
# ─────────────────────────────────────────────────────────────────────────────

_FIXED_LABEL = {
    "고용보험료":       "[고용보험료]",
    "건강보험료":       "[건강보험료]",
    "연금보험료":       "[연금보험료]",
    "퇴직공제부금비":   "[퇴직공제부금비]",
    "노인장기요양보험료": "[노인장기요양보험료]",
}

# 파일에 요율이 "(직노) x 3.595" 같은 텍스트 셀로 저장되어 있다.
# 'x <숫자>' 패턴으로 요율을 추출한다.
_RE_X_NUM = re.compile(r"\bx\s+([\d]+(?:\.[\d]+)?)")


def _formula_rate_near(ws, anchor_row: int, anchor_col: int, span: int = 8) -> float:
    """앵커 주변에서 'x <숫자>' 패턴의 수식 텍스트 셀을 찾아 요율을 파싱한다.

    탐색 범위: anchor_row ~ anchor_row+span, anchor_col ~ anchor_col+span.
    """
    for r in range(anchor_row, anchor_row + span):
        for c in range(anchor_col, anchor_col + span):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                m = _RE_X_NUM.search(v)
                if m:
                    return float(m.group(1))
    raise LookupError(
        f"앵커({anchor_row},{anchor_col}) 인접에서 요율 수식을 찾지 못함 (span={span})"
    )


def _goyong_grade7_rate(ws, anchor_row: int, anchor_col: int) -> float:
    """고용보험료 7등급 요율을 찾아 반환한다.

    앵커 열(anchor_col)에서 아래로 '[7등급]'을 포함한 셀을 찾고,
    앵커 열부터 오른쪽으로 스캔하여 첫 번째 숫자 값을 반환한다.

    요율 셀은 앵커(B열) 오른쪽 AH열 등에 위치하므로,
    앵커 왼쪽 셀(등급 번호·임계값 등)을 실수로 집어가지 않도록
    스캔을 anchor_col 이상 열에서만 시작한다.
    """
    for r in range(anchor_row, anchor_row + 30):
        v = ws.cell(r, anchor_col).value
        if isinstance(v, str) and "[7등급]" in v:
            # 요율은 앵커(B열)의 오른쪽 열에 있음 — anchor_col부터 우측으로만 탐색
            for c in range(anchor_col, anchor_col + 50):
                n = ws.cell(r, c).value
                if isinstance(n, (int, float)):
                    return float(n)
    raise LookupError("고용보험료 [7등급] 요율 셀을 찾지 못함")


def fixed_rate(path: str, item: str) -> float:
    """제비율 파일에서 4대보험·노인장기요양 고정요율을 소수로 반환한다.

    Parameters
    ----------
    path : str
        제비율 xlsx 파일 경로.
    item : str
        항목명 — _FIXED_LABEL 키 중 하나.

    Returns
    -------
    float
        소수 요율 (예: 3.595 % → 0.03595).

    Raises
    ------
    KeyError
        item 이 _FIXED_LABEL 에 없을 때.
    LookupError
        앵커 또는 요율 셀 미발견.
    """
    label = _FIXED_LABEL[item]          # KeyError 로 잘못된 item 알림
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = find_data_sheet(wb)
    anc_row, anc_col = cell_text_search(ws, label)

    if item == "고용보험료":
        raw = _goyong_grade7_rate(ws, anc_row, anc_col)
    else:
        raw = _formula_rate_near(ws, anc_row, anc_col)

    # 파일 값이 퍼센트(예: 3.595)인지 소수(예: 0.03595)인지 정규화
    return raw / 100.0 if raw > 1 else raw


# ─────────────────────────────────────────────────────────────────────────────
# 산업안전보건관리비 구간별 요율 (제비율 파일 참조)
# ─────────────────────────────────────────────────────────────────────────────

def _band_norm(text: str) -> str:
    """대상액 구간 텍스트를 정규화한다 (공백·개행 제거)."""
    return text.strip().replace("\n", "").replace(" ", "")


# params.sanan_band() 정규(canonical) 키 → 파일 내 정규화 텍스트 매핑.
#
# 파일 원문(정규화 전/후):
#   "5억 미만"              → "5억미만"
#   "5억 ~ \n50억 미만"     → "5억~50억미만"
#   "50억 이상"             → "50억이상"
# "2천만미만"은 파일에 별도 행 없음 — 산업안전보건관리비는 총 공사금액 2천만 원 이상
# 건설공사부터 적용되므로 미만 구간은 요율 0으로 처리한다.
_SANAN_CANONICAL_TO_FILE: dict[str, str | None] = {
    "2천만미만": None,           # 파일에 해당 행 없음 → 요율 0 반환
    "5억미만":   "5억미만",      # 파일: "5억 미만"
    "5-50억":    "5억~50억미만", # 파일: "5억 ~ \n50억 미만"
    "50억이상":  "50억이상",     # 파일: "50억 이상"
}


SANAN_KINDS = ("건축공사", "토목공사", "중건설공사", "특수건설공사")


SANAN_EST_THRESHOLD = 80_000_000_000   # 추정금액 800억 (파일의 sub-band 경계)


def sanan_rate(path: str, target_band: str, sanan_kind: str = "토목공사",
               est_cost: int | None = None) -> dict:
    """산안비 대상액 구간·공사종류별 요율과 기초액을 반환한다.

    Parameters
    ----------
    path : str
        제비율 xlsx 파일 경로.
    target_band : str
        대상액 구간 — 공백·개행 없는 형식으로 전달.
        예) '5억미만', '5억~50억미만', '50억이상'

    Returns
    -------
    dict
        {"rate": float (소수), "기초액": int (원) | None}
        기초액은 FA 열 값(천원) × 1000 으로 환산.

    Raises
    ------
    LookupError
        앵커·헤더·구간·토목공사 행을 찾지 못할 때.

    Note
    ----
    "50억이상" 구간은 파일 내에서 두 개의 sub-band로 분리된다:
      - 추정금액 800억 미만 (rate ≈ 2.6%)
      - 추정금액 800억 이상 (rate ≈ 2.73%)
    est_cost가 800억 이상이면 뒤 블록을, 아니면 앞 블록을 사용한다.
    """
    if sanan_kind not in SANAN_KINDS:
        raise ValueError(
            f"알 수 없는 산안비 공사종류: {sanan_kind!r}. 유효값: {list(SANAN_KINDS)}")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = find_data_sheet(wb)

    # 1. [산업안전보건관리비] 앵커 위치 확인
    anc_row, anc_col = cell_text_search(ws, "[산업안전보건관리비]")

    # 2. 헤더 행에서 '구분'·'요율'·'기초액' 열 검색
    #    앵커 열(anc_col) 오른쪽에서만 탐색 (같은 시트 다른 섹션 혼동 방지)
    hdr_row = type_col = rate_col = base_col = None
    for r in range(anc_row + 1, anc_row + 20):
        t_col = r_col = b_col = None
        for c in range(anc_col + 1, anc_col + 80):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                if "구분" in v and t_col is None:
                    t_col = c
                elif "요율" in v and t_col is not None and r_col is None:
                    r_col = c
                elif "기초액" in v and t_col is not None and b_col is None:
                    b_col = c
        if t_col and r_col:
            hdr_row = r
            type_col = t_col
            rate_col = r_col
            base_col = b_col
            break

    if hdr_row is None:
        raise LookupError("[산업안전보건관리비] 헤더 행(구분/요율)을 찾지 못함")

    # 3. 대상액 구간 행 탐색 (anc_col 열에서 정규화 비교)
    # NOTE: "50억이상"은 800억미만 sub-band만 반환 — 800억 이상 공사는 부정확
    #
    # canonical 키(params.sanan_band 반환값) → 파일 내 정규화 텍스트로 변환.
    # "2천만미만"은 파일에 행 없으므로 즉시 요율 0 반환.
    if target_band in _SANAN_CANONICAL_TO_FILE:
        file_key = _SANAN_CANONICAL_TO_FILE[target_band]
        if file_key is None:
            return {"rate": 0.0, "기초액": None}
        norm_target = file_key
    else:
        norm_target = _band_norm(target_band)
    band_row = None
    for r in range(hdr_row + 1, anc_row + 60):
        v = ws.cell(r, anc_col).value
        if isinstance(v, str) and _band_norm(v) == norm_target:
            band_row = r
            break

    if band_row is None:
        raise LookupError(f"산안비 구간 '{target_band}'을 찾지 못함")

    # 3-1. '50억 이상' 구간의 추정금액 800억 sub-band 선택.
    #      파일은 band_row 아래에 '추정금액 800억 미만/이상' 두 블록을 둔다.
    search_from = band_row
    if est_cost is not None and est_cost >= SANAN_EST_THRESHOLD:
        for r in range(band_row, band_row + 20):
            for c in range(anc_col, type_col + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str):
                    t = v.replace(" ", "")
                    if "800억" in t and "이상" in t:
                        search_from = r
                        break
            if search_from != band_row:
                break

    # 4. 해당 구간 내에서 지정한 공사종류 행 탐색
    tomok_row = None
    for r in range(search_from, search_from + 10):
        v = ws.cell(r, type_col).value
        if isinstance(v, str) and v.replace(" ", "") == sanan_kind:
            tomok_row = r
            break

    if tomok_row is None:
        raise LookupError(
            f"산안비 구간 '{target_band}' 내 '{sanan_kind}' 행을 찾지 못함 "
            f"(band_row={band_row})"
        )

    # 5. 요율 읽기 (% → 소수 변환)
    raw_rate = ws.cell(tomok_row, rate_col).value
    if raw_rate is None:
        raise LookupError(f"산안비 요율 셀 비어 있음 (row={tomok_row}, col={rate_col})")
    v = float(raw_rate)
    rate = v / 100.0 if v > 1 else v

    # 6. 기초액 읽기 (천원 → 원 변환)
    base_raw = ws.cell(tomok_row, base_col).value if base_col else None
    base_won = int(float(base_raw) * 1000) if base_raw is not None else None

    return {"rate": rate, "기초액": base_won}


# ─────────────────────────────────────────────────────────────────────────────
# 12개 항목 적용율 일괄 산출
# ─────────────────────────────────────────────────────────────────────────────

def compute_rates(path: str, params: dict) -> dict:
    """파라미터 dict를 받아 12개 항목 적용율(소수)을 한 번에 반환한다.

    Parameters
    ----------
    path : str
        제비율 xlsx 파일 경로.
    params : dict
        jikjeop_cost : int   직접공사비 (원)
        days         : int   공사기간 (일)
        kind         : str   공사종류 ('토목', '건축', '조경', …)
        contract     : str   계약방법 ('경쟁' | '수의')
        sanjae_basis : str   산재보험료 기준 ('한전' | '조달청')
        sanan_target : int   산안비 대상액 (원)

    Returns
    -------
    dict[str, float]
        12개 항목 키에 대한 소수 적용율.
        키: 간접노무비, 공구손료, 산재보험료, 고용보험료, 건강보험료, 연금보험료,
            퇴직공제부금비, 노인장기요양보험료, 산업안전보건관리비, 기타경비, 일반관리비, 이윤.
    """
    from src import params as P

    size = P.size_band(params["jikjeop_cost"])
    dur = P.duration_band(params["days"])
    kind = params["kind"]
    contract = params["contract"]
    sanan = P.sanan_band(params["sanan_target"])

    return {
        "간접노무비":         table_rate(path, "간접노무비",   kind, size, dur),
        "공구손료":           gonggu_rate(),
        "산재보험료":         sanjae_rate(params["sanjae_basis"], path),
        "고용보험료":         fixed_rate(path, "고용보험료"),
        "건강보험료":         fixed_rate(path, "건강보험료"),
        "연금보험료":         fixed_rate(path, "연금보험료"),
        "퇴직공제부금비":     fixed_rate(path, "퇴직공제부금비"),
        "노인장기요양보험료": fixed_rate(path, "노인장기요양보험료"),
        "산업안전보건관리비": sanan_rate(
            path, sanan,
            params.get("sanan_kind") or mapping.default_sanan_kind(kind),
            params.get("est_cost"))["rate"],
        "기타경비":           table_rate(path, "기타경비",     kind, size, dur),
        "일반관리비":         table_rate(path, "일반관리비",   kind, size, dur),
        "이윤":               table_rate(path, "이윤",         kind, size, dur, contract),
    }
