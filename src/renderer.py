"""sheet_model 블록 리스트 → xlsx. 행 커서로 위에서 아래로 그린다.

좌표는 블록 순서와 행 커서로 결정된다(고정 좌표 없음). 분기 판단은 sheet_model이
끝냈고 렌더러는 그리기만 한다.
"""
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from src import styles
from src import sheet_model as M


class _Cursor:
    def __init__(self):
        self.row = 1

    def take(self, n=1):
        r = self.row
        self.row += n
        return r


def _apply_dims(ws):
    for col, w in styles.COL_WIDTHS.items():
        ws.column_dimensions[col].width = w


def render(blocks, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "적용근거"
    _apply_dims(ws)
    cur = _Cursor()
    cur.take(1)  # row 1 여백
    for block in blocks:
        _emit(ws, block, cur)
    wb.save(out_path)
    return out_path


def _emit(ws, block, cur):
    if isinstance(block, M.Title):
        _emit_title(ws, block, cur)
    elif isinstance(block, M.SectionHeader):
        _emit_section(ws, block, cur)
    elif isinstance(block, M.SubHeader):
        _emit_sub(ws, block, cur)
    elif isinstance(block, M.NoteLines):
        _emit_notes(ws, block, cur)
    elif isinstance(block, M.AppliedRate):
        _emit_applied(ws, block, cur)
    elif isinstance(block, M.BandTable):
        _emit_table(ws, block, cur)   # Task 4~5에서 구현
    else:
        raise TypeError(f"unknown block: {block!r}")


def _emit_title(ws, block, cur):
    r = cur.take(1)
    ws.row_dimensions[r].height = styles.ROW_TITLE_H
    ws.merge_cells(f"A{r}:J{r}")
    c = ws[f"A{r}"]
    c.value = block.text
    c.font = styles.title_font()
    c.alignment = styles.center()
    cur.take(1)  # 섹션 사이 여백 한 줄


def _emit_section(ws, block, cur):
    r = cur.take(1)
    ws.row_dimensions[r].height = styles.ROW_BODY_H
    c = ws[f"A{r}"]
    c.value = block.text
    c.font = styles.section_font()
    c.alignment = styles.left()
    if block.note:
        n = ws[f"J{r}"]
        n.value = block.note
        n.font = styles.body_font()
        n.alignment = styles.right()


def _emit_sub(ws, block, cur):
    r = cur.take(1)
    ws.row_dimensions[r].height = styles.ROW_BODY_H
    c = ws[f"A{r}"]
    c.value = block.text
    c.font = styles.body_font()
    c.alignment = styles.left()


def _emit_notes(ws, block, cur):
    for line in block.lines:
        r = cur.take(1)
        ws.row_dimensions[r].height = styles.ROW_BODY_H
        c = ws[f"A{r}"]
        c.value = line
        c.font = styles.body_font()
        c.alignment = styles.left()


def _emit_applied(ws, block, cur):
    r = cur.take(1)
    ws.row_dimensions[r].height = styles.ROW_BODY_H
    lab = ws[f"A{r}"]
    lab.value = block.label
    lab.font = styles.rate_font()
    lab.alignment = styles.left()
    val = ws[f"I{r}"]
    val.value = block.value
    val.number_format = block.fmt
    val.font = styles.rate_font()
    val.alignment = styles.right()
    if block.annotation:
        ann = ws[f"J{r}"]
        ann.value = block.annotation
        ann.font = styles.rate_font()
        ann.alignment = styles.left()


def _emit_table(ws, block, cur):
    # 헤더행(1~2행)
    hr = cur.take(1)
    ws.row_dimensions[hr].height = styles.ROW_BODY_H
    for h in block.headers:
        text, c0, c1, rowspan = _norm_header(h)
        _put(ws, hr, c0, c1, text, fill=styles.header_fill(),
             align=styles.center(wrap=True), border=True, rowspan=rowspan)
    if block.subheaders:
        sr = cur.take(1)
        ws.row_dimensions[sr].height = styles.ROW_BODY_H
        for text, c0, c1 in block.subheaders:
            _put(ws, sr, c0, c1, text, fill=styles.header_fill(),
                 align=styles.center(wrap=True), border=True)
    # 데이터행
    first = cur.row
    for brow in block.rows:
        r = cur.take(1)
        ws.row_dimensions[r].height = styles.ROW_BODY_H
        for cell in brow.cells:
            value, c0, c1, fmt, rowspan = _norm_cell(cell)
            # 강조는 단일 행 셀에만. 세로 병합된 규모/구분 라벨은 강조색 제외.
            fill = styles.highlight_fill() if (brow.highlight and rowspan == 1) else None
            _put(ws, r, c0, c1, value, fill=fill, align=styles.center(),
                 border=True, fmt=fmt, rowspan=rowspan)
    last = cur.row - 1
    # 적용기준 세로 병합 블록(G:J)
    if block.criterion:
        _put(ws, first, 7, 10, block.criterion, align=styles.center(wrap=True),
             border=True, rowspan=last - first + 1)


def _norm_cell(cell):
    """(value, c0, c1, fmt[, rowspan]) → 5-튜플로 정규화."""
    if len(cell) == 5:
        return cell
    value, c0, c1, fmt = cell
    return value, c0, c1, fmt, 1


def _norm_header(h):
    """(text, c0, c1[, rowspan]) → 4-튜플로 정규화."""
    if len(h) == 4:
        return h
    text, c0, c1 = h
    return text, c0, c1, 1


def _put(ws, row, c0, c1, value, fill=None, align=None, border=False, fmt=None, rowspan=1):
    r1 = row + rowspan - 1
    a = f"{get_column_letter(c0)}{row}"
    if c1 > c0 or r1 > row:
        ws.merge_cells(f"{a}:{get_column_letter(c1)}{r1}")
    c = ws[a]
    c.value = value
    c.font = styles.body_font()
    if align:
        c.alignment = align
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if border:
        b = styles.box_border()
        for rr in range(row, r1 + 1):
            for col in range(c0, c1 + 1):
                ws.cell(rr, col).border = b
