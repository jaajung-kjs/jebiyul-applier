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


def _emit_table(ws, block, cur):
    # Task 4~5에서 구현. 현재는 건너뛴다.
    return
