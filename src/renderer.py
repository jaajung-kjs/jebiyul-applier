"""sheet_model 블록 리스트 → xlsx. 행 커서로 위에서 아래로 그린다.

좌표는 블록 순서와 행 커서로 결정된다(고정 좌표 없음). 분기 판단은 sheet_model이
끝냈고 렌더러는 그리기만 한다.
"""
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from src import styles
from src import sheet_model as M
from src import nomu_model as NM


class _Cursor:
    def __init__(self):
        self.row = 1

    def take(self, n=1):
        r = self.row
        self.row += n
        return r


def _apply_dims(ws, col_widths):
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w


def render_sheet(ws, blocks, col_widths=None):
    _apply_dims(ws, col_widths or styles.COL_WIDTHS)
    cur = _Cursor()
    cur.take(1)  # row 1 여백
    for block in blocks:
        _emit(ws, block, cur)
    return ws


def render(blocks, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "적용근거"
    render_sheet(ws, blocks, styles.COL_WIDTHS)
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
    elif isinstance(block, NM.NomuTitle):
        _emit_nomu_title(ws, block, cur)
    elif isinstance(block, NM.NomuHeader):
        _emit_nomu_header(ws, block, cur)
    elif isinstance(block, NM.NomuGroup):
        _emit_nomu_group(ws, block, cur)
    elif isinstance(block, NM.NomuRow):
        _emit_nomu_row(ws, block, cur)
    elif isinstance(block, NM.NomuAvg):
        _emit_nomu_avg(ws, block, cur)
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


def _emit_section(ws, block, cur):
    cur.take(1)  # 대분류(1·2·3·4) 앞 빈 행 하나(가독성)
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
    if isinstance(block.value, (int, float)):
        val.number_format = block.fmt   # 숫자일 때만 % 서식(문자 '적용제외'엔 미적용)
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


_NOMU_C0 = 4  # 과거열 시작(D)


def _nomu_cols(p):
    """과거열 p개일 때 (과거 시작, 현재, 변동율, 비고) 열 인덱스."""
    cur_c = _NOMU_C0 + p
    return _NOMU_C0, cur_c, cur_c + 1, cur_c + 2


def _nomu_hframe(ws, r, last_col):
    """행 1..last_col 전체에 가로 테두리 + 좌/우 외곽선을 깐다.

    내용 있는 셀은 이후 _put(box_border)로 덮어써 세로 구분선까지 갖고, 빈 칸은
    가로·외곽선만 남아 표 테두리가 끊기지 않는다.
    """
    for col in range(1, last_col + 1):
        ws.cell(r, col).border = styles.hframe_border(
            left=(col == 1), right=(col == last_col))


def _emit_nomu_title(ws, block, cur):
    r = cur.take(1)
    ws.row_dimensions[r].height = styles.ROW_TITLE_H
    end_column = _nomu_cols(block.past_count)[-1]
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=end_column)
    c = ws.cell(r, 1)
    c.value = block.text
    c.font = styles.bold_font()
    c.alignment = styles.center()


def _emit_nomu_header(ws, block, cur):
    p = len(block.past_cols)
    past0, cur_c, delta_c, note_c = _nomu_cols(p)
    r1 = cur.take(1)
    r2 = cur.take(1)
    for r in (r1, r2):
        ws.row_dimensions[r].height = styles.ROW_BODY_H

    def hdr(row, c0, c1, text, rowspan=1):
        _put(ws, row, c0, c1, text, fill=styles.header_fill(),
             align=styles.center(wrap=True), border=True, rowspan=rowspan)

    hdr(r1, 1, 1, "번호", rowspan=2)
    hdr(r1, 2, 2, "직  종  명", rowspan=2)
    hdr(r1, 3, 3, "No.", rowspan=2)
    hdr(r1, past0, past0 + p - 1, "공 표 일")
    hdr(r1, cur_c, cur_c, block.current_col, rowspan=2)
    hdr(r1, delta_c, delta_c, "변동율\n(%)", rowspan=2)
    hdr(r1, note_c, note_c, "비  고", rowspan=2)
    for j, label in enumerate(block.past_cols):
        hdr(r2, past0 + j, past0 + j, label)


def _emit_nomu_group(ws, block, cur):
    r = cur.take(1)
    ws.row_dimensions[r].height = styles.ROW_BODY_H
    last_col = _nomu_cols(block.past_count)[-1]
    _nomu_hframe(ws, r, last_col)
    _put(ws, r, 2, 2, f"{block.roman}. {block.name}",
         align=styles.left(), border=True)
    ws.cell(r, 2).font = styles.bold_font()


def _emit_nomu_row(ws, block, cur):
    p = len(block.past_wages)
    past0, cur_c, delta_c, note_c = _nomu_cols(p)
    r = cur.take(1)
    ws.row_dimensions[r].height = styles.ROW_BODY_H
    _put(ws, r, 1, 1, block.no, align=styles.center(), border=True)
    _put(ws, r, 2, 2, block.name, align=styles.left(), border=True)
    _put(ws, r, 3, 3, int(block.code), align=styles.center(), border=True)
    for j, w in enumerate(block.past_wages):
        _put(ws, r, past0 + j, past0 + j, w, align=styles.right(),
             border=True, fmt=styles.COMMA_FMT)
    _put(ws, r, cur_c, cur_c, block.current_wage, align=styles.right(),
         border=True, fmt=styles.COMMA_FMT)
    _put(ws, r, delta_c, delta_c, block.delta, align=styles.center(),
         border=True, fmt=styles.PCT_FMT)
    _put(ws, r, note_c, note_c, block.note or None, align=styles.center(), border=True)


def _emit_nomu_avg(ws, block, cur):
    past0, cur_c, delta_c, note_c = _nomu_cols(block.past_count)
    r = cur.take(1)
    ws.row_dimensions[r].height = styles.ROW_BODY_H
    _nomu_hframe(ws, r, note_c)
    _put(ws, r, 2, 2, "노임변동률평균", align=styles.center(), border=True)
    _put(ws, r, delta_c, delta_c, block.value, align=styles.center(),
         border=True, fmt=styles.PCT_FMT)
