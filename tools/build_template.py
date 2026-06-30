"""표준 적용근거 템플릿 xlsx 생성. 율 셀은 빈 칸으로 두고 라벨만 배치한다."""
import openpyxl
from openpyxl.styles import Font

# 12개 항목 — Task 8 CELL_MAP 계약과 순서가 동일해야 함.
# rate cell = I{3 + i*2}  (i = 0..11)
ITEMS = [
    "간접노무비",
    "공구손료",
    "산재보험료",
    "고용보험료",
    "건강보험료",
    "연금보험료",
    "퇴직공제부금비",
    "노인장기요양보험료",
    "산업안전보건관리비",
    "기타경비",
    "일반관리비",
    "이윤",
]


def build(path: str = "assets/template_적용근거.xlsx") -> None:
    """빈 골격 템플릿을 생성한다.

    레이아웃 계약 (Task 8 builder와 동기화 유지):
      - A1 : 제목
      - A{3 + i*2} : 항목 라벨  (i = 0..11 → 행 3, 5, 7, … 25)
      - G{3 + i*2} : "☞ 적 용 율 :"  (설명 텍스트)
      - I{3 + i*2} : 적용율 값 — builder가 채움 (여기서는 비워 둠)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "적용근거"

    # 제목
    ws["A1"] = "공사비 산출 적용근거"
    ws["A1"].font = Font(bold=True, size=14)

    # 12개 항목 배치
    for i, item in enumerate(ITEMS):
        row = 3 + i * 2
        ws.cell(row=row, column=1, value=item).font = Font(bold=True)
        ws.cell(row=row, column=7, value="☞ 적 용 율 :")
        # column 9 (I열)은 builder가 채움 — 비워 둠

    wb.save(path)
    print(f"Template saved → {path}")


if __name__ == "__main__":
    build()
