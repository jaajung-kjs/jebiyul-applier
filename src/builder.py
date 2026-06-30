"""표준 템플릿에 율을 기입해 결과 xlsx 저장."""
import os
import sys
import shutil

import openpyxl

# Canonical item list — single source of truth for the runtime package.
# Order must match the template row layout: row = 3 + i*2 (i = 0..11).
# tools/build_template.py imports from here at dev time; the reverse is NOT
# allowed (tools/ is not shipped in the PyInstaller bundle).
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

# Absolute path to the bundled template asset.
# When running inside a PyInstaller --onefile bundle, sys._MEIPASS points to
# the temporary extraction directory; otherwise fall back to the repo root.
_BASE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_ASSET = os.path.join(_BASE, "assets", "template_적용근거.xlsx")

# 항목 → 율 기입 셀 좌표.  템플릿 계약: A열 라벨은 행 3,5,7,…(3+i*2), 율은 I열 동일 행.
CELL_MAP: dict[str, str] = {item: f"I{3 + i * 2}" for i, item in enumerate(ITEMS)}


def build_output(
    rates: dict,
    out_path: str,
    template_path: str = _ASSET,
) -> str:
    """템플릿을 복사해 rates dict의 각 항목 율을 I열에 기입한 후 저장한다.

    Args:
        rates:         항목명 → 적용율(소수) dict.
        out_path:      저장할 파일 경로.
        template_path: 기본값은 assets/template_적용근거.xlsx.

    Returns:
        저장된 파일의 경로(out_path).
    """
    shutil.copyfile(template_path, out_path)
    wb = openpyxl.load_workbook(out_path)
    ws = wb["적용근거"]
    for item, cell in CELL_MAP.items():
        if item in rates:
            ws[cell] = rates[item]
    wb.save(out_path)
    return out_path
