"""표준 템플릿에 율을 기입해 결과 xlsx 저장."""
import os
import shutil

import openpyxl

# Single source of truth: import ITEMS from the template module so that the
# cell-row mapping here can never silently drift from the template layout.
# tools/ already has __init__.py so this import works from the repo root.
from tools.build_template import ITEMS

# Absolute path to the bundled template asset.
_ASSET = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    "assets",
    "template_적용근거.xlsx",
)

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
