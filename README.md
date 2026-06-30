# 적용근거 생성기

조달청 제비율 엑셀 + 공사 파라미터 → 표준 적용근거 시트(xlsx) 생성.

## 사용

1. `적용근거생성기.exe` 실행
2. 직접공사비 · 공사기간(일) · 공사종류 · 계약방법 · 산재기준 · 산안비 대상액 입력
3. 제비율 파일 선택 → [적용근거 생성]
4. 바탕화면에 `적용근거_결과.xlsx` 생성됨

## 개발

```bash
pip install -r requirements.txt
python -m pytest -v
pyinstaller app.spec
```

> **주의:** PyInstaller는 크로스 컴파일을 지원하지 않습니다.
> Windows `.exe`는 반드시 **Windows 환경**에서 빌드해야 합니다.
> macOS / Linux에서 `pyinstaller app.spec`을 실행하면 해당 OS용 바이너리만 생성됩니다.
