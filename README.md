# 적용근거 생성기

조달청 제비율 엑셀 + 공사 파라미터 → 표준 적용근거 시트(xlsx) 생성.

## 다운로드 (Windows)

별도 설치 없이 실행파일 하나로 동작합니다.

- **최신 빌드:** 저장소의 **Releases → `latest`** 에서 `JebiyulApplier.exe` 다운로드
- main 브랜치에 푸시될 때마다 GitHub Actions가 Windows에서 자동으로 빌드해 올립니다.
- 특정 버전이 필요하면 `v1.0.0` 같은 태그를 푸시하면 해당 버전 Release가 만들어집니다.

## 사용

1. `JebiyulApplier.exe` 실행
2. 직접공사비 · 공사기간(일) · 공사종류 · 계약방법 · 산재기준 · 산안비 대상액 입력
3. 제비율 파일 선택 → [적용근거 생성]
4. 현재 디렉토리(실행 위치)에 `적용근거_결과.xlsx` 생성됨

## 개발

```bash
pip install -r requirements.txt
python -m pytest -v          # 제비율 원본 파일이 있으면 골든 테스트까지 실행, 없으면 자동 skip
pyinstaller app.spec         # dist/JebiyulApplier.exe 생성
```

> **주의:** PyInstaller는 크로스 컴파일을 지원하지 않습니다.
> Windows `.exe`는 반드시 **Windows 환경**에서 빌드해야 합니다(그래서 GitHub Actions의
> `windows-latest` 러너에서 빌드합니다). macOS / Linux에서 `pyinstaller app.spec`을
> 실행하면 해당 OS용 바이너리만 생성됩니다.

## 알려진 제약

- 산업안전보건관리비 `50억 이상` 구간은 800억 미만 세부구간 요율만 반환합니다(800억 이상
  공사는 수동 보정 필요).
- 수의계약 이윤율은 제비율 파일에 없어 조달청 표준(1000억 미만 10%, 이상 9%)을 적용합니다.
