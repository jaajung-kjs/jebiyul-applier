import os
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 제비율 원본 파일은 저장소에 포함되지 않는다(.gitignore). 로컬에 있으면 골든 테스트가
# 실행되고, CI처럼 파일이 없는 환경에서는 해당 테스트를 자동 skip 한다.
_TOMOK = os.path.join(ROOT, "붙임2. 토목공사 간접공사비 적용기준_조달청(260430 적용).xlsx")
_GEONCHUK = os.path.join(ROOT, "붙임1. 건축공사 간접공사비 적용기준_조달청(260430 적용).xlsx")


@pytest.fixture
def tomok_path():
    if not os.path.exists(_TOMOK):
        pytest.skip("제비율 원본(토목) 파일 없음 — 골든 테스트 skip")
    return _TOMOK


@pytest.fixture
def geonchuk_path():
    if not os.path.exists(_GEONCHUK):
        pytest.skip("제비율 원본(건축) 파일 없음 — 골든 테스트 skip")
    return _GEONCHUK
