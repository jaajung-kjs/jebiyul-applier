import os
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

@pytest.fixture
def tomok_path():
    return os.path.join(ROOT, "붙임2. 토목공사 간접공사비 적용기준_조달청(260430 적용).xlsx")

@pytest.fixture
def geonchuk_path():
    return os.path.join(ROOT, "붙임1. 건축공사 간접공사비 적용기준_조달청(260430 적용).xlsx")
