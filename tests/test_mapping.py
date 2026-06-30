from src.mapping import rate_source_kind, SIZE_BASE_ROW, DURATION_OFFSET

def test_jeonmun_uses_tomok_for_non_ilban():
    assert rate_source_kind("간접노무비", "전기통신소방전문") == "토목"
    assert rate_source_kind("기타경비", "전기통신소방전문") == "토목"

def test_jeonmun_uses_self_for_ilban():
    assert rate_source_kind("일반관리비", "전기통신소방전문") == "전기통신소방전문"

def test_normal_kind_unchanged():
    assert rate_source_kind("간접노무비", "토목") == "토목"

def test_row_arithmetic():
    # 10-50억 / 1095일 → 23 + 4 = 27
    assert SIZE_BASE_ROW["10-50억"] + DURATION_OFFSET["1095"] == 27
