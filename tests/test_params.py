import pytest
from src.params import size_band, duration_band, sanan_band

@pytest.mark.parametrize("cost,expected", [
    (999_999_999, "10억미만"),
    (1_000_000_000, "10-50억"),
    (4_999_999_999, "10-50억"),
    (5_000_000_000, "50-300억"),
    (29_999_999_999, "50-300억"),
    (30_000_000_000, "300-1000억"),
    (99_999_999_999, "300-1000억"),
    (100_000_000_000, "1000억이상"),
])
def test_size_band(cost, expected):
    assert size_band(cost) == expected

@pytest.mark.parametrize("days,expected", [
    (183, "183"), (1, "183"),
    (184, "365"), (365, "365"),
    (366, "1095"), (1095, "1095"),
    (1096, "1096+"), (2000, "1096+"),
])
def test_duration_band(days, expected):
    assert duration_band(days) == expected

@pytest.mark.parametrize("cost,expected", [
    (19_999_999, "2천만미만"),
    (20_000_000, "5억미만"),
    (499_999_999, "5억미만"),
    (500_000_000, "5-50억"),
    (4_999_999_999, "5-50억"),
    (5_000_000_000, "50억이상"),
])
def test_sanan_band(cost, expected):
    assert sanan_band(cost) == expected
