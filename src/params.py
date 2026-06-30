"""공사 파라미터 → 제비율표 구간(밴드) 분류. 순수 함수."""

def size_band(jikjeop_cost: int) -> str:
    c = jikjeop_cost
    if c < 1_000_000_000:
        return "10억미만"
    if c < 5_000_000_000:
        return "10-50억"
    if c < 30_000_000_000:
        return "50-300억"
    if c < 100_000_000_000:
        return "300-1000억"
    return "1000억이상"

def duration_band(days: int) -> str:
    if days <= 183:
        return "183"
    if days <= 365:
        return "365"
    if days <= 1095:
        return "1095"
    return "1096+"

def sanan_band(target_cost: int) -> str:
    c = target_cost
    if c < 20_000_000:
        return "2천만미만"
    if c < 500_000_000:
        return "5억미만"
    if c < 5_000_000_000:
        return "5-50억"
    return "50억이상"
