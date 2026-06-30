"""Tests for validate_inputs in src/gui.py.

All tests are headless — they only import and call validate_inputs,
never instantiate Tk(), so no display is required in CI.
"""
import pytest
from src.gui import validate_inputs


def base():
    return {
        "jikjeop_cost": "500000000",
        "days": "200",
        "kind": "토목",
        "contract": "경쟁",
        "sanjae_basis": "한전",
        "sanan_target": "400000000",
        "jebiyul_path": "x.xlsx",
    }


# --- From the brief ---

def test_valid_inputs_cast_types():
    p = validate_inputs(base())
    assert p["jikjeop_cost"] == 500_000_000 and p["days"] == 200


def test_missing_cost_raises():
    r = base(); r["jikjeop_cost"] = ""
    with pytest.raises(ValueError):
        validate_inputs(r)


def test_non_numeric_days_raises():
    r = base(); r["days"] = "여섯달"
    with pytest.raises(ValueError):
        validate_inputs(r)


def test_bad_kind_raises():
    r = base(); r["kind"] = "조선"
    with pytest.raises(ValueError):
        validate_inputs(r)


# --- Additional tests required by Task 9 ---

def test_bad_contract_raises():
    r = base(); r["contract"] = "직접"
    with pytest.raises(ValueError):
        validate_inputs(r)


def test_bad_sanjae_basis_raises():
    r = base(); r["sanjae_basis"] = "공단"
    with pytest.raises(ValueError):
        validate_inputs(r)


def test_missing_jebiyul_path_raises():
    r = base(); r["jebiyul_path"] = ""
    with pytest.raises(ValueError):
        validate_inputs(r)


def test_comma_formatted_amounts_parse_correctly():
    r = base()
    r["jikjeop_cost"] = "1,000,000,000"
    r["sanan_target"] = "800,000,000"
    p = validate_inputs(r)
    assert p["jikjeop_cost"] == 1_000_000_000
    assert p["sanan_target"] == 800_000_000


def test_all_keys_present_in_result():
    p = validate_inputs(base())
    for key in ("jikjeop_cost", "days", "kind", "contract",
                 "sanjae_basis", "sanan_target", "jebiyul_path"):
        assert key in p


def test_valid_all_kinds():
    from src.mapping import KINDS
    for k in KINDS:
        r = base(); r["kind"] = k
        p = validate_inputs(r)
        assert p["kind"] == k


def test_valid_suui_contract():
    r = base(); r["contract"] = "수의"
    p = validate_inputs(r)
    assert p["contract"] == "수의"


def test_valid_jodalcheong_sanjae():
    r = base(); r["sanjae_basis"] = "조달청"
    p = validate_inputs(r)
    assert p["sanjae_basis"] == "조달청"
