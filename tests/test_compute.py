"""Unit tests for compute layer. Run with: python -m pytest tests/"""
import math

import pandas as pd
import pytest

from compute import altman, dcf, hillegeist, kmv, merton, multiples, ratios, scoring, signals


def test_ratios_safe_div():
    assert ratios.safe_div(10, 2) == 5
    assert ratios.safe_div(10, 0) is None
    assert ratios.safe_div(None, 5) is None


def test_ratios_margins_and_fcf():
    assert ratios.gross_margin(100, 60) == pytest.approx(0.4)
    assert ratios.operating_margin(20, 100) == pytest.approx(0.2)
    assert ratios.fcf(100, -30) == 70
    assert ratios.fcf(100, 30) == 70  # capex absolute-valued


def test_ratios_dupont():
    tb, ib, om, at, em = ratios.dupont_5step(
        net_income=80, pretax_income=100, ebit=120,
        revenue=1000, total_assets=2000, equity=800,
    )
    assert tb == pytest.approx(0.8)
    assert ib == pytest.approx(100 / 120)
    assert om == pytest.approx(0.12)
    assert at == pytest.approx(0.5)
    assert em == pytest.approx(2.5)


def test_owner_earnings():
    assert ratios.owner_earnings(net_income=100, da=20, maintenance_capex=15) == 105
    assert ratios.owner_earnings(net_income=100, da=20, maintenance_capex=-15) == 105  # abs
    assert ratios.owner_earnings(None, 20, 15) is None


def test_capital_allocation_breakdown():
    bd = ratios.capital_allocation_breakdown(
        ocf=1000, capex=-200, buybacks=300, dividends=100,
        m_and_a=50, debt_paydown=150,
    )
    assert bd["CAPEX"] == pytest.approx(0.20)
    assert bd["Buybacks"] == pytest.approx(0.30)
    assert bd["Dividends"] == pytest.approx(0.10)
    assert bd["M&A"] == pytest.approx(0.05)
    assert bd["Debt Paydown"] == pytest.approx(0.15)


def test_ratios_cagr():
    assert ratios.cagr(100, 200, 5) == pytest.approx(2 ** 0.2 - 1)
    assert ratios.cagr(0, 200, 5) is None


def test_scoring_bands():
    assert scoring.score(0.12, scoring.HIGHER_IS_BETTER) == 7
    assert scoring.score(-0.05, scoring.HIGHER_IS_BETTER) == 1
    assert scoring.score(None, scoring.HIGHER_IS_BETTER) is None
    assert scoring.score(0.35, scoring.LOWER_IS_BETTER) == 8


def test_dcf_pv_and_ev():
    inp = dcf.DCFInputs(
        ufcf=[100, 110, 120, 130, 140],
        discount_rate=0.10,
        terminal_value=2000,
        net_debt=500,
        shares_out=100,
    )
    ev = dcf.enterprise_value(inp)
    expected_pv_ufcf = sum(f / (1.1 ** (i + 1)) for i, f in enumerate(inp.ufcf))
    expected_tv = 2000 / (1.1 ** 5)
    assert ev == pytest.approx(expected_pv_ufcf + expected_tv)
    assert dcf.equity_value(inp) == pytest.approx(ev - 500)
    assert dcf.price_per_share(inp) == pytest.approx((ev - 500) / 100)


def test_dcf_reverse():
    g = dcf.reverse_dcf(current_market_cap=1500, current_fcf=100,
                        discount_rate=0.10, years=10, terminal_g=0.025)
    assert g is not None
    assert -0.5 < g < 1.0


def test_dcf_margin_of_safety():
    assert dcf.margin_of_safety(100, 70) == pytest.approx(0.3)
    assert dcf.margin_of_safety(0, 50) is None


def test_kmv_solver_converges():
    inputs = kmv.KMVInputs(
        equity_value=2600.0, equity_vol=0.73,
        risk_free=0.045,
        long_term_liab=10000.0, short_term_liab=4400.0,
        horizon_years=1.2,
    )
    out = kmv.solve(inputs)
    assert out.squared_error < 1.0
    assert 0 <= out.edf <= 1
    assert out.default_point == pytest.approx(0.5 * 10000 + 4400)


def test_merton_outputs():
    inp = merton.MertonInputs(
        equity_value=1631.0, equity_vol=0.729,
        total_liab=14450.0, risk_free=0.045, term_years=1.0,
    )
    out = merton.solve(inp)
    assert out.put_value > 0
    assert 0 <= out.annual_default_prob <= 1


def test_altman_buckets():
    inp = altman.AltmanInputs(
        working_capital=200, retained_earnings=500, ebit=300,
        market_value_equity=2000, total_liabilities=1500,
        sales=4000, total_assets=3000,
    )
    out = altman.solve(inp)
    assert out.z_score == pytest.approx(
        1.2 * (200 / 3000) + 1.4 * (500 / 3000) + 3.3 * (300 / 3000)
        + 0.6 * (2000 / 1500) + 1.0 * (4000 / 3000)
    )
    assert out.bucket in ("Safe", "Grey", "Distress")


def test_hillegeist_probability_in_range():
    inp = hillegeist.HillegeistInputs(
        working_capital=200, retained_earnings=500, ebit=300,
        market_value_equity=2000, total_liabilities=1500, total_assets=3000,
    )
    out = hillegeist.solve(inp)
    assert 0 <= out.default_prob <= 1


def test_multiples_band():
    s = pd.Series([10, 12, 14, 16, 18, 20])
    b = multiples.band_from_series("P/E", current=20, historical=s)
    assert b.current == 20
    assert b.median_5y is not None


def test_signals_basic():
    idx = pd.date_range("2024-01-01", periods=250, freq="D")
    close = pd.Series(range(250), index=idx).astype(float) + 100
    assert signals.daily_buy_signal(close) is True
    assert signals.above_rising_5dma(close) is True
    assert signals.pct_off_52w_high(close) == pytest.approx(0.0)
    regime = signals.market_regime(close)
    assert regime in ("Trend Up", "Trend Down", "Chop", "Unknown")
