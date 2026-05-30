from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd
from scipy.optimize import brentq


@dataclass
class DCFInputs:
    ufcf: list[float]        # forecast unlevered FCF, year 1..N
    discount_rate: float     # WACC or chosen rate (decimal)
    terminal_value: float    # terminal value at end of year N (undiscounted)
    net_debt: float          # debt - cash, current
    minority_interest: float = 0.0
    shares_out: float = 0.0


def pv(values: list[float], r: float) -> float:
    return sum(v / (1 + r) ** (i + 1) for i, v in enumerate(values))


def enterprise_value(inputs: DCFInputs) -> float:
    pv_ufcf = pv(inputs.ufcf, inputs.discount_rate)
    n = len(inputs.ufcf)
    pv_tv = inputs.terminal_value / (1 + inputs.discount_rate) ** n
    return pv_ufcf + pv_tv


def equity_value(inputs: DCFInputs) -> float:
    return enterprise_value(inputs) - (inputs.net_debt or 0) - (inputs.minority_interest or 0)


def price_per_share(inputs: DCFInputs) -> float | None:
    if inputs.shares_out in (None, 0):
        return None
    return equity_value(inputs) / inputs.shares_out


def implied_perpetuity_growth(inputs: DCFInputs, terminal_year_fcf: float) -> float | None:
    """g such that TV = FCF_{N+1} / (r - g)."""
    if inputs.terminal_value in (None, 0) or terminal_year_fcf in (None, 0):
        return None
    return inputs.discount_rate - terminal_year_fcf / inputs.terminal_value


def sensitivity_grid(ufcf: list[float], terminal_ebitda: float,
                     discount_rates: list[float], terminal_multiples: list[float],
                     net_debt: float, minority: float, shares: float,
                     current_price: float) -> dict[str, pd.DataFrame]:
    """Returns dict of DataFrames: tev, equity, price, upside."""
    tev = pd.DataFrame(index=discount_rates, columns=terminal_multiples, dtype=float)
    eq = pd.DataFrame(index=discount_rates, columns=terminal_multiples, dtype=float)
    px = pd.DataFrame(index=discount_rates, columns=terminal_multiples, dtype=float)
    up = pd.DataFrame(index=discount_rates, columns=terminal_multiples, dtype=float)

    for r in discount_rates:
        for m in terminal_multiples:
            tv = terminal_ebitda * m
            inp = DCFInputs(ufcf=ufcf, discount_rate=r, terminal_value=tv,
                            net_debt=net_debt, minority_interest=minority,
                            shares_out=shares)
            t = enterprise_value(inp)
            e = equity_value(inp)
            p = price_per_share(inp)
            tev.loc[r, m] = t
            eq.loc[r, m] = e
            px.loc[r, m] = p if p is not None else float("nan")
            up.loc[r, m] = (p / current_price - 1) if (p and current_price) else float("nan")

    return {"tev": tev, "equity": eq, "price": px, "upside": up}


def reverse_dcf(current_market_cap: float, current_fcf: float,
                discount_rate: float, years: int = 10,
                terminal_g: float = 0.025) -> float | None:
    """Solve for the constant growth rate g over `years` such that the resulting
    DCF (perpetuity terminal at terminal_g) equals current market cap."""

    def model_value(g: float) -> float:
        flows = [current_fcf * (1 + g) ** (i + 1) for i in range(years)]
        pv_flows = pv(flows, discount_rate)
        terminal_fcf = flows[-1] * (1 + terminal_g)
        tv = terminal_fcf / (discount_rate - terminal_g)
        pv_tv = tv / (1 + discount_rate) ** years
        return pv_flows + pv_tv - current_market_cap

    try:
        return brentq(model_value, -0.5, 1.0, xtol=1e-5)
    except (ValueError, RuntimeError):
        return None


def margin_of_safety(intrinsic_value: float, price: float) -> float | None:
    if intrinsic_value in (None, 0):
        return None
    return (intrinsic_value - price) / intrinsic_value
