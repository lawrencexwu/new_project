from __future__ import annotations

from dataclasses import dataclass


@dataclass
class AltmanInputs:
    working_capital: float
    retained_earnings: float
    ebit: float
    market_value_equity: float
    total_liabilities: float
    sales: float
    total_assets: float


@dataclass
class AltmanOutputs:
    x1: float
    x2: float
    x3: float
    x4: float
    x5: float
    z_score: float
    bucket: str


def solve(inp: AltmanInputs) -> AltmanOutputs:
    ta = inp.total_assets or 1e-9
    tl = inp.total_liabilities or 1e-9
    x1 = inp.working_capital / ta
    x2 = inp.retained_earnings / ta
    x3 = inp.ebit / ta
    x4 = inp.market_value_equity / tl
    x5 = inp.sales / ta
    z = 1.2 * x1 + 1.4 * x2 + 3.3 * x3 + 0.6 * x4 + 1.0 * x5
    if z > 2.99:
        bucket = "Safe"
    elif z > 1.81:
        bucket = "Grey"
    else:
        bucket = "Distress"
    return AltmanOutputs(x1=x1, x2=x2, x3=x3, x4=x4, x5=x5,
                        z_score=z, bucket=bucket)
