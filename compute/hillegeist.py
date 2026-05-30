from __future__ import annotations

import math
from dataclasses import dataclass


@dataclass
class HillegeistInputs:
    working_capital: float
    retained_earnings: float
    ebit: float
    market_value_equity: float
    total_liabilities: float
    total_assets: float


@dataclass
class HillegeistOutputs:
    score: float
    default_prob: float


def solve(inp: HillegeistInputs) -> HillegeistOutputs:
    ta = inp.total_assets or 1e-9
    tl = inp.total_liabilities or 1e-9
    x1 = inp.working_capital / ta
    x2 = inp.retained_earnings / ta
    x3 = inp.ebit / ta
    x4 = inp.market_value_equity / tl
    score = 3.835 + 1.13 * x1 + 0.005 * x2 + 0.269 * x3 + 0.399 * x4
    # 1 / (1 + e^score): large positive score → ~0; large negative → ~1.
    if score > 50:
        prob = 0.0
    elif score < -50:
        prob = 1.0
    else:
        prob = 1 / (1 + math.exp(score))
    return HillegeistOutputs(score=score, default_prob=prob)
