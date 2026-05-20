from __future__ import annotations

import math
from dataclasses import dataclass

from scipy.stats import norm

from compute import kmv


@dataclass
class MertonInputs:
    equity_value: float
    equity_vol: float
    total_liab: float
    risk_free: float
    term_years: float = 1.0


@dataclass
class MertonOutputs:
    asset_value: float
    asset_vol: float
    put_value: float
    cds_spread_per_year: float
    default_prob_to_term: float
    annual_default_prob: float


def solve(inputs: MertonInputs) -> MertonOutputs:
    k_inputs = kmv.KMVInputs(
        equity_value=inputs.equity_value,
        equity_vol=inputs.equity_vol,
        risk_free=inputs.risk_free,
        long_term_liab=inputs.total_liab,
        short_term_liab=0.0,
        horizon_years=inputs.term_years,
    )
    k = kmv.solve(k_inputs)
    A, sigma_A, T = k.asset_value, k.asset_vol, inputs.term_years
    K = inputs.total_liab
    rf = inputs.risk_free

    d1 = (math.log(A / K) + (rf + 0.5 * sigma_A ** 2) * T) / (sigma_A * math.sqrt(T))
    d2 = d1 - sigma_A * math.sqrt(T)
    put = K * math.exp(-rf * T) * norm.cdf(-d2) - A * norm.cdf(-d1)
    pd_to_term = norm.cdf(-d2)
    annual_pd = 1 - (1 - pd_to_term) ** (1 / T) if T > 0 else pd_to_term
    cds = put / (K * T) if (K and T) else 0.0

    return MertonOutputs(
        asset_value=A,
        asset_vol=sigma_A,
        put_value=float(put),
        cds_spread_per_year=float(cds),
        default_prob_to_term=float(pd_to_term),
        annual_default_prob=float(annual_pd),
    )
