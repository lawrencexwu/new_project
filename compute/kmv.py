from __future__ import annotations

import math
from dataclasses import dataclass

import numpy as np
from scipy.optimize import fsolve
from scipy.stats import norm


@dataclass
class KMVInputs:
    equity_value: float       # E
    equity_vol: float         # sigma_E (annualized, decimal)
    risk_free: float          # rf (decimal)
    long_term_liab: float
    short_term_liab: float
    horizon_years: float = 1.0


@dataclass
class KMVOutputs:
    asset_value: float
    asset_vol: float
    default_point: float
    distance_to_default: float
    edf: float
    squared_error: float


def default_point(lt: float, st: float) -> float:
    return 0.5 * lt + st


def _equations(asset_params, E, sigma_E, rf, K, T):
    A, sigma_A = asset_params
    if A <= 0 or sigma_A <= 0:
        return (1e9, 1e9)
    d1 = (math.log(A / K) + (rf + 0.5 * sigma_A ** 2) * T) / (sigma_A * math.sqrt(T))
    d2 = d1 - sigma_A * math.sqrt(T)
    eq1 = A * norm.cdf(d1) - K * math.exp(-rf * T) * norm.cdf(d2) - E
    eq2 = (A * norm.cdf(d1) * sigma_A / E) - sigma_E
    return (eq1, eq2)


def solve(inputs: KMVInputs) -> KMVOutputs:
    K = default_point(inputs.long_term_liab, inputs.short_term_liab)
    E = inputs.equity_value
    sigma_E = inputs.equity_vol
    rf = inputs.risk_free
    T = inputs.horizon_years

    A0 = E + K
    sigma_A0 = sigma_E * E / (E + K) if (E + K) else sigma_E

    sol, info, ier, msg = fsolve(
        _equations, x0=(A0, sigma_A0), args=(E, sigma_E, rf, K, T),
        full_output=True
    )
    A, sigma_A = sol
    err1, err2 = _equations((A, sigma_A), E, sigma_E, rf, K, T)
    sq_err = err1 ** 2 + err2 ** 2

    dd = (math.log(A / K) + (rf - 0.5 * sigma_A ** 2) * T) / (sigma_A * math.sqrt(T))
    edf = 1 - norm.cdf(dd)

    return KMVOutputs(
        asset_value=float(A),
        asset_vol=float(sigma_A),
        default_point=float(K),
        distance_to_default=float(dd),
        edf=float(edf),
        squared_error=float(sq_err),
    )
