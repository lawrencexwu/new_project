from __future__ import annotations

from typing import Mapping


def safe_div(num, denom):
    if num is None or denom in (None, 0):
        return None
    try:
        return num / denom
    except ZeroDivisionError:
        return None


def gross_margin(revenue, cost_of_revenue):
    return safe_div((revenue or 0) - (cost_of_revenue or 0), revenue)


def operating_margin(operating_income, revenue):
    return safe_div(operating_income, revenue)


def ebit_margin(ebit, revenue):
    return safe_div(ebit, revenue)


def fcf(ocf, capex):
    if ocf is None or capex is None:
        return None
    return ocf - abs(capex)


def fcf_margin(ocf, capex, revenue):
    return safe_div(fcf(ocf, capex), revenue)


def quality_of_earnings(ocf, net_income):
    return safe_div(ocf, net_income)


def current_ratio(current_assets, current_liabilities):
    return safe_div(current_assets, current_liabilities)


def quick_ratio(current_assets, inventory, current_liabilities):
    return safe_div((current_assets or 0) - (inventory or 0), current_liabilities)


def debt_to_equity(total_debt, equity):
    return safe_div(total_debt, equity)


def debt_to_assets(total_debt, total_assets):
    return safe_div(total_debt, total_assets)


def debt_to_fcf(total_debt, ocf, capex):
    f = fcf(ocf, capex)
    return safe_div(total_debt, f)


def debt_to_ocf(total_debt, ocf):
    return safe_div(total_debt, ocf)


def interest_cover(ebit, interest_expense):
    if interest_expense in (None, 0):
        return None
    return safe_div(ebit, abs(interest_expense))


def fcf_to_interest(ocf, capex, interest_expense):
    if interest_expense in (None, 0):
        return None
    return safe_div(fcf(ocf, capex), abs(interest_expense))


def capex_pct_ocf(capex, ocf):
    if capex is None or ocf in (None, 0):
        return None
    return safe_div(abs(capex), ocf)


def days_sales_outstanding(receivables, revenue, period_days=90):
    if receivables is None or revenue in (None, 0):
        return None
    return (receivables / revenue) * period_days


def inventory_turnover(cost_of_revenue, inventory):
    return safe_div(cost_of_revenue, inventory)


def roe(net_income, equity):
    return safe_div(net_income, equity)


def roa(net_income, total_assets):
    return safe_div(net_income, total_assets)


def roic(nopat, invested_capital):
    return safe_div(nopat, invested_capital)


def nopat(ebit, tax_rate):
    if ebit is None:
        return None
    return ebit * (1 - (tax_rate or 0))


def croci(ebitda, gross_invested_capital):
    return safe_div(ebitda, gross_invested_capital)


def dupont_5step(net_income, pretax_income, ebit, revenue, total_assets, equity):
    """Returns (tax_burden, interest_burden, op_margin, asset_turnover, equity_multiplier)."""
    return (
        safe_div(net_income, pretax_income),
        safe_div(pretax_income, ebit),
        safe_div(ebit, revenue),
        safe_div(revenue, total_assets),
        safe_div(total_assets, equity),
    )


def cagr(start_value, end_value, periods):
    if start_value in (None, 0) or end_value is None or periods in (None, 0):
        return None
    try:
        return (end_value / start_value) ** (1 / periods) - 1
    except (ValueError, ZeroDivisionError):
        return None
