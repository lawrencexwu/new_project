"""Test EDGAR statement assembly against a synthetic company-facts blob."""
from __future__ import annotations

import json

import pandas as pd

from data import edgar_client


def _fake_facts():
    return {
        "entityName": "Acme Corp",
        "facts": {
            "us-gaap": {
                "Revenues": {
                    "units": {
                        "USD": [
                            {"end": "2024-03-31", "val": 100, "form": "10-Q", "fp": "Q1", "filed": "2024-05-01"},
                            {"end": "2024-06-30", "val": 110, "form": "10-Q", "fp": "Q2", "filed": "2024-08-01"},
                            {"end": "2024-09-30", "val": 120, "form": "10-Q", "fp": "Q3", "filed": "2024-11-01"},
                            {"end": "2024-12-31", "val": 130, "form": "10-K", "fp": "Q4", "filed": "2025-02-15"},
                            # Annual rollup we want to filter out
                            {"end": "2024-12-31", "val": 460, "form": "10-K", "fp": "FY", "filed": "2025-02-15"},
                        ]
                    }
                },
                "NetIncomeLoss": {
                    "units": {
                        "USD": [
                            {"end": "2024-03-31", "val": 15, "form": "10-Q", "fp": "Q1", "filed": "2024-05-01"},
                            {"end": "2024-06-30", "val": 18, "form": "10-Q", "fp": "Q2", "filed": "2024-08-01"},
                            # Same quarter restated later — newer filed wins
                            {"end": "2024-06-30", "val": 19, "form": "10-Q", "fp": "Q2", "filed": "2024-09-15"},
                        ]
                    }
                },
                "Assets": {
                    "units": {
                        "USD": [
                            {"end": "2024-03-31", "val": 1000, "form": "10-Q", "fp": "Q1", "filed": "2024-05-01"},
                            {"end": "2024-06-30", "val": 1050, "form": "10-Q", "fp": "Q2", "filed": "2024-08-01"},
                        ]
                    }
                },
                "NetCashProvidedByUsedInOperatingActivities": {
                    "units": {
                        "USD": [
                            {"end": "2024-03-31", "val": 25, "form": "10-Q", "fp": "Q1", "filed": "2024-05-01"},
                            {"end": "2024-06-30", "val": 30, "form": "10-Q", "fp": "Q2", "filed": "2024-08-01"},
                        ]
                    }
                },
            }
        }
    }


def test_quarterly_concept_filters_annuals():
    facts = _fake_facts()
    vals = edgar_client._quarterly_concept_values(
        facts, ["Revenues", "SalesRevenueNet"]
    )
    assert pd.Timestamp("2024-12-31") in vals
    # Annual FY rollup of 460 should not have replaced the Q4 10-K value 130
    assert vals[pd.Timestamp("2024-12-31")] == 130


def test_quarterly_concept_uses_first_matching():
    facts = _fake_facts()
    vals = edgar_client._quarterly_concept_values(
        facts, ["NonExistent", "Revenues"]
    )
    assert vals[pd.Timestamp("2024-03-31")] == 100


def test_quarterly_concept_restatement_picks_latest():
    facts = _fake_facts()
    vals = edgar_client._quarterly_concept_values(facts, ["NetIncomeLoss"])
    assert vals[pd.Timestamp("2024-06-30")] == 19  # later restatement, not 18


def test_statement_from_concepts_yfinance_shape(monkeypatch):
    # Patch company_facts() so we don't hit the network
    monkeypatch.setattr(edgar_client, "company_facts", lambda ticker, force=False: _fake_facts())

    is_df = edgar_client.income_statement("FAKE")
    assert "line" in is_df.columns
    rev_row = is_df[is_df["line"] == "Total Revenue"]
    assert not rev_row.empty
    rev_row = rev_row.drop(columns=["line"]).iloc[0]
    rev_row = rev_row.dropna()
    assert len(rev_row) == 4  # 4 quarters

    bs_df = edgar_client.balance_sheet("FAKE")
    ta_row = bs_df[bs_df["line"] == "Total Assets"]
    assert not ta_row.empty
    ta_row = ta_row.drop(columns=["line"]).iloc[0].dropna()
    assert len(ta_row) == 2

    cf_df = edgar_client.cashflow("FAKE")
    ocf_row = cf_df[cf_df["line"] == "Operating Cash Flow"]
    assert not ocf_row.empty
    ocf_row = ocf_row.drop(columns=["line"]).iloc[0].dropna()
    assert len(ocf_row) == 2


def test_statement_from_concepts_empty_when_no_facts(monkeypatch):
    monkeypatch.setattr(edgar_client, "company_facts", lambda ticker, force=False: None)
    assert edgar_client.income_statement("EMPTY").empty
    assert edgar_client.balance_sheet("EMPTY").empty
    assert edgar_client.cashflow("EMPTY").empty
