from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Band:
    """A 0-10 score band. (lower, upper, score). Inclusive of lower bound."""
    lower: float
    upper: float
    score: int


HIGHER_IS_BETTER = [
    Band(-float("inf"), 0.00, 1),
    Band(0.00, 0.05, 3),
    Band(0.05, 0.10, 5),
    Band(0.10, 0.15, 7),
    Band(0.15, 0.20, 8),
    Band(0.20, 0.30, 9),
    Band(0.30, float("inf"), 10),
]

LOWER_IS_BETTER = [
    Band(-float("inf"), 0.00, 10),
    Band(0.00, 0.20, 9),
    Band(0.20, 0.40, 8),
    Band(0.40, 0.60, 6),
    Band(0.60, 1.00, 4),
    Band(1.00, 2.00, 2),
    Band(2.00, float("inf"), 1),
]

CURRENT_RATIO_BANDS = [
    Band(-float("inf"), 0.5, 1),
    Band(0.5, 0.8, 3),
    Band(0.8, 1.0, 5),
    Band(1.0, 1.5, 7),
    Band(1.5, 2.5, 9),
    Band(2.5, 5.0, 10),
    Band(5.0, float("inf"), 7),  # too much cash sitting also penalized
]

INTEREST_COVER_BANDS = [
    Band(-float("inf"), 1.0, 1),
    Band(1.0, 2.0, 3),
    Band(2.0, 4.0, 5),
    Band(4.0, 8.0, 7),
    Band(8.0, 15.0, 9),
    Band(15.0, float("inf"), 10),
]

DSO_BANDS = [
    Band(-float("inf"), 30.0, 10),
    Band(30.0, 45.0, 9),
    Band(45.0, 60.0, 7),
    Band(60.0, 90.0, 5),
    Band(90.0, 120.0, 3),
    Band(120.0, float("inf"), 1),
]


def score(value: float | None, bands: list[Band]) -> int | None:
    if value is None or value != value:
        return None
    for b in bands:
        if b.lower <= value < b.upper:
            return b.score
    return None
