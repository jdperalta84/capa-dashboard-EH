"""Optional CAPA exclusions approved for the 22 October 2026 investigation."""

import json
from pathlib import Path


def apply_investigation_exclusions(capas, enabled=False):
    """Return included rows and an audit copy of omitted rows with reasons.

    Apply after location selection and before any metrics or charts. Source
    values and dates are preserved. A number only matches within its location.
    """
    if not enabled:
        excluded = capas.iloc[:0].copy()
        excluded["Investigation reason"] = []
        return capas.copy(), excluded

    review = json.loads(
        Path(__file__).with_name("investigation_exclusions.json").read_text(
            encoding="utf-8"
        )
    )
    normalize = lambda value: str(value).strip().casefold()
    reasons = {
        (normalize(row["location"]), normalize(row["number"])): "; ".join(row["reasons"])
        for row in review["records"]
    }
    row_reasons = [
        reasons.get((normalize(location), normalize(number)), "")
        for location, number in zip(capas["Location"], capas["Number"])
    ]
    mask = [bool(reason) for reason in row_reasons]
    excluded = capas.loc[mask].copy()
    excluded["Investigation reason"] = [reason for reason in row_reasons if reason]
    return capas.loc[[not match for match in mask]].copy(), excluded
