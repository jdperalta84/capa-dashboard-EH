import json
from pathlib import Path
import unittest

import pandas as pd
import investigation_exclusions
from investigation_exclusions import apply_investigation_exclusions


class InvestigationExclusionsTest(unittest.TestCase):
    def setUp(self):
        self.rows = pd.DataFrame([
            {"Location": "Latvia", "Number": "CAPA0010", "Date closed": "09/01/2026"},
            {"Location": "Another location", "Number": "CAPA0010", "Date closed": "09/01/2026"},
            {"Location": " latvia ", "Number": " capa0010 ", "Date closed": "09/01/2026"},
        ], index=[0, 0, 1])

    def test_off_preserves_population_and_source_values(self):
        original = self.rows.copy(deep=True)
        included, excluded = apply_investigation_exclusions(self.rows)
        pd.testing.assert_frame_equal(included, original)
        pd.testing.assert_frame_equal(self.rows, original)
        self.assertTrue(excluded.empty)

    def test_location_number_match_and_multiple_reasons(self):
        included, excluded = apply_investigation_exclusions(self.rows, True)
        self.assertEqual(included["Location"].tolist(), ["Another location"])
        self.assertEqual(len(excluded), 2)
        self.assertTrue(excluded["Investigation reason"].str.contains("Closure before notification").all())
        self.assertEqual(excluded["Date closed"].tolist(), ["09/01/2026"] * 2)

    def test_empty_population_and_no_match(self):
        included, excluded = apply_investigation_exclusions(self.rows.iloc[:0], True)
        self.assertTrue(included.empty and excluded.empty)
        included, excluded = apply_investigation_exclusions(self.rows.iloc[1:2], True)
        self.assertEqual(len(included), 1)
        self.assertTrue(excluded.empty)

    def test_review_has_45_distinct_capas_and_49_reasons(self):
        review = json.loads(Path(investigation_exclusions.__file__).with_name(
            "investigation_exclusions.json"
        ).read_text(encoding="utf-8"))
        records = review["records"]
        self.assertEqual(len(records), 45)
        self.assertEqual(len({(r["location"], r["number"]) for r in records}), 45)
        self.assertEqual(sum(len(r["reasons"]) for r in records), 49)


if __name__ == "__main__":
    unittest.main()
