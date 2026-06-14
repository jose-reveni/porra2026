import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import generate_dashboard as gd


@pytest.fixture(scope="module")
def workbook_data():
    xlsx = Path(__file__).resolve().parent.parent / "Porra_Admin_v4_EN.xlsx"
    return gd.parse_workbook(str(xlsx))


@pytest.fixture(scope="module")
def computed_data(workbook_data):
    return gd.compute(workbook_data)


class TestMatchDates:
    def test_match_dates_dict_exists(self):
        assert hasattr(gd, "MATCH_DATES")

    def test_match_dates_has_72_entries(self):
        assert len(gd.MATCH_DATES) == 72

    def test_match_dates_all_codes_present(self, workbook_data):
        match_codes = {m["code"] for m in workbook_data["matches"]}
        assert match_codes == set(gd.MATCH_DATES.keys())

    def test_match_dates_format(self):
        import re
        pattern = re.compile(r"^\d{4}-\d{2}-\d{2}$")
        for code, date in gd.MATCH_DATES.items():
            assert pattern.match(date), f"{code} has invalid date format: {date}"


class TestMatchesIncludeDate:
    def test_each_match_has_date_field(self, workbook_data):
        for m in workbook_data["matches"]:
            assert "date" in m, f"Match {m['code']} missing 'date' field"

    def test_match_date_matches_dict(self, workbook_data):
        for m in workbook_data["matches"]:
            assert m["date"] == gd.MATCH_DATES[m["code"]]


class TestTodayData:
    def test_today_key_exists(self, computed_data):
        assert "today" in computed_data

    def test_today_has_matches_field(self, computed_data):
        assert "matches" in computed_data["today"]

    def test_today_has_all_72_matches(self, computed_data):
        assert len(computed_data["today"]["matches"]) == 72

    def test_today_match_has_picks_per_participant(self, computed_data, workbook_data):
        m = computed_data["today"]["matches"][0]
        assert "picks" in m
        assert len(m["picks"]) == workbook_data["n"]
        for pick in m["picks"]:
            assert "name" in pick
            assert "home" in pick
            assert "away" in pick

    def test_today_match_has_analytics(self, computed_data):
        m = computed_data["today"]["matches"][0]
        assert "outcome_dist" in m
        assert "modal_scoreline" in m
        assert "most_unique_pick" in m

    def test_today_match_has_date(self, computed_data):
        for m in computed_data["today"]["matches"]:
            assert "date" in m
            assert len(m["date"]) == 10


class TestRecentResultsData:
    def test_recent_results_key_exists(self, computed_data):
        assert "recent_results" in computed_data

    def test_recent_results_has_matches_and_total(self, computed_data):
        recent = computed_data["recent_results"]
        assert "matches" in recent
        assert "total" in recent

    def test_recent_results_are_limited_to_latest_six(self, computed_data):
        recent = computed_data["recent_results"]
        assert len(recent["matches"]) <= 6
        dates = [m["date"] for m in recent["matches"]]
        assert dates == sorted(dates, reverse=True)

    def test_recent_result_has_final_score_and_outcome_groups(self, computed_data, workbook_data):
        recent = computed_data["recent_results"]["matches"]
        if not recent:
            pytest.skip("Workbook has no real results loaded")

        m = recent[0]
        assert "result" in m
        assert {"home", "away", "outcome"} <= set(m["result"])
        assert m["result"]["outcome"] in {"1", "X", "2"}
        assert {"exact", "sign", "miss"} <= set(m)
        assert len(m["exact"]) + len(m["sign"]) + len(m["miss"]) == workbook_data["n"]
