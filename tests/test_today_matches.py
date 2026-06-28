import sys
import subprocess
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import generate_dashboard as gd


@pytest.fixture(scope="module")
def workbook_data():
    xlsx = Path(__file__).resolve().parent.parent / "Porra_Admin_v5_EN.xlsx"
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


class TestTodayTrivia:
    def test_no_repeated_trivia_for_same_team(self, computed_data):
        from collections import defaultdict

        trivia_by_team = defaultdict(list)
        for m in computed_data["today"]["matches"]:
            trivia_by_team[m["home_en"]].append(m["home_trivia"]["es"])
            trivia_by_team[m["away_en"]].append(m["away_trivia"]["es"])

        for team, facts in trivia_by_team.items():
            non_empty = [f for f in facts if f]
            assert len(non_empty) == len(set(non_empty)), (
                f"Team {team} has repeated trivia facts"
            )


class TestTodayMatchdayDate:
    def test_spanish_early_morning_uses_previous_matchday(self):
        js = _extract_js_function(gd.JS, "matchdayDateStr")
        script = f"""
{js}
const d = new Date(2026, 5, 15, 2, 0, 0);
process.stdout.write(matchdayDateStr(d));
"""
        result = subprocess.check_output(["node", "-e", script], text=True)

        assert result == "2026-06-14"

    def test_spanish_morning_uses_current_matchday(self):
        js = _extract_js_function(gd.JS, "matchdayDateStr")
        script = f"""
{js}
const d = new Date(2026, 5, 15, 6, 0, 0);
process.stdout.write(matchdayDateStr(d));
"""
        result = subprocess.check_output(["node", "-e", script], text=True)

        assert result == "2026-06-15"


def _extract_js_function(source, name):
    marker = f"function {name}"
    start = source.find(marker)
    assert start != -1, f"{name} is missing from dashboard JS"
    depth = 0
    for i in range(start, len(source)):
        if source[i] == "{":
            depth += 1
        elif source[i] == "}":
            depth -= 1
            if depth == 0:
                return source[start:i + 1]
    raise AssertionError(f"{name} JS function is incomplete")


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


class TestLiveProgressionData:
    def test_live_progression_tracks_every_played_match(self, computed_data):
        live = computed_data["live"]
        if not live:
            pytest.skip("Workbook has no real results loaded")

        assert "progression" in live
        assert live["steps"] == len(live["progression"])
        assert len(live["progression"]) == live["played"] + (
            1 if live.get("standings_ready") else 0
        )
        last = live["progression"][-1]
        if live.get("standings_ready"):
            assert last.get("virtual") is True
            assert last.get("kind") == "standings"
            assert last["table"][0]["pts"] == live["table"][0]["pts"]
        assert {"group_pts", "standings_pts", "thirds_pts", "ko_pts"} <= set(
            live["table"][0]
        )

    def test_live_progression_rows_have_rank_and_delta(self, computed_data, workbook_data):
        live = computed_data["live"]
        if not live:
            pytest.skip("Workbook has no real results loaded")

        first_snapshot = live["progression"][0]
        assert len(first_snapshot["table"]) == workbook_data["n"]
        for row in first_snapshot["table"]:
            assert {
                "name",
                "pts",
                "exact",
                "sign",
                "rank",
                "delta",
                "round_pts",
                "round_exact",
                "round_sign",
            } <= set(row)


class TestKnockoutData:
    def test_knockout_predictions_are_parsed(self, workbook_data):
        assert "knockouts" in workbook_data
        rounds = workbook_data["knockouts"]["rounds"]
        assert [r["key"] for r in rounds] == ["r32", "r16", "qf", "sf"]
        assert [len(r["matches"]) for r in rounds] == [16, 8, 4, 2]

    def test_knockout_schedule_is_attached(self, computed_data):
        first = computed_data["knockout"]["rounds"][0]["matches"][0]
        assert first["code"] == "R32-M1"
        assert first["fixture_home"] == "Sudáfrica"
        assert first["fixture_away"] == "Canadá"
        assert first["date"] == "2026-06-28"
        assert first["time_es"] == "21:00"
        assert first["time_uk"] == "20:00"
        assert first["venue"] == "Los Angeles Stadium"

    def test_knockout_key_exists_in_computed_data(self, computed_data):
        knockout = computed_data["knockout"]
        assert {"ready", "filled", "total", "rounds", "outright", "awards", "scoring"} <= set(knockout)

    def test_knockout_consensus_handles_empty_template(self, computed_data):
        champion = computed_data["knockout"]["outright"]["champion"]
        assert champion["value"] is None
        assert champion["agreement"] == 0

    def test_knockout_coverage_counts_divergent_filled_picks(self):
        data = {
            "names": ["A", "B"],
            "n": 2,
            "knockout_results": {"matches": {}, "outright": {}, "awards": {}},
            "knockouts": {
                "rounds": [{
                    "key": "r32",
                    "label_es": "Dieciseisavos",
                    "label_en": "Round of 32",
                    "advance_points": 1,
                    "matches": [{
                        "code": "R32-M1",
                        "score_picks": [(1, 0), (0, 1)],
                        "penalty_picks": ["México", "Sudáfrica"],
                        "winner_picks": ["México", "Sudáfrica"],
                    }],
                }],
                "final_matches": [],
                "outright": {
                    "champion": {
                        "label_es": "Campeón",
                        "label_en": "Champion",
                        "points": 12,
                        "picks": ["México", "Sudáfrica"],
                    }
                },
                "awards": {},
            },
        }

        knockout = gd.compute_knockout(data)

        assert knockout["filled"] == 8
        assert knockout["total"] == 8
        assert knockout["pct"] == 100.0

    def test_knockout_dashboard_section_is_registered(self):
        assert "buildEliminatorias" in gd.JS
        assert "todayScheduleMatches" in gd.JS
        assert "eliminatorias" in gd.JS
