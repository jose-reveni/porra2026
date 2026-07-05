import copy
import re
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

    def test_r32_matches_have_fourth_trivia_fact(self, computed_data):
        r32 = computed_data["knockout"]["rounds"][0]["matches"]
        for m in r32:
            assert m["home_trivia"]["es"], f"{m['code']} home trivia missing"
            assert m["away_trivia"]["es"], f"{m['code']} away trivia missing"
            home_en = m["fixture_home_en"]
            away_en = m["fixture_away_en"]
            assert m["home_trivia"]["es"] == gd.TRIVIA[home_en][3][0]
            assert m["away_trivia"]["es"] == gd.TRIVIA[away_en][3][0]


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


class TestMeTodayOutcome:
    def test_knockout_result_score_counts_exact_pick(self):
        js = "\n".join([
            _extract_js_function(gd.JS, "mePickInMatch"),
            _extract_js_function(gd.JS, "meTodayOutcome"),
        ])
        script = f"""
let ME = 'JE';
{js}
const match = {{
  result: {{score: {{home: 2, away: 1}}}},
  picks: [{{name: 'JE', home: 2, away: 1}}],
}};
process.stdout.write(meTodayOutcome(match).kind);
"""
        result = subprocess.check_output(["node", "-e", script], text=True)

        assert result == "exact"


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


class TestTodayStake:
    def test_played_group_match_has_no_stake(self, computed_data):
        today = computed_data["today"]["matches"]
        if not today:
            pytest.skip("No today matches in workbook")
        played = [m for m in today if m.get("stake") is None]
        assert played, "Expected played matches without stake"

    def test_compute_match_stake_for_unplayed_match(self):
        match = {
            "code": "GX-M1",
            "picks": [
                {"name": "A", "home": 1, "away": 0},
                {"name": "B", "home": 0, "away": 1},
            ],
        }
        live_table = [
            {"name": "A", "pts": 10, "rank": 2},
            {"name": "B", "pts": 12, "rank": 1},
        ]
        stake = gd.compute_match_stake(match, {}, live_table)
        assert stake["max_swing"] == 1
        assert stake["min_swing"] == 1
        assert stake["max_points"] == 8
        assert stake["picks"] == 2

    def test_scenario_stake_unanimous_pick_no_movement(self):
        names = ["A", "B", "C"]
        live_table = [{"name": n, "pts": 100, "rank": i + 1} for i, n in enumerate(names)]
        match = {
            "code": "G1",
            "picks": [{"name": n, "home": 2, "away": 0} for n in names],
        }
        stake = gd.compute_match_stake(match, {}, live_table)
        assert stake["max_swing"] == 0
        assert stake["min_swing"] == 0
        assert all(p["swing_up"] == 0 and p["swing_down"] == 0 for p in stake["people"])

    def test_scenario_stake_sign_only_beats_you(self):
        """Si aciertas signo pero otro hace pleno, puedes bajar."""
        live_table = [
            {"name": "A", "pts": 100, "rank": 1},
            {"name": "B", "pts": 99, "rank": 2},
        ]
        match = {
            "code": "G1",
            "picks": [
                {"name": "A", "home": 2, "away": 0},
                {"name": "B", "home": 1, "away": 0},
            ],
        }
        stake = gd.compute_match_stake(match, {}, live_table)
        a, b = {p["name"]: p for p in stake["people"]}["A"], {p["name"]: p for p in stake["people"]}["B"]
        assert a["swing_up"] == 0
        assert a["swing_down"] == 1
        assert a["worst_result"]["score"] == "1-0"
        assert b["swing_up"] == 1
        assert b["swing_down"] == 0
        assert b["best_result"]["score"] == "1-0"
        assert stake["max_swing_result"]["score"] == "1-0"
        assert stake["min_swing_result"]["score"] == "1-0"

    def test_stake_section_removed_from_js(self):
        assert "buildStakeToday" not in gd.JS
        assert "koStakeHtml" not in gd.JS
        assert "stakeSwingHtml" in gd.JS
        assert "stakeResultLbl" in gd.JS


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
        # KO knockout cards add a `voided` bucket (right outcome but wrong branch),
        # so every participant lands in exactly one of the four groups.
        voided = m.get("voided", [])
        assert len(m["exact"]) + len(m["sign"]) + len(m["miss"]) + len(voided) == workbook_data["n"]

    def test_recent_results_include_knockout_matches(self, computed_data, workbook_data):
        recent = computed_data["recent_results"]["matches"]
        ko_matches = [m for m in recent if m.get("is_knockout")]

        assert ko_matches, "recent results should include at least one knockout match"
        ko_match = ko_matches[0]
        assert ko_match["phase_es"]
        assert {"home", "away", "outcome"} <= set(ko_match["result"])
        assert ko_match["result"]["winner"]
        assert {"exact", "sign", "miss", "advance"} <= set(ko_match)
        # Cada participante cae en exactamente un grupo (incluido el "cementerio").
        buckets = ("exact", "sign", "miss", "voided")
        assert sum(len(ko_match.get(b, [])) for b in buckets) == workbook_data["n"]


class TestLiveProgressionData:
    def test_live_progression_tracks_every_played_match(self, computed_data):
        live = computed_data["live"]
        if not live:
            pytest.skip("Workbook has no real results loaded")

        assert "progression" in live
        assert live["steps"] == len(live["progression"])
        virtual_steps = (1 if live.get("standings_ready") else 0)
        ko_progression = computed_data["knockout"].get("progression")
        if ko_progression:
            virtual_steps += ko_progression["steps"]
        elif computed_data["knockout"].get("scoring"):
            virtual_steps += 1
        assert len(live["progression"]) == live["played"] + virtual_steps
        last = live["progression"][-1]
        if computed_data["knockout"].get("scoring"):
            assert last.get("virtual") is True
            assert last.get("kind") == "ko"
            assert last["table"][0]["pts"] == live["table"][0]["pts"]
        elif live.get("standings_ready"):
            assert last.get("virtual") is True
            assert last.get("kind") == "standings"
            assert last["table"][0]["pts"] == live["table"][0]["pts"]
        assert {"group_pts", "standings_pts", "thirds_pts", "ko_pts"} <= set(
            live["table"][0]
        )

    def test_live_progression_includes_knockout_step(self, computed_data):
        live = computed_data["live"]
        if not live or not computed_data["knockout"].get("scoring"):
            pytest.skip("Workbook has no knockout results loaded")

        latest_ko = computed_data["knockout"]["progression"]["progression"][-1]
        ko_step = live["progression"][-1]
        assert ko_step["code"] == latest_ko["code"]
        assert ko_step["kind"] == "ko"
        assert ko_step["label_es"] == "Eliminatorias"
        assert ko_step["phase_es"] == latest_ko["phase_es"]
        assert ko_step["home"] == latest_ko["home"]
        assert ko_step["away"] == latest_ko["away"]
        assert ko_step["result"] == latest_ko["result"]
        assert {"ko_pts", "round_ko_pts", "round_exact", "round_sign", "round_advance"} <= set(
            ko_step["table"][0]
        )
        assert ko_step["table"][0]["pts"] == live["table"][0]["pts"]

    def test_live_progression_keeps_knockout_matches_separate(self, computed_data):
        live = computed_data["live"]
        ko_progression = computed_data["knockout"]["progression"]
        if not live or not ko_progression or ko_progression["steps"] < 2:
            pytest.skip("Workbook needs at least two knockout results loaded")

        ko_steps = ko_progression["progression"]
        ko_codes = [step["code"] for step in ko_steps]
        assert ko_codes[:2] == ["R32-M3", "R32-M9"]
        live_ko_steps = [
            step for step in live["progression"]
            if step.get("kind") == "ko" and step["code"] in ko_codes
        ]
        assert [step["code"] for step in live_ko_steps] == ko_codes

        latest_ko = ko_steps[-1]
        carryover = next(
            row for row in latest_ko["table"]
            if row["pts"] != row["round_pts"]
        )
        live_row = next(
            row for row in live_ko_steps[-1]["table"]
            if row["name"] == carryover["name"]
        )
        assert live_row["ko_pts"] == carryover["pts"]
        assert live_row["round_ko_pts"] == carryover["round_pts"]

    def test_race_hover_shows_standings_and_knockout_sources(self):
        assert "function buildLiveRanking" in gd.JS
        assert "function rankTip" in gd.JS
        assert "Fase de grupos" in gd.JS
        assert "Eliminatorias" in gd.JS
        assert "ko_exact" in gd.JS
        assert "ko_outcomes" in gd.JS
        assert "ko_advance" in gd.JS
        assert "RANKING_VARIANTS" not in gd.JS
        assert "race-feed" not in gd.JS

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
        assert first["fixture_home"] == "Alemania"
        assert first["fixture_away"] == "Paraguay"
        assert first["date"] == "2026-06-29"
        assert first["time_es"] == "22:30"
        assert first["time_uk"] == "21:30"
        assert first["venue"] == "Boston Stadium"

    def test_knockout_key_exists_in_computed_data(self, computed_data, workbook_data):
        knockout = computed_data["knockout"]
        assert {"ready", "filled", "total", "results_started", "rounds", "outright", "awards", "scoring", "metrics"} <= set(knockout)
        assert knockout["results_started"] is True
        expected_played = sum(
            1 for result in workbook_data["knockout_results"]["matches"].values()
            if "score" in result
        )
        assert knockout["scoring"]["played"] == expected_played
        metrics = knockout["metrics"]
        assert metrics is not None
        assert metrics["champRank"][0]["team"] == "España"
        assert metrics["champRank"][0]["count"] > 0
        assert len(metrics["people"]) == len(computed_data["cards"])

    def test_knockout_consensus_shape(self, computed_data):
        champion = computed_data["knockout"]["outright"]["champion"]
        assert {"value", "count", "agreement", "dist"} <= set(champion)
        assert champion["agreement"] >= 0

    def test_knockout_round_of_32_winner_comes_from_score(self, workbook_data):
        match = workbook_data["knockouts"]["rounds"][0]["matches"][4]
        assert match["code"] == "R32-M5"
        # R32 has real teams from the start, so the winner pick is derived from the
        # scoreline: a decisive score picks that side; it is always one of the two.
        h, a = match["score_picks"][0]
        wp = match["winner_picks"][0]
        sides = {gd._cmp_team(match["fixture_home"]), gd._cmp_team(match["fixture_away"])}
        if h is not None and a is not None and h != a:
            expected = match["fixture_home"] if h > a else match["fixture_away"]
            assert gd._cmp_team(wp) == gd._cmp_team(expected)
        else:
            assert wp is None or gd._cmp_team(wp) in sides

    def test_knockout_result_is_exposed_for_live_view(self, computed_data):
        match = computed_data["knockout"]["rounds"][0]["matches"][2]
        assert match["code"] == "R32-M3"
        assert match["result"] == {
            "score": {"home": 0, "away": 1},
            "winner": "Canadá",
            "winner_flag": "🇨🇦",
        }

    def test_knockout_bracket_resolves_winner_into_next_round(self, computed_data):
        r32m3 = computed_data["knockout"]["rounds"][0]["matches"][2]
        r16m2 = computed_data["knockout"]["rounds"][1]["matches"][1]
        assert r32m3["fixture_home"] == "Sudáfrica"
        assert r32m3["fixture_away"] == "Canadá"
        assert r16m2["fixture_home"] == "W75"
        assert r16m2["fixture_away"] == "W76"
        assert r16m2["resolved_home"] == "Canadá"
        assert r16m2["resolved_home_flag"] == "🇨🇦"

    def test_knockout_consensus_handles_empty_template(self):
        empty_data = {
            "names": ["A"],
            "n": 1,
            "knockout_results": {"matches": {}, "outright": {}, "awards": {}},
            "knockouts": {
                "rounds": [],
                "final_matches": [],
                "outright": {
                    "champion": {
                        "label_es": "Campeón",
                        "label_en": "Champion",
                        "points": 12,
                        "picks": [None],
                    }
                },
                "awards": {},
            },
        }
        knockout = gd.compute_knockout(empty_data)
        champion = knockout["outright"]["champion"]
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
                        "date": "2026-06-29",
                        "dt": "2026-06-29T20:00:00",
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

        assert knockout["filled"] == 6
        assert knockout["total"] == 6
        assert knockout["pct"] == 100.0

    def test_knockout_dashboard_section_is_registered(self):
        assert "buildEliminatorias" in gd.JS
        assert "todayScheduleMatches" in gd.JS
        assert "eliminatorias" in gd.JS
        assert "m.winner.agreement || 0) * 100" not in gd.JS
        assert "bk-cbar split" in gd.JS
        assert "Resultado final" in gd.JS
        assert "koBracketPrecisionCard" in gd.JS

    def test_nav_badges_use_match_counts_not_participants(self):
        assert "if(key === 'groups') return D.hero.matches" in gd.JS
        assert "D.recent_results && D.recent_results.total" in gd.JS
        assert "if(key === 'groups') return N" not in gd.JS
        assert "D.today.matches.length" not in gd.JS

    def test_knockout_matches_expose_picks_and_stake(self, computed_data):
        unplayed = next(
            m for rnd in computed_data["knockout"]["rounds"] for m in rnd["matches"]
            if not m.get("result") and m.get("stake") and not m["stake"].get("deferred")
        )
        assert "picks" in unplayed
        assert len(unplayed["picks"]) == computed_data["hero"]["participants"]
        assert "outcome_dist" in unplayed
        assert "stake" in unplayed
        # max per person = 3 (signo) + 2 (pleno) + puntos de pase de la ronda.
        assert unplayed["stake"]["max_one"] == 5 + unplayed["stake"].get("advance_points", 0)
        assert unplayed["stake"]["picks"] > 0
        assert unplayed["stake"]["max_swing"] >= 0

    def test_ko_match_stake_uses_live_ranking(self, computed_data):
        unplayed = next(
            m for rnd in computed_data["knockout"]["rounds"] for m in rnd["matches"]
            if not m.get("result") and m.get("stake") and not m["stake"].get("deferred")
        )
        nadia_stake = next(
            p for p in unplayed["stake"]["people"] if p["name"] == "Nadia"
        )
        assert nadia_stake["swing_up"] < 10
        assert "swing_down" in nadia_stake
        assert unplayed["stake"]["min_swing"] >= 0

    def test_knockout_progression_after_r32_m3(self, computed_data, workbook_data):
        prog = computed_data["knockout"]["progression"]
        expected_played = sum(
            1 for result in workbook_data["knockout_results"]["matches"].values()
            if "score" in result
        )
        assert prog is not None
        assert prog["played"] == expected_played
        assert prog["progression"][0]["code"] == "R32-M3"
        row = prog["progression"][0]["table"][0]
        assert {"rank", "delta", "round_pts", "round_advance"} <= set(row)

    def test_knockout_champion_survival_not_penalized_for_other_picks(
        self, computed_data, workbook_data,
    ):
        metrics = computed_data["knockout"]["metrics"]
        match = workbook_data["knockouts"]["rounds"][0]["matches"][2]
        result = workbook_data["knockout_results"]["matches"]["R32-M3"]
        winner_key = gd._cmp_team(result["winner"])
        still_alive = []
        for i, name in enumerate(workbook_data["names"]):
            champ = workbook_data["knockouts"]["outright"]["champion"]["picks"][i]
            pick = match["winner_picks"][i]
            if not champ or not pick:
                continue
            if gd._cmp_team(pick) != winner_key and gd._cmp_team(champ) != gd._cmp_team("South Africa"):
                person = next(p for p in metrics["people"] if p["name"] == name)
                still_alive.append(person["fell"])
        assert still_alive
        assert all(fell == len(metrics["rounds"]) for fell in still_alive)

    def test_knockout_champion_survival_falls_when_champion_loses(self):
        data = {
            "names": ["Ana", "Bob"],
            "n": 2,
            "qualifiers": {},
            "knockout_results": {
                "matches": {
                    "R32-M1": {"score": (0, 1), "winner": "Mexico"},
                },
                "outright": {},
                "awards": {},
            },
            "knockouts": {
                "rounds": [{
                    "key": "r32",
                    "label_es": "Dieciseisavos",
                    "label_en": "Round of 32",
                    "advance_points": 1,
                    "matches": [{
                        "code": "R32-M1",
                        "fixture_home": "Mexico",
                        "fixture_away": "South Africa",
                        "score_picks": [(0, 1), (1, 0)],
                        "penalty_picks": ["Mexico", "South Africa"],
                        "winner_picks": ["Mexico", "South Africa"],
                    }],
                }],
                "final_matches": [],
                "outright": {
                    "champion": {
                        "label_es": "Campeón",
                        "label_en": "Champion",
                        "points": 12,
                        "picks": ["South Africa", "Mexico"],
                    }
                },
                "awards": {},
            },
        }
        metrics = gd.compute_knockout_metrics(data)
        ana = next(p for p in metrics["people"] if p["name"] == "Ana")
        bob = next(p for p in metrics["people"] if p["name"] == "Bob")
        assert ana["fell"] == 0
        assert bob["fell"] == len(metrics["rounds"])

    def test_knockout_bracket_precision_by_round(self, computed_data, workbook_data):
        metrics = computed_data["knockout"]["metrics"]
        assert metrics["bracketRounds"]
        assert len(metrics["bracketRounds"]) == 5
        r32_results = workbook_data["knockout_results"]["matches"]
        r32_matches = [
            m for m in workbook_data["knockouts"]["rounds"][0]["matches"]
            if m["code"] in r32_results and "score" in r32_results[m["code"]]
        ]
        total_hits = total_misses = 0
        for i, name in enumerate(workbook_data["names"]):
            person = next(p for p in metrics["people"] if p["name"] == name)
            r32 = person["bracket"][0]
            r16 = person["bracket"][1]
            assert r32["total"] == 16
            assert r16["total"] == 8
            assert len(person["bracket"]) == 5
            expected_hits = sum(
                1 for match in r32_matches
                if gd._cmp_team(match["winner_picks"][i])
                == gd._cmp_team(r32_results[match["code"]]["winner"])
            )
            expected_misses = len(r32_matches) - expected_hits
            total_hits += expected_hits
            total_misses += expected_misses
            assert r32["hits"] == expected_hits
            assert r32["misses"] == expected_misses
            assert r16["drift"] >= 0
        assert total_hits > 0
        assert total_misses > 0
        assert total_hits + total_misses == len(r32_matches) * len(workbook_data["names"])

    def test_knockout_dashboard_has_bracket_precision(self):
        assert "koBracketPrecisionCard" in gd.JS
        assert "koSurvivalCard" not in gd.JS
        match = {
            "score_picks": [(1, 0), (0, 1)],
            "winner_picks": ["México", "Sudáfrica"],
            "fixture_home": "México",
            "fixture_away": "Sudáfrica",
        }
        rnd = {"advance_points": 1}
        live_table = [
            {"name": "A", "pts": 10, "rank": 2},
            {"name": "B", "pts": 12, "rank": 1},
        ]
        stake = gd.compute_ko_match_stake(match, rnd, ["A", "B"], live_table)
        assert stake["max_one"] == 6
        assert stake["picks"] == 2
        assert stake["max_swing"] == 1
        assert stake["min_swing"] == 1
        assert stake["people"][0]["max_pts"] == 6

    def test_same_day_later_match_defers_stake_until_earlier_played(self, workbook_data):
        # R32-M9 y R32-M1 son el mismo día (M9 antes que M1). Con ambos sin jugar,
        # el stake del posterior (M1) se aplaza hasta que caiga el anterior (M9).
        data = copy.deepcopy(workbook_data)
        data["knockout_results"]["matches"].pop("R32-M9", None)
        data["knockout_results"]["matches"].pop("R32-M1", None)
        c = gd.compute(data)
        earlier = next(m for rnd in c["knockout"]["rounds"] for m in rnd["matches"] if m["code"] == "R32-M9")
        later = next(m for rnd in c["knockout"]["rounds"] for m in rnd["matches"] if m["code"] == "R32-M1")
        assert earlier["stake"] and not earlier["stake"].get("deferred")
        assert later["stake"]["deferred"] is True
        assert later["stake"]["pending_after"][0]["code"] == "R32-M9"

    def test_same_day_later_stake_updates_after_earlier_result(self, workbook_data):
        # R32-M9 mantiene su resultado real (jugado); solo el posterior (M1) queda
        # sin jugar, así que su stake ya no se aplaza.
        data = copy.deepcopy(workbook_data)
        data["knockout_results"]["matches"].pop("R32-M1", None)
        c = gd.compute(data)
        earlier = next(m for rnd in c["knockout"]["rounds"] for m in rnd["matches"] if m["code"] == "R32-M9")
        later = next(m for rnd in c["knockout"]["rounds"] for m in rnd["matches"] if m["code"] == "R32-M1")
        assert earlier.get("stake") is None
        assert later["stake"] and not later["stake"].get("deferred")
        assert later["stake"]["max_swing"] > 0

    def test_knockout_phase3_metrics(self, computed_data):
        metrics = computed_data["knockout"]["metrics"]
        assert "honors" in metrics
        assert "championPath" in metrics
        assert "vsPuebloRank" in metrics
        assert metrics["expApprox"] is False
        person = metrics["people"][0]
        assert {"variance", "vsPueblo", "vsPuebloTotal", "boldPct", "reventador", "expApprox"} <= set(person)
        assert 0 <= person["variance"] <= 100
        assert metrics["honors"]["profeta"]["name"]
        assert metrics["honors"]["agorero"]["name"]
        assert metrics["honors"]["manual"]["name"]

    def test_ko_risk_reward_uses_consensus_share(self):
        data = {
            "names": ["Ana", "Bob", "Cara"],
            "n": 3,
            "qualifiers": {},
            "knockout_results": {"matches": {}, "outright": {}, "awards": {}},
            "knockouts": {
                "rounds": [{
                    "key": "r32",
                    "label_es": "Dieciseisavos",
                    "label_en": "Round of 32",
                    "advance_points": 1,
                    "matches": [{
                        "code": "R32-M1",
                        "fixture_home": "Mexico",
                        "fixture_away": "South Africa",
                        "score_picks": [(1, 0), (1, 0), (0, 1)],
                        "winner_picks": ["Mexico", "Mexico", "South Africa"],
                    }],
                }],
                "final_matches": [],
                "outright": {
                    "champion": {
                        "label_es": "Campeón",
                        "label_en": "Champion",
                        "points": 12,
                        "picks": ["Mexico", "Mexico", "South Africa"],
                    }
                },
                "awards": {},
            },
        }
        metrics = gd.compute_knockout_metrics(data)
        ana = next(p for p in metrics["people"] if p["name"] == "Ana")
        cara = next(p for p in metrics["people"] if p["name"] == "Cara")
        assert ana["variance"] < cara["variance"]
        assert ana["exp"] > cara["exp"]
        assert metrics["expApprox"] is True
        assert metrics["honors"]["agorero"]["name"] == "Cara"
        assert metrics["honors"]["manual"]["name"] in ("Ana", "Bob")

    def test_ko_champion_path_follows_consensus(self):
        data = {
            "names": ["A", "B"],
            "n": 2,
            "qualifiers": {},
            "knockout_results": {"matches": {}, "outright": {}, "awards": {}},
            "knockouts": {
                "rounds": [{
                    "key": "r32",
                    "label_es": "Dieciseisavos",
                    "label_en": "Round of 32",
                    "advance_points": 1,
                    "matches": [{
                        "code": "R32-M1",
                        "fixture_home": "Spain",
                        "fixture_away": "France",
                        "score_picks": [(1, 0), (1, 0)],
                        "winner_picks": ["Spain", "Spain"],
                    }],
                }],
                "final_matches": [],
                "outright": {
                    "champion": {
                        "label_es": "Campeón",
                        "label_en": "Champion",
                        "points": 12,
                        "picks": ["Spain", "Spain"],
                    }
                },
                "awards": {},
            },
        }
        metrics = gd.compute_knockout_metrics(data)
        assert metrics["championTeam"] == "España"
        assert len(metrics["championPath"]) == 1
        assert metrics["championPath"][0]["opponent"] == "Francia"

    def test_knockout_dashboard_has_phase3_ui(self):
        assert "koChampionPathHtml" in gd.JS
        assert "koHonorsHtml" in gd.JS
        assert "koVsPuebloHtml" in gd.JS
        assert "attachBracketHovers" in gd.JS
        assert "koBoldBarsHtml" in gd.JS
        assert "Acto 10" in gd.JS

    def test_alive_teams_drops_r16_loser_via_feeder(self, workbook_data):
        """Perdedor de octavos (cruce con feeders W##) debe salir de `alive`.

        Regresión: `_ko_alive_teams` solo descartaba lados con nombre real, así
        que octavos+ (fixtures W##) dejaban vivos a los perdedores. Brasil cae
        0-2 con Noruega en R16-M5 y debe quedar eliminado."""
        alive = gd._ko_alive_teams(workbook_data)
        # Perdedores reales de octavos ya jugados.
        assert "brazil" not in alive
        assert "canada" not in alive
        assert "paraguay" not in alive
        # Los que avanzaron siguen vivos.
        assert {"norway", "france", "morocco"} <= alive

    def test_fallen_champions_include_brazil_backers(self, computed_data, workbook_data):
        """El cementerio (Acto 6) lista a quienes pusieron un campeón ya fuera."""
        metrics = computed_data["knockout"]["metrics"]
        grave_champs = {gd._cmp_team(g["champ"]) for g in metrics["grave"]}
        brazil_backers = [
            workbook_data["names"][i]
            for i, pick in enumerate(
                workbook_data["knockouts"]["outright"]["champion"]["picks"]
            )
            if pick and gd._cmp_team(pick) == "brazil"
        ]
        assert brazil_backers  # el escenario tiene apostantes por Brasil
        assert "brazil" in grave_champs
        grave_names = {g["name"] for g in metrics["grave"]}
        assert set(brazil_backers) <= grave_names
