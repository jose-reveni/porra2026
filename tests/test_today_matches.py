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
        assert len(m["exact"]) + len(m["sign"]) + len(m["miss"]) == workbook_data["n"]

    def test_recent_results_include_knockout_matches(self, computed_data, workbook_data):
        recent = computed_data["recent_results"]["matches"]
        ko_match = next((m for m in recent if m["code"] == "R32-M3"), None)

        assert ko_match is not None
        assert ko_match["is_knockout"] is True
        assert ko_match["phase_es"] == "Dieciseisavos"
        assert ko_match["result"]["home"] == 0
        assert ko_match["result"]["away"] == 1
        assert ko_match["result"]["winner"] == "Canadá"
        assert {"exact", "sign", "miss", "advance"} <= set(ko_match)
        assert (
            len(ko_match["exact"]) + len(ko_match["sign"]) + len(ko_match["miss"])
            == workbook_data["n"]
        )


class TestLiveProgressionData:
    def test_live_progression_tracks_every_played_match(self, computed_data):
        live = computed_data["live"]
        if not live:
            pytest.skip("Workbook has no real results loaded")

        assert "progression" in live
        assert live["steps"] == len(live["progression"])
        virtual_steps = (1 if live.get("standings_ready") else 0)
        if computed_data["knockout"].get("scoring"):
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

        ko_step = live["progression"][-1]
        assert ko_step["code"] == "R32-M3"
        assert ko_step["kind"] == "ko"
        assert ko_step["label_es"] == "Eliminatorias"
        assert ko_step["phase_es"] == "Dieciseisavos"
        assert ko_step["home"] == "Sudáfrica"
        assert ko_step["away"] == "Canadá"
        assert ko_step["result"] == {"home": 0, "away": 1, "outcome": "2"}
        assert {"ko_pts", "round_ko_pts", "round_exact", "round_sign", "round_advance"} <= set(
            ko_step["table"][0]
        )
        assert ko_step["table"][0]["pts"] == live["table"][0]["pts"]

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

    def test_knockout_key_exists_in_computed_data(self, computed_data):
        knockout = computed_data["knockout"]
        assert {"ready", "filled", "total", "results_started", "rounds", "outright", "awards", "scoring", "metrics"} <= set(knockout)
        assert knockout["results_started"] is True
        assert knockout["scoring"]["played"] == 1
        metrics = knockout["metrics"]
        assert metrics is not None
        assert metrics["champRank"][0]["team"] == "España"
        assert metrics["champRank"][0]["count"] == 8
        assert len(metrics["people"]) == len(computed_data["cards"])

    def test_knockout_consensus_shape(self, computed_data):
        champion = computed_data["knockout"]["outright"]["champion"]
        assert {"value", "count", "agreement", "dist"} <= set(champion)
        assert champion["agreement"] >= 0

    def test_knockout_round_of_32_winner_comes_from_score(self, workbook_data):
        match = workbook_data["knockouts"]["rounds"][0]["matches"][4]
        assert match["code"] == "R32-M5"
        assert match["fixture_home"] == "Portugal"
        assert match["winner_picks"][0] == "Portugal"

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
        assert "koProgressionCard" in gd.JS

    def test_nav_badges_use_match_counts_not_participants(self):
        assert "if(key === 'groups') return D.hero.matches" in gd.JS
        assert "D.recent_results && D.recent_results.total" in gd.JS
        assert "if(key === 'groups') return N" not in gd.JS
        assert "D.today.matches.length" not in gd.JS

    def test_knockout_matches_expose_picks_and_stake(self, computed_data):
        unplayed = next(
            m for rnd in computed_data["knockout"]["rounds"] for m in rnd["matches"]
            if m["code"] == "R32-M9"
        )
        assert "picks" in unplayed
        assert len(unplayed["picks"]) == computed_data["hero"]["participants"]
        assert "outcome_dist" in unplayed
        assert "stake" in unplayed
        assert unplayed["stake"]["max_one"] == 6
        assert unplayed["stake"]["picks"] > 0
        assert unplayed["stake"]["max_swing"] <= 10

    def test_ko_match_stake_uses_live_ranking(self, computed_data):
        unplayed = next(
            m for rnd in computed_data["knockout"]["rounds"] for m in rnd["matches"]
            if m["code"] == "R32-M9"
        )
        nadia_stake = next(
            p for p in unplayed["stake"]["people"] if p["name"] == "Nadia"
        )
        assert nadia_stake["swing_up"] < 10
        assert "swing_down" in nadia_stake
        assert unplayed["stake"]["min_swing"] >= 0

    def test_knockout_progression_after_r32_m3(self, computed_data):
        prog = computed_data["knockout"]["progression"]
        assert prog is not None
        assert prog["played"] == 1
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
        match = workbook_data["knockouts"]["rounds"][0]["matches"][2]
        result = workbook_data["knockout_results"]["matches"]["R32-M3"]
        winner_key = gd._cmp_team(result["winner"])
        ok_r32 = miss_r32 = drift_r16 = 0
        for i, name in enumerate(workbook_data["names"]):
            person = next(p for p in metrics["people"] if p["name"] == name)
            pick = match["winner_picks"][i]
            r32 = person["bracket"][0]
            r16 = person["bracket"][1]
            assert r32["total"] == 16
            assert r16["total"] == 8
            assert len(person["bracket"]) == 5
            if gd._cmp_team(pick) == winner_key:
                ok_r32 += 1
                assert r32["hits"] == 1 and r32["misses"] == 0
                assert r16["drift"] == 0
            else:
                miss_r32 += 1
                assert r32["hits"] == 0 and r32["misses"] == 1
                assert r16["drift"] >= 1
                drift_r16 += 1
        assert ok_r32 == 20
        assert miss_r32 == 8
        assert drift_r16 == 8

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
        data = workbook_data
        c = gd.compute(data)
        brasil = next(m for rnd in c["knockout"]["rounds"] for m in rnd["matches"] if m["code"] == "R32-M9")
        germany = next(m for rnd in c["knockout"]["rounds"] for m in rnd["matches"] if m["code"] == "R32-M1")
        assert brasil["stake"] and not brasil["stake"].get("deferred")
        assert germany["stake"]["deferred"] is True
        assert germany["stake"]["pending_after"][0]["code"] == "R32-M9"

    def test_same_day_later_stake_updates_after_earlier_result(self, workbook_data):
        import copy
        data = copy.deepcopy(workbook_data)
        data["knockout_results"]["matches"]["R32-M9"] = {
            "score": (2, 1),
            "winner": "Brasil",
        }
        c = gd.compute(data)
        brasil = next(m for rnd in c["knockout"]["rounds"] for m in rnd["matches"] if m["code"] == "R32-M9")
        germany = next(m for rnd in c["knockout"]["rounds"] for m in rnd["matches"] if m["code"] == "R32-M1")
        assert brasil.get("stake") is None
        assert germany["stake"] and not germany["stake"].get("deferred")
        assert germany["stake"]["max_swing"] > 0

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
