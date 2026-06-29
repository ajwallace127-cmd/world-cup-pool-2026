#!/usr/bin/env python3
"""
World Cup Pool 2026 — Standings Tracker
=========================================
Reads Google Forms pool entries (exported as .xlsx) and generates
a self-contained HTML leaderboard website with live scoring.

Usage:
    python process_pool.py                          # uses pool_entries.xlsx
    python process_pool.py my_entries.xlsx          # custom filename
    python process_pool.py entries.xlsx --no-fetch  # skip ESPN, use manual data
"""

import json
import sys
import os
import re
import time
import html as html_mod
import unicodedata
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, timedelta, datetime
from collections import defaultdict, OrderedDict

# ============================================================
# CONFIGURATION — Edit these as needed
# ============================================================

CHALLENGE_NAME   = "World Cup Pool 2026"
EXCEL_FILE       = "pool_entries.xlsx"
OUTPUT_FILE      = "index.html"
TOURNAMENT_START = date(2026, 6, 11)   # First game kick-off date
TOURNAMENT_END   = date(2026, 7, 19)   # Final

# ── SCORING ──────────────────────────────────────────────────
WIN_PTS       = 300   # regulation win or ET/PK win
TIE_PTS       = 100   # regulation draw, or ET/PK loss
LOSS_PTS      = 0     # regulation loss
WINNER_BONUS  = 450   # correctly picked the tournament winner
ADV_BONUS_AB  = 100   # group stage advancement bonus, Tiers A & B
ADV_BONUS_CD  = 200   # group stage advancement bonus, Tiers C & D
GOAL_PTS      = 150   # per goal for each goal scorer selection

# ── TIERS ────────────────────────────────────────────────────
TIER_DEFS = {
    'A': ['Spain', 'France', 'England', 'Portugal', 'Brazil', 'Argentina', 'Germany', 'Netherlands'],
    'B': ['Norway', 'Colombia', 'Belgium', 'Morocco', 'USA', 'Switzerland', 'Japan', 'Uruguay',
          'Ecuador', 'Mexico', 'Turkey', 'Croatia', 'Senegal', 'Sweden'],
    'C': ['Austria', 'Scotland', 'Canada', 'Czech Republic', 'Ivory Coast', 'Ghana', 'Egypt',
          'Paraguay', 'Algeria', 'South Korea', 'Tunisia', 'Bosnia', 'Australia'],
    'D': ['Iran', 'DR Congo', 'South Africa', 'Cape Verde', 'Saudi Arabia', 'Panama',
          'Uzbekistan', 'Qatar', 'New Zealand', 'Iraq', 'Haiti', 'Curacao', 'Jordan'],
}
TIER_LABELS = {
    'A': 'Tier A — Favorites',
    'B': 'Tier B — Dark Horses',
    'C': 'Tier C — Longshots',
    'D': "Tier D — We're Glad You Came",
}
TIER_ADVANCE_BONUS = {'A': ADV_BONUS_AB, 'B': ADV_BONUS_AB, 'C': ADV_BONUS_CD, 'D': ADV_BONUS_CD}
TIER_PICKS = {'A': 2, 'B': 4, 'C': 4, 'D': 2}

# Build reverse lookup: normalized team name → tier
_TEAM_TIER: dict[str, str] = {}
for _tier, _teams in TIER_DEFS.items():
    for _t in _teams:
        _TEAM_TIER[_t.lower()] = _tier

# ── IMAGES ───────────────────────────────────────────────────
# Place photos in the images/ subfolder of this repo.
# The HTML references them as relative paths — they work locally
# and on GitHub Pages. Fallback to pitch-green gradient if missing.
# Suggested filenames (save from the World Cup photos):
#   images/hero.jpg         — main banner (e.g. Van Persie header)
#   images/standings_bg.jpg — Standings tab strip
#   images/groups_bg.jpg    — Groups tab strip (e.g. Maradona)
#   images/bracket_bg.jpg   — Bracket tab strip (e.g. Götze goal)
#   images/scorers_bg.jpg   — Scorers tab strip (e.g. Ronaldo running)
#   images/allpicks_bg.jpg  — All Picks tab strip (e.g. England celebrating)

# ── MANUAL OVERRIDES ─────────────────────────────────────────
# Goal scorer override — ESPN game summaries are fetched automatically.
# Only fill this in if the automated fetch is failing for a specific player.
GOAL_SCORER_OVERRIDE: dict = {
    # "Kylian Mbappe": 5,
}

MANUAL_ADVANCED: set = {
    # Group A: Mexico 1st, South Africa 2nd
    "Mexico", "South Africa",
    # Group B: Switzerland 1st, Canada 2nd, Bosnia 3rd (best-3rd)
    "Switzerland", "Canada", "Bosnia",
    # Group C: Brazil 1st, Morocco 2nd
    "Brazil", "Morocco",
    # Group D: United States 1st, Australia 2nd, Paraguay 3rd (best-3rd)
    "United States", "Australia", "Paraguay",
    # Group E: Germany 1st, Ivory Coast 2nd, Ecuador 3rd (best-3rd)
    "Germany", "Ivory Coast", "Ecuador",
    # Group F: Netherlands 1st, Japan 2nd, Sweden 3rd (best-3rd)
    "Netherlands", "Japan", "Sweden",
    # Group G: Belgium 1st, Egypt 2nd
    "Belgium", "Egypt",
    # Group H: Spain 1st, Cape Verde 2nd
    "Spain", "Cape Verde",
    # Group I: France 1st, Norway 2nd, Senegal 3rd (best-3rd)
    "France", "Norway", "Senegal",
    # Group J: Argentina 1st, Austria 2nd, Algeria 3rd (best-3rd)
    "Argentina", "Austria", "Algeria",
    # Group K: Colombia 1st, Portugal 2nd, DR Congo 3rd (best-3rd)
    "Colombia", "Portugal", "DR Congo",
    # Group L: England 1st, Croatia 2nd, Ghana 3rd (best-3rd)
    "England", "Croatia", "Ghana",
}  # All 32 Round of 32 teams — group stage complete

# All 16 eliminated teams: 12 fourth-place + 4 third-place teams that missed best-8
MANUAL_ELIMINATED: set = {
    # 4th-place teams (all 12 groups)
    "Czech Republic", # Group A, 4th
    "Qatar",          # Group B, 4th
    "Haiti",          # Group C, 4th
    "Turkey",         # Group D, 4th
    "Curacao",        # Group E, 4th
    "Tunisia",        # Group F, 4th
    "New Zealand",    # Group G, 4th
    "Saudi Arabia",   # Group H, 4th
    "Iraq",           # Group I, 4th
    "Jordan",         # Group J, 4th
    "Uzbekistan",     # Group K, 4th
    "Panama",         # Group L, 4th
    # 3rd-place teams that missed the best-8
    "South Korea",    # Group A, 3rd (3 pts — not enough)
    "Scotland",       # Group C, 3rd (3 pts — not enough)
    "Iran",           # Group G, 3rd (3 pts — not enough)
    "Uruguay",        # Group H, 3rd (2 pts — not enough)
}

MANUAL_TEAM_OVERRIDE: dict = {}  # Override ESPN data for specific teams

# ── TEAM NAME ALIASES ────────────────────────────────────────
TEAM_ALIASES: dict[str, str] = {
    # Tier A
    "spain":                "spain",
    "france":               "france",
    "england":              "england",
    "portugal":             "portugal",
    "brazil":               "brazil",
    "argentina":            "argentina",
    "germany":              "germany",
    "netherlands":          "netherlands",
    "holland":              "netherlands",
    # Tier B
    "norway":               "norway",
    "colombia":             "colombia",
    "belgium":              "belgium",
    "morocco":              "morocco",
    "usa":                  "united states",
    "us":                   "united states",
    "united states":        "united states",
    "switzerland":          "switzerland",
    "japan":                "japan",
    "uruguay":              "uruguay",
    "ecuador":              "ecuador",
    "mexico":               "mexico",
    "turkey":               "turkey",
    "türkiye":              "turkey",
    "turkiye":              "turkey",
    "croatia":              "croatia",
    "senegal":              "senegal",
    "sweden":               "sweden",
    # Tier C
    "austria":              "austria",
    "scotland":             "scotland",
    "canada":               "canada",
    "czech republic":       "czech republic",
    "czechia":              "czech republic",
    "ivory coast":          "ivory coast",
    "cote d'ivoire":        "ivory coast",
    "côte d'ivoire":        "ivory coast",
    "cote divoire":         "ivory coast",
    "ghana":                "ghana",
    "egypt":                "egypt",
    "paraguay":             "paraguay",
    "algeria":              "algeria",
    "south korea":          "korea republic",
    "korea":                "korea republic",
    "korea republic":       "korea republic",
    "republic of korea":    "korea republic",
    "tunisia":              "tunisia",
    "bosnia":                 "bosnia",
    "bosnia and herzegovina": "bosnia",
    "bosnia & herzegovina":   "bosnia",
    "bosnia-herzegovina":     "bosnia",
    "bosnia herzegovina":     "bosnia",
    "australia":            "australia",
    # Tier D
    "iran":                 "iran",
    "ir iran":              "iran",
    "dr congo":             "dr congo",
    "congo dr":             "dr congo",
    "democratic republic of congo": "dr congo",
    "drc":                  "dr congo",
    "south africa":         "south africa",
    "cape verde":           "cape verde",
    "cape verde islands":   "cape verde",
    "saudi arabia":         "saudi arabia",
    "panama":               "panama",
    "uzbekistan":           "uzbekistan",
    "qatar":                "qatar",
    "new zealand":          "new zealand",
    "iraq":                 "iraq",
    "haiti":                "haiti",
    "curacao":              "curacao",
    "curaçao":              "curacao",
    "jordan":               "jordan",
}

ESPN_SCOREBOARD = "https://site.api.espn.com/apis/site/v2/sports/soccer/fifa.world/scoreboard"
ESPN_SUMMARY    = "https://site.api.espn.com/apis/site/v2/sports/soccer/fifa.world/summary"
ESPN_STANDINGS  = "https://site.api.espn.com/apis/v2/sports/soccer/fifa.world/standings"


# ============================================================
# NAME NORMALIZATION
# ============================================================

def _strip_accents(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )

def _normalize(name: str) -> str:
    name = str(name).strip().lower()
    name = _strip_accents(name)
    name = re.sub(r"[''`ʻʼʹ]", "", name)
    if name in TEAM_ALIASES:
        name = TEAM_ALIASES[name]
    name = re.sub(r"\s+", " ", name).strip()
    return name

def _h(s: str) -> str:
    return html_mod.escape(str(s), quote=True)

def _pretty(raw: str) -> str:
    if not raw or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    return " ".join(w.capitalize() for w in str(raw).strip().split())

def get_tier(team_name: str) -> str:
    norm = _normalize(team_name)
    for t_name, tier in _TEAM_TIER.items():
        if _normalize(t_name) == norm:
            return tier
    for t_name, tier in _TEAM_TIER.items():
        if norm and norm in _normalize(t_name):
            return tier
    return '?'


# ============================================================
# ESPN API
# ============================================================

def _round_category(event: dict) -> str:
    """Map ESPN event to round category: GS / R32 / R16 / QF / SF / 3P / F"""
    rn = ""
    comps = event.get("competitions", [{}])
    if comps:
        for note in comps[0].get("notes", []):
            h = note.get("headline", "")
            if h:
                rn = h.lower()
                break
    if not rn:
        rn = (event.get("name", "") + " " + event.get("shortName", "")).lower()

    if any(k in rn for k in ["3rd place", "third place", "bronze"]):
        return "3P"
    if "final" in rn and "semi" not in rn and "quarter" not in rn:
        return "F"
    if any(k in rn for k in ["semi", "semifinal"]):
        return "SF"
    if any(k in rn for k in ["quarter", "quarterfinal"]):
        return "QF"
    if any(k in rn for k in ["round of 16", "r16", "last 16", "second round"]):
        return "R16"
    # Round of 32 (2026 format): ESPN may call it "Round of 32", "Last 32",
    # "First Knockout Round", "1st Round", "Round 1", etc.
    if any(k in rn for k in ["round of 32", "r32", "last 32", "first knockout",
                              "1st round", "round 1", "first round"]):
        return "R32"
    if "group" in rn:
        return "GS"
    # Fallback: check season type (3 = post-season / knockout in ESPN)
    stype = event.get("season", {}).get("type", 2) if isinstance(event.get("season"), dict) else 2
    return "GS" if stype != 3 else "R32"


def _parse_scoring_plays(data: dict) -> dict:
    """
    Extract {normalized_name: goal_count} from a single ESPN game summary response.

    Four-pass strategy — tries every known ESPN data location:
      Pass 1 – scoringPlays[] with participants (legacy structure)
      Pass 2 – plays[] with scoringPlay:True and participants
      Pass 3 – keyEvents[] with scoringPlay:True and participants (2026 WC structure)
      Pass 4 – leaders[] categories for goals (cross-check / last resort)
    """
    local: dict = defaultdict(int)
    ASSIST_ROLES = {"assist", "assister", "assisted by"}

    def _extract_from_play_list(plays: list) -> dict:
        """Parse participants from a list of play/event dicts. Returns {norm_name: count}."""
        found: dict = defaultdict(int)
        found_idx: set = set()
        # Sub-pass A: structured participants
        for i, play in enumerate(plays):
            play_type = play.get("type", {}).get("text", "").lower()
            if "own" in play_type:
                found_idx.add(i)
                continue
            for participant in play.get("participants", []):
                role = participant.get("type", {}).get("text", "").lower()
                if role in ASSIST_ROLES:
                    continue
                name = participant.get("athlete", {}).get("displayName", "")
                if name:
                    found[_normalize(name)] += 1
                    found_idx.add(i)
                    break
        # Sub-pass B: text-based fallback for plays with no participants
        for i, play in enumerate(plays):
            if i in found_idx:
                continue
            play_type = play.get("type", {}).get("text", "").lower()
            if "own" in play_type:
                continue
            text = (play.get("text") or play.get("shortText") or "").strip()
            if not text:
                continue
            name = re.sub(r"[\s\(]*\d+\+?\d*'?\)?$", "", text).strip()
            name = re.sub(r"(?i)^goal[\s\-:]+", "", name).strip()
            if name and 2 < len(name) < 60:
                found[_normalize(name)] += 1
        return dict(found)

    # ── Pass 1: scoringPlays[] ────────────────────────────────────────────
    sp = data.get("scoringPlays", [])
    if sp:
        local.update(_extract_from_play_list(sp))

    # ── Pass 2: plays[] with scoringPlay:True ─────────────────────────────
    if not local:
        scoring_from_plays = [p for p in data.get("plays", []) if p.get("scoringPlay")]
        if scoring_from_plays:
            local.update(_extract_from_play_list(scoring_from_plays))

    # ── Pass 3: keyEvents[] with scoringPlay:True (2026 WC structure) ─────
    if not local:
        scoring_key_events = [e for e in data.get("keyEvents", []) if e.get("scoringPlay")]
        if scoring_key_events:
            local.update(_extract_from_play_list(scoring_key_events))

    # ── Pass 4: leaders[] goals category ─────────────────────────────────
    if not local:
        for category in data.get("leaders", []):
            cat_name = (category.get("name") or category.get("displayName") or "").lower()
            if "goal" not in cat_name:
                continue
            for leader in category.get("leaders", []):
                name = leader.get("athlete", {}).get("displayName", "")
                goals = int(float(leader.get("value", 0)))
                if name and goals > 0:
                    local[_normalize(name)] += goals

    # ── Pass 5: boxscore player stats ────────────────────────────────────
    if not local:
        for team_box in data.get("boxscore", {}).get("players", []):
            for stat_group in team_box.get("statistics", []):
                for athlete_entry in stat_group.get("athletes", []):
                    stats = {s["name"]: s.get("value", 0)
                             for s in athlete_entry.get("stats", [])}
                    goals = int(float(stats.get("goals", stats.get("Goals", 0))))
                    if goals > 0:
                        name = athlete_entry.get("athlete", {}).get("displayName", "")
                        if name:
                            local[_normalize(name)] += goals

    return dict(local)


def fetch_goal_scorers(event_ids: list) -> dict:
    """
    Fetch per-player goal totals from ESPN game summaries in parallel.
    Returns {normalized_name: goal_count}.
    """
    if not event_ids:
        return {}
    print(f"  Fetching scorer data from {len(event_ids)} game summaries (parallel)…")

    def fetch_one(event_id: str) -> dict:
        try:
            resp = requests.get(ESPN_SUMMARY, params={"event": event_id}, timeout=15)
            if resp.status_code == 200:
                data = resp.json()
                result = _parse_scoring_plays(data)
                return result
        except Exception as e:
            print(f"  [DEBUG] fetch_one {event_id} error: {e}")
        return {}

    goals: dict = defaultdict(int)
    # 10 workers keeps ESPN happy while cutting wall-time from ~30s → ~5s
    with ThreadPoolExecutor(max_workers=10) as pool:
        futures = {pool.submit(fetch_one, eid): eid for eid in event_ids}
        for fut in as_completed(futures):
            for name, count in fut.result().items():
                goals[name] += count

    print(f"  ✓  {len(goals)} distinct scorers found")
    return dict(goals)


def fetch_group_standings() -> dict:
    """
    Fetch group stage standings from ESPN.
    Returns {group_name: [{"team": str, "played": N, "won": N, "drawn": N,
                           "lost": N, "gf": N, "ga": N, "gd": N, "pts": N}]}
    """
    result = {}
    try:
        resp = requests.get(ESPN_STANDINGS, timeout=15)
        if resp.status_code != 200:
            return {}
        data = resp.json()

        # ESPN returns standings under "children" (one per group) or "groups"
        groups = (
            data.get("children") or
            data.get("groups") or
            data.get("standings", {}).get("groups") or
            []
        )

        for group in groups:
            g_name = group.get("name", group.get("abbreviation", ""))
            entries = (
                group.get("standings", {}).get("entries") or
                group.get("entries") or
                []
            )
            if not entries:
                continue

            rows = []
            for entry in entries:
                team_name = (
                    entry.get("team", {}).get("displayName") or
                    entry.get("team", {}).get("name") or "?"
                )
                stats_list = entry.get("stats", [])
                stats = {s["name"]: s.get("value", 0) for s in stats_list}

                gp   = int(stats.get("gamesPlayed",    stats.get("played",    0)))
                wins = int(stats.get("wins",            stats.get("won",       0)))
                draws= int(stats.get("ties",            stats.get("drawn",     stats.get("draws", 0))))
                loss = int(stats.get("losses",          stats.get("lost",      0)))
                gf   = int(stats.get("pointsFor",       stats.get("goalsFor",  stats.get("gf", 0))))
                ga   = int(stats.get("pointsAgainst",   stats.get("goalsAgainst", stats.get("ga", 0))))
                gd   = int(stats.get("pointDifferential", stats.get("goalDifference", gf - ga)))
                pts  = int(stats.get("points", wins * 3 + draws))

                rows.append({
                    "team": team_name, "played": gp,
                    "won": wins, "drawn": draws, "lost": loss,
                    "gf": gf, "ga": ga, "gd": gd, "pts": pts,
                })
            rows.sort(key=lambda r: (-r["pts"], -r["gd"], -r["gf"]))
            result[g_name] = rows

    except Exception as e:
        print(f"  ⚠  Group standings fetch failed: {e}")
    return result


def fetch_results() -> tuple[dict, set, set, set, dict, dict]:
    """
    Fetch completed World Cup results from ESPN.

    Returns:
        team_results  — {display_name: {"wins": N, "ties": N, "losses": N}}
        advanced      — teams confirmed in knockout rounds
        champion      — tournament winner (empty until final played)
        knocked_out   — eliminated teams
        scorer_goals  — {normalized_player_name: goal_count}
        round_matches — {round: [{"home": str, "away": str, "home_score": N,
                                  "away_score": N, "winner": str, "date": str}]}
    """
    team_results     = defaultdict(lambda: {"wins": 0, "ties": 0, "losses": 0})
    group_game_count = defaultdict(int)
    advanced         = set()
    knocked_out      = set()
    champion         = set()
    event_ids        = []
    round_matches    = defaultdict(list)
    fetch_succeeded  = False   # True if ESPN returned any response (even 0 completed games)

    today   = date.today()
    current = TOURNAMENT_START

    while current <= min(today, TOURNAMENT_END):
        params = {"dates": current.strftime("%Y%m%d"), "limit": "25"}
        try:
            resp = requests.get(ESPN_SCOREBOARD, params=params, timeout=15)
            resp.raise_for_status()
            data = resp.json()
            fetch_succeeded = True   # API responded successfully

            for event in data.get("events", []):
                status = event.get("status", {}).get("type", {})
                if not status.get("completed", False):
                    continue

                comps_list = event.get("competitions", [{}])
                if not comps_list:
                    continue
                comp        = comps_list[0]
                competitors = comp.get("competitors", [])
                if len(competitors) != 2:
                    continue

                round_cat = _round_category(event)
                is_group  = round_cat == "GS"
                is_final  = round_cat == "F"

                c0, c1 = competitors[0], competitors[1]
                n0 = c0.get("team", {}).get("displayName", "")
                n1 = c1.get("team", {}).get("displayName", "")
                s0 = int(c0.get("score", 0) or 0)
                s1 = int(c1.get("score", 0) or 0)
                w0 = bool(c0.get("winner", False))
                w1 = bool(c1.get("winner", False))

                if not n0 or not n1:
                    continue

                event_ids.append(event.get("id", ""))

                if is_group:
                    group_game_count[n0] += 1
                    group_game_count[n1] += 1
                else:
                    advanced.add(n0)
                    advanced.add(n1)

                winner_name = ""
                if w0 and not w1:
                    team_results[n0]["wins"] += 1
                    if is_group:
                        team_results[n1]["losses"] += 1
                    else:
                        team_results[n1]["ties"] += 1
                        knocked_out.add(_normalize(n1))
                    winner_name = n0
                    if is_final:
                        champion.add(n0)
                        knocked_out.add(_normalize(n1))
                elif w1 and not w0:
                    team_results[n1]["wins"] += 1
                    if is_group:
                        team_results[n0]["losses"] += 1
                    else:
                        team_results[n0]["ties"] += 1
                        knocked_out.add(_normalize(n0))
                    winner_name = n1
                    if is_final:
                        champion.add(n1)
                else:
                    team_results[n0]["ties"] += 1
                    team_results[n1]["ties"] += 1

                round_matches[round_cat].append({
                    "home": n0, "away": n1,
                    "home_score": s0, "away_score": s1,
                    "winner": winner_name,
                    "date": current.strftime("%b %d"),
                })

        except requests.exceptions.ConnectionError:
            print(f"  ⚠  Network error for {current} — skipping")
        except requests.exceptions.HTTPError as e:
            print(f"  ⚠  HTTP {e.response.status_code} for {current} — skipping")
        except Exception as e:
            print(f"  ⚠  Error for {current}: {e}")

        current += timedelta(days=1)

    # Apply manual overrides BEFORE computing group-stage knockouts so that
    # teams in MANUAL_ADVANCED are never incorrectly flagged as eliminated.
    advanced |= MANUAL_ADVANCED

    for team, data in MANUAL_TEAM_OVERRIDE.items():
        if "wins"   in data: team_results[team]["wins"]   = data["wins"]
        if "ties"   in data: team_results[team]["ties"]   = data["ties"]
        if "losses" in data: team_results[team]["losses"] = data["losses"]
        if data.get("advanced"):  advanced.add(team)
        if data.get("champion"):  champion.add(team)

    # In the 48-team format 8 of 12 third-place teams advance, so we never
    # auto-eliminate based on games played. Only add confirmed 4th-place teams.
    for name in MANUAL_ELIMINATED:
        knocked_out.add(_normalize(name))

    # Normalize advanced and champion so match_team can use pick_norm lookups
    # consistently (same as knocked_out). This handles ESPN name variants like
    # "Congo DR" vs "DR Congo", "Korea Republic" vs "South Korea", etc.
    advanced    = {_normalize(a) for a in advanced}
    champion    = {_normalize(c) for c in champion}

    total_games = len(event_ids)
    print(f"  ESPN: {total_games} games | {len(team_results)} teams | "
          f"{len(advanced)} advanced | {len(knocked_out)} eliminated")

    scorer_goals = fetch_goal_scorers([eid for eid in event_ids if eid])
    for raw_name, goals in GOAL_SCORER_OVERRIDE.items():
        scorer_goals[_normalize(raw_name)] = goals

    return dict(team_results), advanced, champion, knocked_out, scorer_goals, dict(round_matches), fetch_succeeded


# ============================================================
# TEAM / SCORER MATCHING
# ============================================================

def match_team(pick: str, team_results: dict, advanced: set,
               champion: set, knocked_out: set) -> dict:
    _empty = {
        "wins": 0, "ties": 0, "losses": 0,
        "advanced": False, "champion": False, "knocked_out": False,
        "matched_name": _pretty(pick), "has_played": False, "pts_so_far": 0,
    }
    if not pick or (isinstance(pick, float) and pd.isna(pick)):
        return _empty

    pick_norm = _normalize(str(pick))
    # knocked_out is a set of NORMALIZED strings — never include it in all_names
    # (it would create phantom entries that match pick_norm but have no team_results data).
    # Instead check knocked_out separately via pick_norm comparison.
    all_names = set(team_results.keys()) | advanced | champion

    # advanced, champion, knocked_out are all normalized sets — use pick_norm directly.
    for name in all_names:
        if _normalize(name) == pick_norm:
            r   = team_results.get(name, {"wins": 0, "ties": 0, "losses": 0})
            pts = r["wins"] * WIN_PTS + r["ties"] * TIE_PTS
            return {
                "wins": r["wins"], "ties": r["ties"], "losses": r["losses"],
                "advanced":    pick_norm in advanced,
                "champion":    pick_norm in champion,
                "knocked_out": pick_norm in knocked_out,
                "matched_name": name,
                "has_played": r["wins"] + r["ties"] + r["losses"] > 0 or pick_norm in advanced,
                "pts_so_far": pts,
            }

    pick_words  = set(pick_norm.split())
    best_name, best_score = None, 0.0
    for name in all_names:
        nw = set(_normalize(name).split())
        if pick_words and pick_words.issubset(nw):
            score = len(pick_words & nw) / len(pick_words | nw)
            if score > best_score:
                best_score, best_name = score, name

    if best_name:
        r   = team_results.get(best_name, {"wins": 0, "ties": 0, "losses": 0})
        pts = r["wins"] * WIN_PTS + r["ties"] * TIE_PTS
        return {
            "wins": r["wins"], "ties": r["ties"], "losses": r["losses"],
            "advanced":    pick_norm in advanced,
            "champion":    pick_norm in champion,
            "knocked_out": pick_norm in knocked_out,
            "matched_name": best_name,
            "has_played": r["wins"] + r["ties"] + r["losses"] > 0 or pick_norm in advanced,
            "pts_so_far": pts,
        }
    return _empty


def match_scorer(pick: str, scorer_goals: dict) -> int:
    if not pick or (isinstance(pick, float) and pd.isna(pick)):
        return 0
    pick_norm  = _normalize(str(pick))
    if pick_norm in scorer_goals:
        return scorer_goals[pick_norm]
    pick_words = set(pick_norm.split())
    for espn_norm, goals in scorer_goals.items():
        espn_words = set(espn_norm.split())
        if pick_words and pick_words.issubset(espn_words):
            return goals
        if espn_words and espn_words.issubset(pick_words):
            return goals
    return 0


# ============================================================
# READ ENTRIES
# ============================================================

def _detect_columns(df: pd.DataFrame) -> dict:
    result = {'A': [], 'B': [], 'C': [], 'D': [],
              'winner': None, 'scorer1': None, 'scorer2': None, 'name': None}
    tier_patterns = {
        'A': [r'^tier.?a', r'^favorites?', r'^favourites?'],
        'B': [r'^tier.?b', r'^dark.?horse'],
        'C': [r'^tier.?c', r'^longshots?'],
        'D': [r'^tier.?d', r'^glad.?you.?came', r"^we.?re.?glad", r'^we.*glad'],
    }
    scorer_cols = []
    for col in df.columns:
        cl = str(col).lower().strip()
        if re.search(r'^(your\s+)?name$|^full[\s_]?name$|^first.*last', cl) and result['name'] is None:
            result['name'] = col
            continue
        if re.search(r'winner|champion|tournament.?winner|tournament.?pick', cl) and 'scorer' not in cl:
            if result['winner'] is None:
                result['winner'] = col
            continue
        if re.search(r'goal.?scorer|goalscorer|scorer', cl):
            scorer_cols.append(col)
            continue
        for tier, patterns in tier_patterns.items():
            if any(re.search(p, cl) for p in patterns):
                result[tier].append(col)
                break
    if scorer_cols:
        result['scorer1'] = scorer_cols[0]
        if len(scorer_cols) > 1:
            result['scorer2'] = scorer_cols[1]
    return result


def _parse_picks(row, cols: list, expected: int) -> list:
    picks = []
    if not cols:
        return picks
    if len(cols) == 1:
        raw = str(row.get(cols[0], "") or "").strip()
        if not raw or raw.lower() == "nan":
            return picks
        for sep in [";", ","]:
            parts = [p.strip() for p in raw.split(sep) if p.strip()]
            if len(parts) > 1:
                picks = parts
                break
        if not picks:
            picks = [raw]
    else:
        for col in cols:
            val = str(row.get(col, "") or "").strip()
            if val and val.lower() != "nan":
                picks.append(val)
    return [p for p in picks if p][:expected]


def read_entries(excel_file: str) -> list:
    if excel_file.lower().endswith(".csv"):
        df = pd.read_csv(excel_file, header=0)
    else:
        loaded = False
        for sheet in ["Form Responses 1", "Raw Data", 0]:
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet, header=0)
                loaded = True
                break
            except Exception:
                continue
        if not loaded:
            raise ValueError(f"Could not read {excel_file}.")

    email_candidates = [c for c in df.columns if "email" in str(c).lower() or c == "Username"]
    if not email_candidates:
        raise ValueError(f"No email column found. Columns: {list(df.columns)}")
    email_col = max(email_candidates, key=lambda c: df[c].notna().sum())

    if "Timestamp" in df.columns:
        df = df.sort_values("Timestamp")
    df = df[df[email_col].notna()].copy()

    col_map = _detect_columns(df)
    print(f"  Column map:")
    for tier in 'ABCD':
        print(f"    Tier {tier}: {col_map[tier]}")
    print(f"    Winner:  {col_map['winner']}")
    print(f"    Scorer1: {col_map['scorer1']}  |  Scorer2: {col_map['scorer2']}")

    def _extract(row) -> dict:
        entry = {}
        for tier in 'ABCD':
            entry[tier] = _parse_picks(row, col_map[tier], TIER_PICKS[tier])
        winner_val = ""
        if col_map['winner']:
            v = str(row.get(col_map['winner'], "") or "").strip()
            if v.lower() != "nan":
                winner_val = v
        entry['winner'] = winner_val
        scorers = []
        for key in ['scorer1', 'scorer2']:
            if col_map[key]:
                v = str(row.get(col_map[key], "") or "").strip()
                if v and v.lower() != "nan":
                    scorers.append(v)
        entry['scorers'] = scorers[:2]
        return entry

    groups: dict = OrderedDict()
    for _, row in df.iterrows():
        email = str(row[email_col]).strip().lower()
        if not email or email == "nan":
            continue
        groups.setdefault(email, []).append(row)

    entries = []
    for email, rows in groups.items():
        multi = len(rows) > 1
        for idx, row in enumerate(rows, start=1):
            picks = _extract(row)
            display = ""
            if col_map['name']:
                v = str(row.get(col_map['name'], "") or "").strip()
                if v and v.lower() != "nan":
                    display = v
            if not display:
                username = email.split("@")[0]
                words    = re.split(r"[._\-]+", username)
                words    = [re.sub(r"\d+$", "", w).capitalize() for w in words
                            if re.sub(r"\d+$", "", w)]
                display  = " ".join(words) if words else username.title()
            if multi:
                display = f"{display} ({idx})"
            entries.append({"email": email, "name": display, "picks": picks})
    return entries


# ============================================================
# SCORING
# ============================================================

def calculate_scores(entries: list, team_results: dict, advanced: set,
                     champion: set, knocked_out: set, scorer_goals: dict) -> list:
    standings = []
    for entry in entries:
        picks  = entry["picks"]
        total  = 0
        detail = []
        alive  = 0
        for tier in 'ABCD':
            for team_raw in picks.get(tier, []):
                tm       = match_team(team_raw, team_results, advanced, champion, knocked_out)
                adv_bonus = TIER_ADVANCE_BONUS[tier] if tm["advanced"] else 0
                pts       = tm["pts_so_far"] + adv_bonus
                total    += pts
                if not tm["knocked_out"]:
                    alive += 1
                detail.append({
                    "tier": tier,
                    "team": _pretty(tm["matched_name"] or team_raw),
                    "wins": tm["wins"], "ties": tm["ties"], "losses": tm["losses"],
                    "advanced": tm["advanced"], "champion": tm["champion"],
                    "knocked_out": tm["knocked_out"],
                    "pts": pts, "adv_bonus": adv_bonus, "has_played": tm["has_played"],
                })

        winner_pick = picks.get("winner", "")
        winner_pts  = 0
        got_winner  = False
        if winner_pick and champion:
            wn = _normalize(winner_pick)
            for champ in champion:
                if _normalize(champ) == wn or wn in _normalize(champ):
                    got_winner = True
                    winner_pts = WINNER_BONUS
                    break
        total += winner_pts

        scorer_detail = []
        scorer_pts    = 0
        for scorer_raw in picks.get("scorers", []):
            goals  = match_scorer(scorer_raw, scorer_goals)
            pts_sc = goals * GOAL_PTS
            scorer_pts += pts_sc
            scorer_detail.append({"name": _pretty(scorer_raw), "goals": goals, "pts": pts_sc})
        total += scorer_pts

        # For teams confirmed through/out of the group stage, credit at least 3 games
        # even if ESPN data is incomplete (name mismatch, API gap, etc.)
        games_played = sum(
            max(d["wins"] + d["ties"] + d["losses"], 3)
            if (d["advanced"] or d["knocked_out"])
            else (d["wins"] + d["ties"] + d["losses"])
            for d in detail
        )
        standings.append({
            "email": entry["email"], "name": entry["name"],
            "total_points": total, "teams_alive": alive,
            "games_played": games_played,
            "team_detail": detail,
            "winner_pick": _pretty(winner_pick), "winner_pts": winner_pts, "got_winner": got_winner,
            "scorer_detail": scorer_detail, "scorer_pts": scorer_pts,
        })

    standings.sort(key=lambda x: (-x["total_points"], x["name"]))
    # Assign tied ranks: same score → same rank number, skip intervening numbers
    rank = 1
    for i, s in enumerate(standings):
        if i > 0 and s["total_points"] == standings[i-1]["total_points"]:
            s["rank"] = standings[i-1]["rank"]
        else:
            s["rank"] = rank
        rank = i + 2
    return standings


def pick_popularity(entries: list) -> dict:
    tier_pop   = {t: defaultdict(int) for t in 'ABCD'}
    winner_pop = defaultdict(int)
    scorer_pop = defaultdict(int)
    for e in entries:
        picks = e["picks"]
        for tier in 'ABCD':
            for team in picks.get(tier, []):
                tier_pop[tier][_pretty(team)] += 1
        if picks.get("winner"):
            winner_pop[_pretty(picks["winner"])] += 1
        for sc in picks.get("scorers", []):
            if sc:
                scorer_pop[_pretty(sc)] += 1
    return {
        "tiers":   {t: dict(sorted(tier_pop[t].items(), key=lambda x: -x[1])) for t in 'ABCD'},
        "winner":  dict(sorted(winner_pop.items(),  key=lambda x: -x[1])),
        "scorers": dict(sorted(scorer_pop.items(),  key=lambda x: -x[1])),
    }


# ============================================================
# HTML TAB BUILDERS
# ============================================================

TIER_COLORS = {'A': '#7C3AED', 'B': '#1A56DB', 'C': '#D97706', 'D': '#64748B'}
TIER_BG     = {'A': '#F5F3FF', 'B': '#EFF6FF', 'C': '#FFFBEB', 'D': '#F8FAFC'}

ROUND_LABELS = {
    "GS":  "Group Stage",
    "R32": "Round of 32",
    "R16": "Round of 16",
    "QF":  "Quarter-finals",
    "SF":  "Semi-finals",
    "3P":  "Third Place",
    "F":   "Final",
}
ROUND_ORDER = ["GS", "R32", "R16", "QF", "SF", "3P", "F"]


def _chip_html(d: dict) -> str:
    team  = _h(d["team"])
    pts   = f"+{d['pts']}" if d["pts"] else "0"
    tier  = d["tier"]
    color = TIER_COLORS.get(tier, '#64748B')
    played = d["wins"] + d["ties"] + d["losses"] > 0
    if d["knocked_out"]:
        css, icon = "chip-out", "✗"
    elif d["champion"]:
        css, icon = "chip-winner", "🏆"
    elif d["advanced"]:
        css, icon = "chip-adv", "✓"
    elif played and d["pts"] == 0:
        css, icon = "chip-zero", "–"
    else:
        css, icon = "chip-alive", "●"
    label = f'<span class="chip-tier" style="background:{color}20;color:{color}">{tier}</span>'
    return f'<span class="pick-chip {css}">{icon} {label} {team} <em>({pts})</em></span>\n'


def _build_group_tab(group_standings: dict) -> str:
    if not group_standings:
        return """
    <div class="tab-pane fade" id="tab-groups" role="tabpanel">
      <div class="tab-hero-lg" style="--tab-bg:url('images/groups_bg.jpg')">
        <div class="tab-hero-inner">
          <div class="tab-hero-title">Group Stage</div>
          <div class="tab-hero-sub">Standings will appear once the tournament kicks off (June 11)</div>
        </div>
      </div>
      <div class="container-inner">
        <div class="empty-state">
          ⚽ <strong>Tournament hasn't started yet.</strong><br>
          Group standings will appear here automatically once games are played.
        </div>
      </div>
    </div>"""

    groups_html = ""
    sorted_groups = sorted(group_standings.items())
    for i, (g_name, rows) in enumerate(sorted_groups):
        teams_html = ""
        for pos, row in enumerate(rows, 1):
            adv_cls = "gs-adv" if pos <= 2 else ""  # Top 2 advance (simplified)
            teams_html += f"""
            <tr class="{adv_cls}">
              <td class="gs-pos">{pos}</td>
              <td class="gs-team">{_h(row['team'])}</td>
              <td class="text-center">{row['played']}</td>
              <td class="text-center fw-bold">{row['won']}</td>
              <td class="text-center">{row['drawn']}</td>
              <td class="text-center">{row['lost']}</td>
              <td class="text-center">{row['gf']}</td>
              <td class="text-center">{row['ga']}</td>
              <td class="text-center">{'+' if row['gd'] > 0 else ''}{row['gd']}</td>
              <td class="text-center gs-pts">{row['pts']}</td>
            </tr>"""

        groups_html += f"""
        <div class="gs-card">
          <div class="gs-card-head">{_h(g_name)}</div>
          <table class="gs-table">
            <thead>
              <tr>
                <th>#</th><th>Team</th>
                <th title="Played">P</th><th title="Won">W</th>
                <th title="Drawn">D</th><th title="Lost">L</th>
                <th title="Goals For">GF</th><th title="Goals Against">GA</th>
                <th title="Goal Difference">GD</th><th title="Points">Pts</th>
              </tr>
            </thead>
            <tbody>{teams_html}</tbody>
          </table>
        </div>"""

    return f"""
    <div class="tab-pane fade" id="tab-groups" role="tabpanel">
      <div class="tab-hero-lg" style="--tab-bg:url('images/groups_bg.jpg')">
        <div class="tab-hero-inner">
          <div class="tab-hero-title">Group Stage</div>
          <div class="tab-hero-sub">{len(group_standings)} groups · Top 2 advance + 8 best 3rd-place teams</div>
        </div>
      </div>
      <div class="container-inner">
        <div class="gs-grid">{groups_html}</div>
        <p class="gs-legend">
          <span class="gs-adv-dot"></span> Advancing position
        </p>
      </div>
    </div>"""


def _build_bracket_tab(round_matches: dict) -> str:
    has_data = any(v for k, v in round_matches.items() if k != "GS")

    if not has_data:
        # Placeholder bracket — 2026 World Cup structure
        placeholder_rounds = [
            ("Round of 32",    16),
            ("Round of 16",     8),
            ("Quarter-finals",  4),
            ("Semi-finals",     2),
            ("Final",           1),
        ]
        rounds_html = ""
        for label, n_matches in placeholder_rounds:
            cards = ""
            for _ in range(n_matches):
                cards += """
            <div class="bk-card bk-placeholder">
              <div class="bk-match-row">
                <span class="bk-team bk-tbd">TBD</span>
                <span class="bk-score bk-tbd-score">–</span>
              </div>
              <div class="bk-divider"></div>
              <div class="bk-match-row">
                <span class="bk-team bk-tbd">TBD</span>
                <span class="bk-score bk-tbd-score">–</span>
              </div>
              <div class="bk-date bk-tbd">Pending</div>
            </div>"""
            rounds_html += f"""
        <div class="bk-round">
          <div class="bk-round-label">{label}
            <span class="bk-round-count">{n_matches} match{'es' if n_matches != 1 else ''}</span>
          </div>
          <div class="bk-cards">{cards}</div>
        </div>"""

        return f"""
    <div class="tab-pane fade" id="tab-bracket" role="tabpanel">
      <div class="tab-hero-lg" style="--tab-bg:url('images/bracket_bg.jpg')">
        <div class="tab-hero-inner">
          <div class="tab-hero-title">Knockout Bracket</div>
          <div class="tab-hero-sub">Round of 32 · Round of 16 · Quarter-finals · Semi-finals · Final</div>
        </div>
      </div>
      <div class="container-inner">
        <p class="text-muted small mb-4">Fills in automatically as matches are completed · Round of 32 runs June 28–July 3</p>
        {rounds_html}
      </div>
    </div>"""

    rounds_html = ""
    for round_key in ROUND_ORDER:
        if round_key == "GS":
            continue
        matches = round_matches.get(round_key, [])
        if not matches:
            continue

        label = ROUND_LABELS[round_key]
        cards  = ""
        for m in matches:
            hw = "bk-winner" if m["winner"] == m["home"] else ""
            aw = "bk-winner" if m["winner"] == m["away"] else ""
            hl = "bk-loser"  if m["winner"] and m["winner"] != m["home"] else ""
            al = "bk-loser"  if m["winner"] and m["winner"] != m["away"] else ""

            home_score = m["home_score"] if m["winner"] else "–"
            away_score = m["away_score"] if m["winner"] else "–"

            trophy = " 🏆" if round_key == "F" and m["winner"] == m["home"] else ""
            trophy_a = " 🏆" if round_key == "F" and m["winner"] == m["away"] else ""

            cards += f"""
            <div class="bk-card">
              <div class="bk-match-row {hw} {hl}">
                <span class="bk-team">{_h(m['home'])}{trophy}</span>
                <span class="bk-score">{home_score}</span>
              </div>
              <div class="bk-divider"></div>
              <div class="bk-match-row {aw} {al}">
                <span class="bk-team">{_h(m['away'])}{trophy_a}</span>
                <span class="bk-score">{away_score}</span>
              </div>
              <div class="bk-date">{m['date']}</div>
            </div>"""

        rounds_html += f"""
        <div class="bk-round">
          <div class="bk-round-label">{label}
            <span class="bk-round-count">{len(matches)} match{'es' if len(matches) != 1 else ''}</span>
          </div>
          <div class="bk-cards">{cards}</div>
        </div>"""

    return f"""
    <div class="tab-pane fade" id="tab-bracket" role="tabpanel">
      <div class="tab-hero-lg" style="--tab-bg:url('images/bracket_bg.jpg')">
        <div class="tab-hero-inner">
          <div class="tab-hero-title">Knockout Bracket</div>
          <div class="tab-hero-sub">Round of 32 · Round of 16 · Quarter-finals · Semi-finals · Final</div>
        </div>
      </div>
      <div class="container-inner">{rounds_html}</div>
    </div>"""


def _build_scorers_tab(standings: list, scorer_goals: dict, entries: list) -> str:
    """Show only the goal scorers that pool participants actually picked."""
    # Collect all unique scorer picks + their pick counts
    scorer_picks: dict = defaultdict(lambda: {"picks": 0, "pickers": [], "goals": 0, "pts": 0})

    for s in standings:
        for sc in s["scorer_detail"]:
            key = sc["name"].lower()
            scorer_picks[sc["name"]]["picks"] += 1
            scorer_picks[sc["name"]]["goals"] = sc["goals"]
            scorer_picks[sc["name"]]["pts"]   = sc["pts"]
            scorer_picks[sc["name"]]["pickers"].append(s["name"])

    if not scorer_picks:
        return """
    <div class="tab-pane fade" id="tab-scorers" role="tabpanel">
      <div class="tab-hero-lg" style="--tab-bg:url('images/scorers_bg.jpg')">
        <div class="tab-hero-inner">
          <div class="tab-hero-title">⚽ Goal Scorers</div>
          <div class="tab-hero-sub">+150 pts per goal · automatically tracked from ESPN</div>
        </div>
      </div>
      <div class="container-inner">
        <div class="empty-state">No scorer picks found.</div>
      </div>
    </div>"""

    # Sort: goals desc, then picks desc
    sorted_scorers = sorted(scorer_picks.items(),
                            key=lambda x: (-x[1]["goals"], -x[1]["picks"]))
    max_goals = max((v["goals"] for _, v in sorted_scorers), default=1) or 1

    rows_html = ""
    for rank, (name, data) in enumerate(sorted_scorers, 1):
        goals = data["goals"]
        pts   = data["pts"]
        picks = data["picks"]
        pct   = round(picks / len(entries) * 100) if entries else 0
        bar_w = round(goals / max_goals * 100) if max_goals and goals else 0

        if rank == 1 and goals > 0:
            row_cls, rank_cls = "sc-row-gold", "rank-1"
        elif rank == 2 and goals > 0:
            row_cls, rank_cls = "sc-row-silver", "rank-2"
        elif rank == 3 and goals > 0:
            row_cls, rank_cls = "sc-row-bronze", "rank-3"
        else:
            row_cls, rank_cls = "", "rank-n"

        goal_bar = (f'<div class="goal-bar-wrap"><div class="goal-bar" '
                    f'style="width:{bar_w}%"></div></div>') if goals else ""
        pts_html = (f'<span class="sc-pts">+{pts} pts</span>') if pts else '<span class="sc-pts sc-zero">0 pts</span>'
        goals_html = (f'<span class="sc-goals-badge">{goals} ⚽</span>'
                      if goals else '<span class="sc-goals-badge sc-no-goal">0 ⚽</span>')

        expand_id = f"sc-exp-{rank}"
        if picks > 4:
            all_pickers_html = "".join(
                f'<span class="sc-picker-chip">{_h(p)}</span>' for p in data["pickers"]
            )
            picker_section = f"""
            <div class="sc-pickers">
              Picked by: {_h(", ".join(data["pickers"][:4]))}
              <button class="sc-expand-btn" onclick="toggleScPickers('{expand_id}',this)">
                +{picks-4} more ▾
              </button>
            </div>
            <div class="sc-pickers-full" id="{expand_id}" style="display:none">
              {all_pickers_html}
            </div>"""
        else:
            picker_section = f'<div class="sc-pickers">Picked by: {_h(", ".join(data["pickers"]))}</div>'

        rows_html += f"""
        <div class="sc-row {row_cls}">
          <div class="sc-rank"><span class="rank-badge {rank_cls}" style="width:28px;height:28px;font-size:.75rem">{rank}</span></div>
          <div class="sc-info">
            <div class="sc-name">{_h(name)}</div>
            {picker_section}
            {goal_bar}
          </div>
          <div class="sc-stats">
            {goals_html}
            {pts_html}
            <span class="sc-pickcount">{picks} pick{'s' if picks != 1 else ''} · {pct}%</span>
          </div>
        </div>"""

    # Summary stat: total goals scored by pool picks
    total_goals = sum(v["goals"] for _, v in sorted_scorers)
    total_pts   = sum(v["pts"]   for _, v in sorted_scorers)

    return f"""
    <div class="tab-pane fade" id="tab-scorers" role="tabpanel">
      <div class="tab-hero-lg" style="--tab-bg:url('images/scorers_bg.jpg')">
        <div class="tab-hero-inner">
          <div class="tab-hero-title">⚽ Goal Scorers</div>
          <div class="tab-hero-sub">+150 pts per goal · automatically tracked from ESPN</div>
        </div>
      </div>
      <div class="container-inner">
        <div class="sc-list">{rows_html}</div>
      </div>
    </div>"""


def _build_raw_data_tab(standings: list, team_results: dict,
                        advanced: set, champion: set, knocked_out: set,
                        scorer_goals: dict, rank_counts: dict = None) -> str:
    if rank_counts is None:
        rank_counts = defaultdict(int)
    tier_headers = ""
    for tier in 'ABCD':
        for i in range(1, TIER_PICKS[tier] + 1):
            color = TIER_COLORS[tier]
            tier_headers += (
                f'<th class="text-center rp-tier-hdr" '
                f'style="border-bottom:2px solid {color}" '
                f'title="{TIER_LABELS[tier]}">'
                f'<span style="color:{color};font-size:.65rem">Tier {tier}</span><br>'
                f'<span style="font-size:.7rem">#{i}</span></th>'
            )
    tier_headers += (
        '<th class="text-center rp-tier-hdr" style="border-bottom:2px solid #16A34A">'
        '<span style="color:#16A34A;font-size:.65rem">Winner</span><br>'
        '<span style="font-size:.7rem">Pick</span></th>'
        '<th class="text-center rp-tier-hdr" style="border-bottom:2px solid #EA580C">'
        '<span style="color:#EA580C;font-size:.65rem">⚽ Scorer</span><br>'
        '<span style="font-size:.7rem">#1</span></th>'
        '<th class="text-center rp-tier-hdr" style="border-bottom:2px solid #EA580C">'
        '<span style="color:#EA580C;font-size:.65rem">⚽ Scorer</span><br>'
        '<span style="font-size:.7rem">#2</span></th>'
    )

    rows_html = ""
    for s in standings:
        picks = []
        for tier in 'ABCD':
            slot_picks = [d["team"] for d in s["team_detail"] if d["tier"] == tier]
            for i in range(TIER_PICKS[tier]):
                picks.append(("team", tier, slot_picks[i] if i < len(slot_picks) else ""))

        picks.append(("winner", None, s["winner_pick"]))
        for sc in s["scorer_detail"]:
            picks.append(("scorer", None, sc["name"], sc["goals"]))
        while sum(1 for p in picks if p[0] == "scorer") < 2:
            picks.append(("scorer", None, ""))

        cells = ""
        for pick_info in picks:
            kind = pick_info[0]
            if kind == "team":
                _, tier, team = pick_info
                if not team:
                    cells += '<td class="rp-empty">—</td>'
                    continue
                tm = match_team(team, team_results, advanced, champion, knocked_out)
                if tm["champion"]:      css = "rp-champ"
                elif tm["knocked_out"]: css = "rp-out"
                elif tm["has_played"]:  css = "rp-alive"
                else:                   css = "rp-notplayed"
                pts_label = f'<span class="rp-pts">+{tm["pts_so_far"]}</span>' if tm["pts_so_far"] else ""
                cells += f'<td class="{css}">{_h(team)}{pts_label}</td>'
            elif kind == "winner":
                _, _, wpick = pick_info
                if not wpick:
                    cells += '<td class="rp-empty">—</td>'
                elif s["got_winner"]:
                    cells += f'<td class="rp-champ">🏆 {_h(wpick)}<span class="rp-pts">+{WINNER_BONUS}</span></td>'
                elif champion:
                    cells += f'<td class="rp-out">{_h(wpick)}</td>'
                else:
                    cells += f'<td class="rp-alive">{_h(wpick)}</td>'
            elif kind == "scorer":
                sname = pick_info[2]
                goals = pick_info[3] if len(pick_info) > 3 else 0
                if not sname:
                    cells += '<td class="rp-empty">—</td>'
                elif goals > 0:
                    cells += f'<td class="rp-alive">{_h(sname)}<span class="rp-pts">+{goals * GOAL_PTS}</span></td>'
                else:
                    cells += f'<td class="rp-notplayed">{_h(sname)}</td>'

        rank = s["rank"]
        rp_tied = rank_counts[rank] > 1
        rp_rank_display = f"T-{rank}" if rp_tied else str(rank)
        if   rank == 1: rank_css = "rank-1"
        elif rank == 2: rank_css = "rank-2"
        elif rank == 3: rank_css = "rank-3"
        else:           rank_css = "rank-n"

        rows_html += f"""
        <tr class="rp-row" data-name="{_h(s['name'].lower())}">
          <td class="rp-rank-cell"><span class="rank-badge {rank_css}" style="width:22px;height:22px;font-size:.58rem">{rp_rank_display}</span></td>
          <td class="rp-name-cell"><span class="rp-pname">{_h(s['name'])}</span></td>
          <td class="rp-pts-cell">{s['total_points']:,}</td>
          {cells}
        </tr>"""

    return f"""
    <div class="tab-pane fade" id="tab-raw" role="tabpanel">
      <div class="tab-hero-lg" style="--tab-bg:url('images/allpicks_bg.jpg')">
        <div class="tab-hero-inner">
          <div class="tab-hero-title">📋 All Picks</div>
          <div class="tab-hero-sub">Every participant's 12 teams + winner + goal scorers</div>
        </div>
      </div>
      <div class="container-inner">
        <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
          <p class="text-muted small mb-0">
            <span class="rp-legend rp-alive-leg">● Active</span>
            <span class="rp-legend rp-out-leg">✗ Out</span>
            <span class="rp-legend rp-champ-leg">🏆 Champion</span>
          </p>
          <input id="rp-search" type="text" class="lb-search-input"
                 placeholder="🔍  Filter by name…" oninput="filterRaw()">
        </div>
        <div class="rp-wrap table-responsive">
          <table class="table mb-0 table-sm" id="rp-table">
            <thead>
              <tr>
                <th class="rp-rank-hdr">#</th>
                <th class="rp-name-hdr">Name</th>
                <th class="rp-pts-hdr">Pts</th>
                {tier_headers}
              </tr>
            </thead>
            <tbody id="rp-body">{rows_html}</tbody>
          </table>
        </div>
      </div>
    </div>"""


# ============================================================
# MAIN HTML GENERATOR
# ============================================================

def generate_html(standings: list, team_results: dict, advanced: set,
                  champion: set, knocked_out: set, entries: list,
                  pop: dict, scorer_goals: dict, group_standings: dict,
                  round_matches: dict, fetch_attempted: bool = False,
                  fetch_succeeded: bool = False) -> str:

    last_updated = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    total_p      = len(standings)
    total_games  = sum(v["wins"] + v["losses"] + v["ties"] for v in team_results.values()) // 2
    leader       = standings[0] if standings else {"name": "TBD", "total_points": 0}
    all_teams_flat = TIER_DEFS['A'] + TIER_DEFS['B'] + TIER_DEFS['C'] + TIER_DEFS['D']
    alive_count  = sum(1 for t in all_teams_flat
                       if not any(_normalize(t) == _normalize(k) for k in knocked_out))
    today = date.today()
    active = TOURNAMENT_START <= today <= TOURNAMENT_END

    # Only show stale banner if the ESPN fetch actually failed (not just 0 completed games yet)
    stale = (
        '<div class="stale-banner">⚠️  Score data could not be retrieved from ESPN. '
        'Standings may be out of date.</div>'
        if (fetch_attempted and active and not fetch_succeeded) else ""
    )

    # ── Leaderboard rows ───────────────────────────────────────────────────
    # Pre-compute which ranks are tied
    rank_counts = defaultdict(int)
    for s in standings:
        rank_counts[s["rank"]] += 1

    rows_html = ""
    for idx, s in enumerate(standings):
        rank = s["rank"]
        tied = rank_counts[rank] > 1
        rank_display = f"T-{rank}" if tied else str(rank)
        if   rank == 1: badge_cls, row_cls = "rank-1", "row-gold"
        elif rank == 2: badge_cls, row_cls = "rank-2", "row-silver"
        elif rank == 3: badge_cls, row_cls = "rank-3", "row-bronze"
        else:           badge_cls, row_cls = "rank-n", ""

        fire       = ' <span class="fire-tag">🔥</span>' if rank == 1 else ""
        alive_list = [d for d in s["team_detail"] if not d["knocked_out"]]
        alive_str  = ", ".join(d["team"] for d in alive_list[:3])
        if len(alive_list) > 3:
            alive_str += f" +{len(alive_list)-3} more"

        chips = "".join(_chip_html(d) for d in s["team_detail"])
        if s["winner_pick"]:
            wicon = "🏆" if s["got_winner"] else "⭐"
            wcss  = "chip-winner" if s["got_winner"] else "chip-alive"
            wpts  = f"+{s['winner_pts']}" if s["winner_pts"] else "0"
            chips += (f'<span class="pick-chip {wcss} chip-winner-pick">'
                      f'{wicon} <span class="chip-tier" style="background:#F0FDF4;color:#16A34A">WIN</span>'
                      f' {_h(s["winner_pick"])} <em>({wpts})</em></span>\n')
        for sc in s["scorer_detail"]:
            spts   = f"+{sc['pts']}" if sc["pts"] else "0 pts"
            sgoals = f"{sc['goals']} goal{'s' if sc['goals'] != 1 else ''}" if sc["goals"] else "0 goals"
            chips += (f'<span class="pick-chip chip-scorer">'
                      f'⚽ <span class="chip-tier" style="background:#FFF7ED;color:#EA580C">GOAL</span>'
                      f' {_h(sc["name"])} <em>({sgoals} · {spts})</em></span>\n')

        row_id     = f"pr-{idx}"
        alive_html = (f'<span class="badge-alive">{s["teams_alive"]}</span>'
                      if s["teams_alive"] > 0
                      else '<span class="badge-dead">0</span>')

        rows_html += f"""
        <tr class="lb-row {row_cls}" onclick="toggleRow('{row_id}')">
          <td><span class="rank-badge {badge_cls}">{rank_display}</span></td>
          <td>
            <span class="p-name">{_h(s['name'])}{fire}</span>
            <span class="p-email">{_h(s['email'])}</span>
          </td>
          <td class="text-end"><span class="pts-val">{s['total_points']:,}</span></td>
          <td class="text-center">{s['games_played']}</td>
          <td class="text-center">{alive_html}</td>
          <td class="d-none d-lg-table-cell still-in-col">{alive_str or '—'}</td>
        </tr>
        <tr class="picks-row" id="{row_id}">
          <td colspan="6"><div class="picks-inner">{chips}</div></td>
        </tr>"""

    # ── Team tracker ───────────────────────────────────────────────────────
    all_teams_data = []
    for tier, teams in TIER_DEFS.items():
        for team in teams:
            r = team_results.get(team, {})
            if not r:
                for espn_name, data in team_results.items():
                    if _normalize(team) == _normalize(espn_name):
                        r = data
                        break
            wins   = r.get("wins", 0)
            ties   = r.get("ties", 0)
            losses = r.get("losses", 0)
            game_pts  = wins * WIN_PTS + ties * TIE_PTS
            adv   = any(_normalize(team) == _normalize(a) for a in advanced)
            champ = any(_normalize(team) == _normalize(c) for c in champion)
            kout  = any(_normalize(team) == _normalize(k) for k in knocked_out)
            played = wins + ties + losses > 0
            if champ:      status = "champion"
            elif kout:     status = "out"
            elif adv:      status = "advanced"
            elif played:   status = "active"
            else:          status = "not_played"
            pick_count = sum(
                1 for e in entries
                for t in e["picks"].get(tier, [])
                if _normalize(t) == _normalize(team) or _normalize(team) in _normalize(t)
            )
            pct = round(pick_count / len(entries) * 100, 1) if entries else 0.0
            adv_bonus = TIER_ADVANCE_BONUS[tier] if adv else 0
            all_teams_data.append({
                "tier": tier, "team": team, "status": status,
                "wins": wins, "ties": ties, "losses": losses,
                "game_pts": game_pts, "adv_bonus": adv_bonus,
                "pts": game_pts + adv_bonus,
                "advanced": adv, "picks": pick_count, "pct": pct,
            })

    status_order = {"champion": 0, "advanced": 1, "active": 1, "not_played": 2, "out": 3}
    all_teams_data.sort(key=lambda t: (status_order[t["status"]], t["tier"], -t["picks"]))

    teams_rows = ""
    prev_sg    = None
    for t in all_teams_data:
        sg = "out" if t["status"] == "out" else "alive"
        if sg == "out" and prev_sg != "out":
            teams_rows += '<tr class="elim-divider"><td colspan="10">— Eliminated —</td></tr>'
        prev_sg = sg
        if t["status"] == "champion":
            row_cls, status_html = "tr-champ", '<span class="ts-champ">🏆 Champion</span>'
        elif t["status"] == "advanced":
            row_cls, status_html = "tr-alive", '<span class="ts-alive">✓ Advanced</span>'
        elif t["status"] == "active":
            row_cls, status_html = "tr-alive", '<span class="ts-alive">● Active</span>'
        elif t["status"] == "not_played":
            row_cls, status_html = "tr-notplayed", '<span class="ts-notplayed">· Pre-tournament</span>'
        else:
            row_cls, status_html = "tr-out", '<span class="ts-out">✗ Out</span>'
        color    = TIER_COLORS.get(t["tier"], "#64748B")
        tier_pill = (f'<span class="tier-pill" '
                     f'style="background:{TIER_BG[t["tier"]]};color:{color};border:1px solid {color}40">'
                     f'{t["tier"]}</span>')
        gp           = t["wins"] + t["ties"] + t["losses"]
        record       = f'{t["wins"]}W-{t["ties"]}D-{t["losses"]}L'
        gp_str       = str(gp) if gp > 0 else '—'
        game_pts_str = f'{t["game_pts"]:,}' if (t["game_pts"] or gp > 0) else '—'
        adv_pts_str  = f'+{t["adv_bonus"]:,}' if t["adv_bonus"] else '—'
        total_str    = f'{t["pts"]:,}' if (t["pts"] or gp > 0) else '—'
        teams_rows += f"""
        <tr class="{row_cls}">
          <td class="text-center">{tier_pill}</td>
          <td class="team-name-cell">{_h(t['team'])}</td>
          <td class="text-center">{status_html}</td>
          <td class="text-center text-muted small">{record}</td>
          <td class="text-center">{gp_str}</td>
          <td class="text-end">{game_pts_str}</td>
          <td class="text-end" style="color:var(--green)">{adv_pts_str}</td>
          <td class="text-end fw-semibold">{total_str}</td>
          <td class="text-center fw-semibold">{t['picks']}</td>
          <td class="text-end text-muted">{t['pct']}%</td>
        </tr>"""

    # ── Build new tabs ─────────────────────────────────────────────────────
    groups_tab  = _build_group_tab(group_standings)
    bracket_tab = _build_bracket_tab(round_matches)
    scorers_tab = _build_scorers_tab(standings, scorer_goals, entries)
    raw_tab     = _build_raw_data_tab(standings, team_results, advanced,
                                      champion, knocked_out, scorer_goals,
                                      rank_counts=rank_counts)

    js_pop = json.dumps(pop, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{CHALLENGE_NAME}</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Bangers&family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
  <style>
    :root {{
      --pitch:   #064E3B;
      --pitch2:  #065F46;
      --pitch3:  #047857;
      --gold:    #F59E0B;
      --gold2:   #D97706;
      --white15: rgba(255,255,255,.15);
      --white30: rgba(255,255,255,.3);
      --green:   #16A34A;
      --lgreen:  #22C55E;
      --bg:      #F0FDF4;
      --card:    #FFFFFF;
      --border:  #D1FAE5;
      --text:    #1E293B;
      --muted:   #64748B;
    }}
    *, body {{ font-family:'Inter',sans-serif; color:var(--text); }}
    body {{ background:#ECFDF5; }}

    /* ── HERO ────────────────────────────────────────────── */
    .hero {{
      position:relative;
      background: linear-gradient(to bottom, #064E3B 0%, #053D2E 100%);
      padding:32px 0 26px;
      color:white;
      text-align:center;
      overflow:hidden;
    }}
    .hero-ball  {{ font-size:2.8rem; line-height:1; margin-bottom:10px; filter:drop-shadow(0 2px 8px rgba(0,0,0,.4)); }}
    .hero-title {{
      font-family:'Bangers',cursive;
      font-size:clamp(2rem,6vw,4rem);
      letter-spacing:4px;
      text-shadow:3px 4px 0 rgba(0,0,0,.4), 0 0 40px rgba(245,158,11,.3);
      line-height:1; margin-bottom:6px;
      background:linear-gradient(135deg, #FFFFFF 30%, #FDE68A 100%);
      -webkit-background-clip:text; -webkit-text-fill-color:transparent;
    }}
    .hero-sub     {{ font-size:.9rem; opacity:.7; margin-top:8px; letter-spacing:.5px; color:white; }}
    .hero-updated {{ font-size:.72rem; opacity:.5; margin-top:5px; color:white; }}
    .hero-updated span {{ color:white; }}
    .stat-card {{
      background:rgba(255,255,255,.1);
      border:1px solid rgba(255,255,255,.18);
      backdrop-filter:blur(8px);
      border-radius:12px; padding:18px 14px; text-align:center; color:white;
    }}
    .sc-label {{ font-size:.65rem; font-weight:700; text-transform:uppercase;
                 letter-spacing:1px; opacity:.65; margin-bottom:6px; }}
    .sc-value {{ font-family:'Bangers',cursive; font-size:2.6rem; line-height:1; }}
    .sc-sub   {{ font-size:.72rem; opacity:.55; margin-top:3px; }}

    /* ── TAB NAVIGATION ──────────────────────────────────── */
    .tabs-wrap {{
      background:var(--pitch);
      position:sticky; top:0; z-index:100;
      box-shadow:0 2px 12px rgba(0,0,0,.3);
    }}
    .nav-tabs {{
      border-bottom:none;
      gap:0;
      flex-wrap:nowrap;
      overflow-x:auto;
      -webkit-overflow-scrolling:touch;
      scrollbar-width:none;
    }}
    .nav-tabs::-webkit-scrollbar {{ display:none; }}
    .nav-tabs .nav-link {{
      font-weight:700; font-size:.78rem; color:rgba(255,255,255,.6);
      border:none; border-bottom:3px solid transparent;
      padding:14px 16px; border-radius:0; background:none;
      white-space:nowrap; letter-spacing:.3px;
      transition:all .15s;
    }}
    .nav-tabs .nav-link:hover  {{ color:var(--gold); background:rgba(255,255,255,.06); }}
    .nav-tabs .nav-link.active {{
      color:var(--gold);
      border-bottom-color:var(--gold);
      background:rgba(255,255,255,.08);
    }}
    .tab-content {{ padding-top:0; }}

    /* ── STANDINGS STATS BAR ─────────────────────────────── */
    .standings-stats-bar {{
      background:
        linear-gradient(to bottom, rgba(6,78,59,.55) 0%, rgba(6,78,59,.3) 50%, rgba(6,78,59,.55) 100%),
        url('images/standings_bg.jpg') center/cover no-repeat;
      padding:22px 0 20px;
      border-bottom:1px solid var(--pitch3);
    }}

    /* ── TAB HERO STRIPS ─────────────────────────────────── */
    .tab-hero-lg {{
      min-height:160px;
      background:
        linear-gradient(to bottom, rgba(6,78,59,.45) 0%, rgba(6,78,59,.18) 50%, rgba(6,78,59,.45) 100%),
        var(--tab-bg, linear-gradient(135deg,#064E3B,#065F46)) center/cover no-repeat;
      display:flex; align-items:center; justify-content:center; text-align:center;
    }}
    .tab-hero-lg .tab-hero-inner {{
      padding:18px 24px; text-align:center;
    }}
    .tab-hero-lg .tab-hero-title {{
      font-size:clamp(2rem,5vw,3.2rem);
      text-shadow:2px 3px 0 rgba(0,0,0,.5);
    }}
    .tab-hero-lg .tab-hero-sub {{
      font-size:.88rem; margin-top:8px;
    }}
    .tab-hero-inner {{ padding:0 20px; }}
    .tab-hero-title {{
      font-family:'Bangers',cursive;
      font-size:clamp(1.4rem,3.5vw,2.2rem);
      color:white; letter-spacing:2px;
      text-shadow:1px 2px 0 rgba(0,0,0,.4);
      line-height:1.1;
    }}
    .tab-hero-sub {{ font-size:.78rem; color:rgba(255,255,255,.65); margin-top:4px; }}

    /* ── CONTAINER INNER ─────────────────────────────────── */
    .container-inner {{ max-width:1200px; margin:0 auto; padding:24px 16px 40px; }}

    /* ── LEADERBOARD ─────────────────────────────────────── */
    #tab-standings {{ background:var(--bg); }}
    .lb-wrap {{
      border-radius:12px;
      box-shadow:0 4px 20px rgba(0,0,0,.08);
      overflow-x:auto; -webkit-overflow-scrolling:touch;
    }}
    .lb-wrap table {{ min-width:520px; }}
    .lb-wrap thead th {{
      font-size:.7rem; font-weight:700; text-transform:uppercase;
      letter-spacing:.5px; color:white; padding:13px 14px;
      background:var(--pitch); border-bottom:2px solid var(--pitch3);
      white-space:nowrap;
    }}
    .lb-row {{ cursor:pointer; transition:background .1s; background:var(--card); }}
    .lb-row:hover {{ background:#F0FDF4 !important; }}
    .lb-row td {{ padding:11px 14px; vertical-align:middle; border-bottom:1px solid var(--border); }}
    .row-gold   {{ box-shadow:inset 4px 0 0 var(--gold); background:#FFFBEB !important; }}
    .row-silver {{ box-shadow:inset 4px 0 0 #94A3B8; }}
    .row-bronze {{ box-shadow:inset 4px 0 0 #B45309; }}
    .rank-badge {{
      display:inline-flex; align-items:center; justify-content:center;
      min-width:30px; height:30px; padding:0 5px; border-radius:6px; font-weight:800; font-size:.85rem;
      white-space:nowrap;
    }}
    .rank-1 {{ background:#FEF3C7; color:#92400E; border:1px solid var(--gold); }}
    .rank-2 {{ background:#F1F5F9; color:#475569; border:1px solid #CBD5E1; }}
    .rank-3 {{ background:#FEF9F0; color:#92400E; border:1px solid #D97706; }}
    .rank-n {{ background:#ECFDF5; color:var(--green); border:1px solid #A7F3D0; font-size:.72rem; }}
    .p-name  {{ display:block; font-weight:700; font-size:.92rem; }}
    .p-email {{ display:block; font-size:.67rem; color:var(--muted); }}
    .fire-tag {{ font-size:.8rem; }}
    .pts-val {{ font-weight:800; font-size:1.05rem; color:var(--pitch); }}
    .badge-alive {{
      display:inline-block; padding:3px 9px; border-radius:20px; font-size:.72rem;
      font-weight:700; background:#DCFCE7; color:#15803D; border:1px solid #BBF7D0;
    }}
    .badge-dead {{
      display:inline-block; padding:3px 9px; border-radius:20px; font-size:.72rem;
      font-weight:600; background:#F8FAFC; color:var(--muted); border:1px solid var(--border);
    }}
    .still-in-col {{ font-size:.78rem; color:var(--muted); }}
    .picks-row {{ display:none; }}
    .picks-row.open {{ display:table-row; }}
    .picks-row td {{ padding:10px 14px 14px; background:#F8FAFC !important; border-bottom:1px solid var(--border); }}
    .picks-inner {{ display:flex; flex-wrap:wrap; gap:5px; }}
    .pick-chip {{
      display:inline-flex; align-items:center; gap:4px;
      padding:4px 10px; border-radius:20px; font-size:.73rem; font-weight:600;
    }}
    .pick-chip em {{ font-style:normal; opacity:.65; font-size:.67rem; font-weight:500; }}
    .chip-tier {{
      display:inline-block; padding:1px 5px; border-radius:4px;
      font-size:.62rem; font-weight:800; letter-spacing:.3px;
    }}
    .chip-alive   {{ background:#DCFCE7; color:#15803D; }}
    .chip-adv     {{ background:#DBEAFE; color:#1E40AF; }}
    .chip-winner  {{ background:#FEF3C7; color:#92400E; border:1px solid var(--gold); }}
    .chip-out     {{ background:#FEE2E2; color:#B91C1C; }}
    .chip-zero    {{ background:#F1F5F9; color:#64748B; border:1px solid #CBD5E1; }}
    .chip-scorer  {{ background:#FFF7ED; color:#C2410C; }}
    .chip-winner-pick {{ border:1px dashed var(--gold); }}

    #lb-search, .lb-search-input {{
      border:1px solid var(--border); border-radius:8px;
      padding:8px 14px; font-size:.85rem; width:100%; max-width:280px;
      outline:none; background:var(--card);
    }}
    #lb-search:focus, .lb-search-input:focus {{
      border-color:var(--green); box-shadow:0 0 0 3px rgba(22,163,74,.12);
    }}

    /* ── GROUP STANDINGS ─────────────────────────────────── */
    #tab-groups {{ background:#ECFDF5; }}
    .gs-grid {{
      display:grid;
      grid-template-columns:repeat(auto-fill, minmax(320px, 1fr));
      gap:16px;
    }}
    .gs-card {{
      background:var(--card); border-radius:12px;
      box-shadow:0 2px 10px rgba(0,0,0,.06);
      border:1px solid var(--border); overflow:hidden;
    }}
    .gs-card-head {{
      background:var(--pitch); color:white;
      font-family:'Bangers',cursive; font-size:1.1rem;
      letter-spacing:1.5px; padding:10px 14px;
    }}
    .gs-table {{ width:100%; border-collapse:collapse; font-size:.82rem; }}
    .gs-table thead th {{
      font-size:.65rem; font-weight:700; text-transform:uppercase;
      letter-spacing:.4px; color:var(--muted); padding:7px 8px;
      border-bottom:1px solid var(--border); text-align:center;
    }}
    .gs-table thead th:nth-child(2) {{ text-align:left; }}
    .gs-table tbody td {{ padding:8px 8px; border-bottom:1px solid #F1F5F9; }}
    .gs-table tbody tr:last-child td {{ border-bottom:none; }}
    .gs-pos {{ color:var(--muted); font-size:.75rem; text-align:center; width:24px; }}
    .gs-team {{ font-weight:600; }}
    .gs-pts {{ font-weight:800; color:var(--pitch); }}
    .gs-adv {{ background:#F0FDF4; }}
    .gs-adv .gs-pos {{ color:var(--green); font-weight:700; }}
    .gs-legend {{ font-size:.72rem; color:var(--muted); margin-top:12px; display:flex; align-items:center; gap:6px; }}
    .gs-adv-dot {{ display:inline-block; width:10px; height:10px; background:#DCFCE7; border:1px solid #86EFAC; border-radius:2px; }}

    /* ── BRACKET ─────────────────────────────────────────── */
    #tab-bracket {{ background:#ECFDF5; }}
    .bk-round {{ margin-bottom:28px; }}
    .bk-round-label {{
      font-family:'Bangers',cursive; font-size:1.3rem; color:var(--pitch);
      letter-spacing:2px; border-left:4px solid var(--gold);
      padding-left:12px; margin-bottom:14px;
      display:flex; align-items:baseline; gap:10px;
    }}
    .bk-round-count {{ font-family:'Inter',sans-serif; font-size:.72rem; color:var(--muted); font-weight:600; letter-spacing:0; }}
    .bk-cards {{
      display:grid;
      grid-template-columns:repeat(auto-fill, minmax(200px, 1fr));
      gap:10px;
    }}
    .bk-card {{
      background:var(--card); border-radius:10px;
      box-shadow:0 2px 8px rgba(0,0,0,.07);
      border:1px solid var(--border); overflow:hidden;
    }}
    .bk-match-row {{
      display:flex; align-items:center; justify-content:space-between;
      padding:9px 12px; gap:6px;
    }}
    .bk-team {{ font-size:.82rem; font-weight:600; }}
    .bk-score {{ font-weight:800; font-size:.9rem; color:var(--muted); }}
    .bk-winner {{ background:#F0FDF4; }}
    .bk-winner .bk-team {{ color:var(--pitch); }}
    .bk-winner .bk-score {{ color:var(--green); }}
    .bk-loser  {{ opacity:.5; }}
    .bk-divider {{ height:1px; background:var(--border); margin:0 12px; }}
    .bk-date {{ font-size:.65rem; color:var(--muted); text-align:center; padding:4px 0 6px; }}
    .bk-placeholder {{ opacity:.45; border:1px dashed #CBD5E1 !important; background:#F8FAFC !important; }}
    .bk-tbd {{ color:#94A3B8 !important; font-style:italic; font-weight:500 !important; }}
    .bk-tbd-score {{ color:#CBD5E1 !important; }}

    /* ── GOAL SCORERS ────────────────────────────────────── */
    #tab-scorers {{ background:#ECFDF5; }}
    .sc-summary {{
      display:grid; grid-template-columns:repeat(3,1fr); gap:12px; margin-bottom:24px;
    }}
    .sc-sum-card {{
      background:var(--pitch); color:white; border-radius:10px;
      padding:18px 16px; text-align:center;
    }}
    .sc-sum-val {{ font-family:'Bangers',cursive; font-size:2.4rem; line-height:1; }}
    .sc-sum-lbl {{ font-size:.7rem; opacity:.7; margin-top:4px; letter-spacing:.5px; text-transform:uppercase; }}
    .sc-list {{ display:flex; flex-direction:column; gap:10px; }}
    .sc-row {{
      background:var(--card); border-radius:10px;
      border:1px solid var(--border);
      box-shadow:0 1px 6px rgba(0,0,0,.05);
      padding:12px 16px;
      display:flex; align-items:center; gap:14px;
    }}
    .sc-row-gold   {{ box-shadow:inset 4px 0 0 var(--gold), 0 2px 8px rgba(0,0,0,.08); background:#FFFBEB; }}
    .sc-row-silver {{ box-shadow:inset 4px 0 0 #94A3B8; }}
    .sc-row-bronze {{ box-shadow:inset 4px 0 0 #B45309; }}
    .sc-rank {{ flex-shrink:0; }}
    .sc-info {{ flex:1; min-width:0; }}
    .sc-name {{ font-weight:700; font-size:.95rem; }}
    .sc-pickers {{ font-size:.72rem; color:var(--muted); margin-top:2px; }}
    .sc-expand-btn {{
      background:none; border:none; padding:0 4px; cursor:pointer;
      font-size:.72rem; color:var(--green); font-weight:700;
    }}
    .sc-expand-btn:hover {{ text-decoration:underline; }}
    .sc-pickers-full {{
      margin-top:6px; display:flex; flex-wrap:wrap; gap:4px;
    }}
    .sc-picker-chip {{
      display:inline-block; padding:2px 8px; border-radius:20px;
      background:#ECFDF5; color:var(--pitch); font-size:.68rem; font-weight:600;
      border:1px solid #A7F3D0;
    }}
    .goal-bar-wrap {{ height:4px; background:#E2E8F0; border-radius:2px; margin-top:7px; }}
    .goal-bar {{ height:4px; background:var(--green); border-radius:2px; transition:width .4s; }}
    .sc-stats {{ flex-shrink:0; text-align:right; }}
    .sc-goals-badge {{
      display:block; font-weight:800; font-size:1rem; color:var(--pitch);
    }}
    .sc-no-goal {{ color:var(--muted); font-weight:600; }}
    .sc-pts {{ display:block; font-size:.72rem; font-weight:700; color:var(--green); margin-top:2px; }}
    .sc-zero {{ color:var(--muted); }}
    .sc-pickcount {{ display:block; font-size:.68rem; color:var(--muted); margin-top:3px; }}

    /* ── TEAM TRACKER ────────────────────────────────────── */
    #tab-teams {{ background:#ECFDF5; }}
    .tracker-card {{
      background:var(--card); border-radius:12px; overflow:hidden;
      box-shadow:0 2px 10px rgba(0,0,0,.06); border:1px solid var(--border);
    }}
    .tracker-card thead th {{
      font-size:.7rem; font-weight:700; text-transform:uppercase; letter-spacing:.4px;
      color:white; padding:12px 10px; background:var(--pitch);
      border-bottom:2px solid var(--pitch3); white-space:nowrap;
      cursor:pointer; user-select:none;
    }}
    .tracker-card thead th:hover {{ background:var(--pitch3); }}
    .tracker-card tbody td {{ padding:9px 10px; border-bottom:1px solid var(--border); font-size:.85rem; }}
    .tr-champ td    {{ background:#FFFBEB; }}
    .tr-out td      {{ color:#CBD5E1; }}
    .tr-out .team-name-cell {{ text-decoration:line-through; text-decoration-color:#CBD5E1; }}
    .tr-notplayed td {{ color:var(--muted); }}
    .ts-champ     {{ color:#92400E; font-weight:700; font-size:.8rem; }}
    .ts-alive     {{ color:var(--green); font-weight:700; font-size:.8rem; }}
    .ts-out       {{ color:#CBD5E1; font-weight:600; font-size:.8rem; }}
    .ts-notplayed {{ color:var(--muted); font-weight:500; font-size:.8rem; }}
    .team-name-cell {{ font-weight:600; }}
    .tier-pill {{
      display:inline-block; padding:2px 8px; border-radius:10px;
      font-size:.72rem; font-weight:800;
    }}
    .elim-divider td {{
      text-align:center; font-size:.7rem; font-weight:600; text-transform:uppercase;
      letter-spacing:2px; color:var(--muted); padding:7px;
      background:#F8FAFC !important; border-bottom:1px solid var(--border);
    }}

    /* ── CHARTS ──────────────────────────────────────────── */
    #tab-picks {{ background:#ECFDF5; }}
    .chart-card {{
      background:var(--card); border-radius:10px; padding:14px 16px; height:100%;
      border:1px solid var(--border); box-shadow:0 1px 4px rgba(0,0,0,.05);
    }}
    .chart-label {{ display:flex; justify-content:space-between; align-items:baseline; margin-bottom:10px; }}
    .tier-tag {{ font-weight:800; font-size:.88rem; }}
    .tier-sub {{ font-size:.72rem; color:var(--muted); font-weight:600; }}

    /* ── RAW DATA ────────────────────────────────────────── */
    #tab-raw {{ background:#ECFDF5; }}
    .rp-wrap {{ border-radius:12px; overflow:auto; box-shadow:0 4px 16px rgba(0,0,0,.07); background:var(--card); max-height:75vh; }}
    #rp-table {{ border-collapse:separate; border-spacing:0; }}
    #rp-table thead {{ position:sticky; top:0; z-index:10; }}
    #rp-table thead th {{
      background:var(--pitch); font-size:.65rem; font-weight:700; text-transform:uppercase;
      letter-spacing:.4px; color:white; padding:9px 6px; white-space:nowrap;
      border-bottom:none; user-select:none;
    }}
    /* Sticky left columns: # → Name → Pts */
    .rp-rank-hdr {{
      width:36px; position:sticky; left:0; z-index:11;
      background:var(--pitch); text-align:center !important;
    }}
    .rp-name-hdr {{
      min-width:130px; position:sticky; left:36px; z-index:11;
      background:var(--pitch);
    }}
    .rp-pts-hdr {{
      width:54px; position:sticky; left:166px; z-index:11;
      background:var(--pitch); text-align:right !important;
      border-right:2px solid rgba(255,255,255,.25) !important;
    }}
    .rp-tier-hdr {{ min-width:90px; }}
    .rp-row td {{ padding:5px 6px; border-bottom:1px solid var(--border); font-size:.78rem; vertical-align:middle; text-align:center; }}
    .rp-row:hover td {{ background:#F0FDF4 !important; }}
    .rp-rank-cell {{
      position:sticky; left:0; z-index:5; background:var(--card);
      width:36px; padding:4px 6px !important; text-align:center !important;
    }}
    .rp-name-cell {{
      position:sticky; left:36px; z-index:5; background:var(--card);
      min-width:130px; padding:6px 8px !important; white-space:nowrap; text-align:left !important;
    }}
    .rp-pts-cell {{
      position:sticky; left:166px; z-index:5; background:var(--card);
      width:54px; padding:6px 6px !important; text-align:right !important;
      font-weight:700; font-size:.8rem; color:var(--pitch);
      border-right:2px solid var(--border) !important;
    }}
    .rp-row:hover .rp-rank-cell,
    .rp-row:hover .rp-name-cell,
    .rp-row:hover .rp-pts-cell {{ background:#F0FDF4 !important; }}
    .rp-pname {{ font-weight:700; font-size:.82rem; }}
    .rp-pts   {{ display:block; font-size:.6rem; font-weight:700; opacity:.7; margin-top:1px; }}
    td.rp-alive     {{ background:#F0FDF4; color:#15803D; font-weight:600; }}
    td.rp-champ     {{ background:#FFFBEB; color:#92400E; font-weight:700; border:1px solid var(--gold); }}
    td.rp-out       {{ background:#FEF2F2; color:#B91C1C; font-weight:500; opacity:.8; }}
    td.rp-notplayed {{ background:#F8FAFC; color:var(--muted); }}
    td.rp-empty     {{ background:#F8FAFC; color:#CBD5E1; }}
    .rp-legend {{ display:inline-block; padding:2px 8px; border-radius:10px; font-size:.7rem; font-weight:600; margin-left:4px; }}
    span.rp-alive-leg {{ background:#DCFCE7; color:#15803D; }}
    span.rp-out-leg   {{ background:#FEE2E2; color:#B91C1C; }}
    span.rp-champ-leg {{ background:#FEF3C7; color:#92400E; }}

    /* ── SHARED ──────────────────────────────────────────── */
    .empty-state {{
      text-align:center; padding:60px 20px; color:var(--muted);
      font-size:.92rem; line-height:1.8;
      background:var(--card); border-radius:12px;
      border:1px dashed var(--border);
    }}
    .stale-banner {{
      background:#FEF3C7; border:1px solid var(--gold); color:#92400E;
      border-radius:8px; padding:10px 16px; margin:16px 0;
      font-size:.83rem; font-weight:600; text-align:center;
    }}
    footer {{
      text-align:center; padding:28px; font-size:.75rem; color:rgba(255,255,255,.5);
      background:var(--pitch); border-top:1px solid var(--pitch3); margin-top:0;
    }}
    footer a {{ color:var(--gold); }}
    code {{ font-size:.85em; }}
  </style>
</head>
<body>

<!-- ═══ HERO ═══════════════════════════════════════════════ -->
<div class="hero">
  <div class="container">
    <div class="hero-ball">⚽</div>
    <div class="hero-title">{_h(CHALLENGE_NAME)}</div>
    <div class="hero-sub">Live Standings &amp; Stats</div>
    <div class="hero-updated">Updated <span id="ts" data-utc="{last_updated}"></span></div>
  </div>
</div>

<!-- ═══ NAVIGATION ═════════════════════════════════════════ -->
<div class="tabs-wrap">
  <div class="container">
    <ul class="nav nav-tabs" id="mainTabs" role="tablist">
      <li class="nav-item" role="presentation">
        <button class="nav-link active" data-bs-toggle="tab" data-bs-target="#tab-standings"
                type="button" role="tab">🏆 Standings</button>
      </li>
      <li class="nav-item" role="presentation">
        <button class="nav-link" data-bs-toggle="tab" data-bs-target="#tab-raw"
                type="button" role="tab">📋 All Picks</button>
      </li>
      <li class="nav-item" role="presentation">
        <button class="nav-link" data-bs-toggle="tab" data-bs-target="#tab-teams"
                type="button" role="tab">🌍 Teams</button>
      </li>
      <li class="nav-item" role="presentation">
        <button class="nav-link" data-bs-toggle="tab" data-bs-target="#tab-scorers"
                type="button" role="tab">⚽ Goal Scorers</button>
      </li>
      <li class="nav-item" role="presentation">
        <button class="nav-link" id="tab-picks-btn" data-bs-toggle="tab" data-bs-target="#tab-picks"
                type="button" role="tab">📊 Popularity</button>
      </li>
      <li class="nav-item" role="presentation">
        <button class="nav-link" data-bs-toggle="tab" data-bs-target="#tab-groups"
                type="button" role="tab">🗂 Groups</button>
      </li>
      <li class="nav-item" role="presentation">
        <button class="nav-link" data-bs-toggle="tab" data-bs-target="#tab-bracket"
                type="button" role="tab">🥊 Bracket</button>
      </li>
    </ul>
  </div>
</div>

<div class="tab-content" id="mainTabContent">

  <!-- ═══ STANDINGS ══════════════════════════════════════════ -->
  <div class="tab-pane fade show active" id="tab-standings" role="tabpanel">
    <div class="standings-stats-bar">
      <div class="container-inner" style="padding-top:0;padding-bottom:0;">
        <div class="row g-3 justify-content-center">
          <div class="col-6 col-sm-3">
            <div class="stat-card">
              <div class="sc-label">Entries</div>
              <div class="sc-value">{total_p}</div>
              <div class="sc-sub">participants</div>
            </div>
          </div>
          <div class="col-6 col-sm-3">
            <div class="stat-card">
              <div class="sc-label">Leader</div>
              <div class="sc-value" style="font-size:1.6rem;line-height:1.3">{_h(leader['name'])}</div>
              <div class="sc-sub">{leader['total_points']:,} pts</div>
            </div>
          </div>
          <div class="col-6 col-sm-3">
            <div class="stat-card">
              <div class="sc-label">Games Played</div>
              <div class="sc-value">{total_games}</div>
              <div class="sc-sub">of 104</div>
            </div>
          </div>
          <div class="col-6 col-sm-3">
            <div class="stat-card">
              <div class="sc-label">Teams In</div>
              <div class="sc-value">{alive_count}</div>
              <div class="sc-sub">of 48</div>
            </div>
          </div>
        </div>
      </div>
    </div>
    <div class="container-inner">
      {stale}
      <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
        <p class="text-muted small mb-0">
          <span class="pick-chip chip-alive" style="font-size:.68rem">● active</span>
          <span class="pick-chip chip-zero"  style="font-size:.68rem">– played/0pts</span>
          <span class="pick-chip chip-adv"   style="font-size:.68rem">✓ advanced</span>
          <span class="pick-chip chip-winner" style="font-size:.68rem">🏆 champion</span>
          <span class="pick-chip chip-out"   style="font-size:.68rem">✗ out</span>
        </p>
        <input id="lb-search" type="text" placeholder="🔍  Search participant…" oninput="filterTable()">
      </div>
      <div class="lb-wrap">
        <table class="table mb-0" id="lb-table">
          <thead>
            <tr>
              <th>#</th><th>Participant</th>
              <th class="text-end">Points</th>
              <th class="text-center">Played</th>
              <th class="text-center">Alive</th>
              <th class="d-none d-lg-table-cell">Teams Still In</th>
            </tr>
          </thead>
          <tbody>{rows_html}</tbody>
        </table>
      </div>
    </div>
  </div>

  {raw_tab}

  <!-- ═══ TEAM TRACKER ═══════════════════════════════════════ -->
  <div class="tab-pane fade" id="tab-teams" role="tabpanel">
    <div class="tab-hero-lg" style="--tab-bg:url('images/teams_bg.jpg')">
      <div class="tab-hero-inner">
        <div class="tab-hero-title">🌍 Team Tracker</div>
        <div class="tab-hero-sub">All 48 teams · status, record, and pick popularity</div>
      </div>
    </div>
    <div class="container-inner">
      <div class="tracker-card" style="overflow-x:auto;-webkit-overflow-scrolling:touch;">
        <div style="min-width:640px;">
          <table class="table mb-0" id="teams-table">
            <thead>
              <tr>
                <th class="text-center" onclick="sortTeams(0)">Tier ↕</th>
                <th onclick="sortTeams(1)">Team ↕</th>
                <th class="text-center" onclick="sortTeams(2)">Status ↕</th>
                <th class="text-center" onclick="sortTeams(3)">Record ↕</th>
                <th class="text-center" onclick="sortTeams(4)">Played ↕</th>
                <th class="text-end"    onclick="sortTeams(5)">Game Pts ↕</th>
                <th class="text-end"    onclick="sortTeams(6)">Adv Pts ↕</th>
                <th class="text-end"    onclick="sortTeams(7)">Total ↕</th>
                <th class="text-center" onclick="sortTeams(8)">Picked ↕</th>
                <th class="text-end"    onclick="sortTeams(9)">% ↕</th>
              </tr>
            </thead>
            <tbody id="teams-body">{teams_rows}</tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  {scorers_tab}

  <!-- ═══ PICK POPULARITY ════════════════════════════════════ -->
  <div class="tab-pane fade" id="tab-picks" role="tabpanel">
    <div class="tab-hero-lg" style="--tab-bg:url('images/popularity_bg.jpg')">
      <div class="tab-hero-inner">
        <div class="tab-hero-title">📊 Pick Popularity</div>
        <div class="tab-hero-sub">How the pool split across tiers, winner, and scorers</div>
      </div>
    </div>
    <div class="container-inner">
      <div class="row g-4">
        <div class="col-12 col-lg-6">
          <div class="chart-card">
            <div class="chart-label">
              <span class="tier-tag" style="color:{TIER_COLORS['A']}">Tier A — Favorites</span>
              <span class="tier-sub">Pick 2 · +100pts advancement</span>
            </div>
            <canvas id="cA" height="160"></canvas>
          </div>
        </div>
        <div class="col-12 col-lg-6">
          <div class="chart-card">
            <div class="chart-label">
              <span class="tier-tag" style="color:{TIER_COLORS['B']}">Tier B — Dark Horses</span>
              <span class="tier-sub">Pick 4 · +100pts advancement</span>
            </div>
            <canvas id="cB" height="160"></canvas>
          </div>
        </div>
        <div class="col-12 col-lg-6">
          <div class="chart-card">
            <div class="chart-label">
              <span class="tier-tag" style="color:{TIER_COLORS['C']}">Tier C — Longshots</span>
              <span class="tier-sub">Pick 4 · +200pts advancement</span>
            </div>
            <canvas id="cC" height="160"></canvas>
          </div>
        </div>
        <div class="col-12 col-lg-6">
          <div class="chart-card">
            <div class="chart-label">
              <span class="tier-tag" style="color:{TIER_COLORS['D']}">Tier D — We're Glad You Came</span>
              <span class="tier-sub">Pick 2 · +200pts advancement</span>
            </div>
            <canvas id="cD" height="160"></canvas>
          </div>
        </div>
        <div class="col-12 col-md-6">
          <div class="chart-card">
            <div class="chart-label">
              <span class="tier-tag" style="color:#16A34A">⭐ Tournament Winner Picks</span>
              <span class="tier-sub">+450 pts bonus</span>
            </div>
            <canvas id="cW" height="160"></canvas>
          </div>
        </div>
        <div class="col-12 col-md-6">
          <div class="chart-card">
            <div class="chart-label">
              <span class="tier-tag" style="color:#EA580C">⚽ Goal Scorer Picks</span>
              <span class="tier-sub">+150 pts/goal</span>
            </div>
            <canvas id="cS" height="220"></canvas>
          </div>
        </div>
      </div>
    </div>
  </div>

  {groups_tab}
  {bracket_tab}

</div><!-- /tab-content -->

<footer>
  Data sourced from ESPN &bull;
  Auto-updates every 2 hours via GitHub Actions &bull;
  <a href="https://docs.github.com/en/pages">Hosted on GitHub Pages</a>
</footer>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
<script>
const POP = {js_pop};

// ── Timestamp ────────────────────────────────────────────────
(function(){{
  const el = document.getElementById('ts');
  if (!el) return;
  const d = new Date(el.dataset.utc);
  el.textContent = d.toLocaleString('en-US',{{month:'long',day:'numeric',year:'numeric',hour:'numeric',minute:'2-digit',hour12:true}});
}})();

// ── Leaderboard row expand ───────────────────────────────────
function toggleRow(id) {{
  document.getElementById(id)?.classList.toggle('open');
}}

function toggleScPickers(id, btn) {{
  const el = document.getElementById(id);
  if (!el) return;
  const open = el.style.display === 'none';
  el.style.display = open ? 'flex' : 'none';
  const count = btn.textContent.match(/\+(\d+)/)?.[1] || '';
  btn.textContent = open ? `▴ collapse` : `+${{count}} more ▾`;
}}

// ── Search / filter ──────────────────────────────────────────
function filterTable() {{
  const q = document.getElementById('lb-search').value.toLowerCase();
  document.querySelectorAll('#lb-table tbody .lb-row').forEach(tr => {{
    const show = tr.innerText.toLowerCase().includes(q);
    tr.style.display = show ? '' : 'none';
    const next = tr.nextElementSibling;
    if (next?.classList.contains('picks-row')) {{
      if (!show) next.classList.remove('open');
      next.style.display = show ? '' : 'none';
    }}
  }});
}}

function filterRaw() {{
  const q = document.getElementById('rp-search').value.toLowerCase();
  document.querySelectorAll('#rp-body .rp-row').forEach(tr => {{
    tr.style.display = tr.dataset.name.includes(q) ? '' : 'none';
  }});
}}

// ── Team tracker sort ────────────────────────────────────────
const _sd = {{}};
function sortTeams(col) {{
  const tbody = document.getElementById('teams-body');
  const rows  = Array.from(tbody.querySelectorAll('tr:not(.elim-divider)'));
  _sd[col] = !_sd[col];
  const dir = _sd[col] ? 1 : -1;
  rows.sort((a, b) => {{
    const av = (a.cells[col]?.innerText||'').replace(/[^0-9.\-WDL%]/g,'');
    const bv = (b.cells[col]?.innerText||'').replace(/[^0-9.\-WDL%]/g,'');
    const an = parseFloat(av), bn = parseFloat(bv);
    return isNaN(an)||isNaN(bn)
      ? dir*(a.cells[col]?.innerText||'').localeCompare(b.cells[col]?.innerText||'')
      : dir*(an-bn);
  }});
  const div = tbody.querySelector('.elim-divider');
  rows.forEach(r => r.remove()); if (div) div.remove();
  const alive = rows.filter(r => !r.classList.contains('tr-out'));
  const out   = rows.filter(r =>  r.classList.contains('tr-out'));
  alive.forEach(r => tbody.appendChild(r));
  if (div && out.length) tbody.appendChild(div);
  out.forEach(r => tbody.appendChild(r));
}}

// ── Charts (lazy init) ───────────────────────────────────────
document.getElementById('tab-picks-btn')?.addEventListener('shown.bs.tab', () => {{
  if (!window._chartsInit) {{ initCharts(); window._chartsInit = true; }}
}});

function makeBar(id, labels, counts, color) {{
  const el = document.getElementById(id);
  if (!el || !labels.length) return;
  const total = counts.reduce((a,b)=>a+b,0)||1;
  Chart.register(ChartDataLabels);
  new Chart(el.getContext('2d'), {{
    type:'bar',
    data:{{labels,datasets:[{{data:counts,backgroundColor:color+'22',borderColor:color,borderWidth:1.5,borderRadius:4}}]}},
    options:{{
      indexAxis:'y', responsive:true, layout:{{padding:{{right:48}}}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{callbacks:{{label:c=>` ${{c.parsed.x}} picks (${{Math.round(c.parsed.x/total*100)}}%)`}}}},
        datalabels:{{anchor:'end',align:'right',color:'#64748B',font:{{size:11,weight:'600'}},formatter:v=>Math.round(v/total*100)+'%',padding:{{left:4}}}},
      }},
      scales:{{
        x:{{beginAtZero:true,max:Math.max(...counts,1)*1.35,ticks:{{stepSize:1,font:{{size:10}},color:'#94A3B8'}},grid:{{color:'#F1F5F9'}}}},
        y:{{ticks:{{font:{{size:11,weight:'600'}},color:'#1E293B'}},grid:{{display:false}}}},
      }},
    }},
  }});
}}

function initCharts() {{
  const tc = {{ A:'{TIER_COLORS['A']}', B:'{TIER_COLORS['B']}', C:'{TIER_COLORS['C']}', D:'{TIER_COLORS['D']}' }};
  for (const [t,c] of Object.entries(tc)) {{
    const d = POP.tiers?.[t]||{{}};
    makeBar('c'+t, Object.keys(d).slice(0,12), Object.values(d).slice(0,12), c);
  }}
  const wd = POP.winner||{{}};
  makeBar('cW', Object.keys(wd).slice(0,10), Object.values(wd).slice(0,10), '#16A34A');
  const sd = POP.scorers||{{}};
  const trunc = s => s.length > 18 ? s.slice(0,17)+'…' : s;
  makeBar('cS', Object.keys(sd).slice(0,18).map(trunc), Object.values(sd).slice(0,18), '#EA580C');
}}
</script>
</body>
</html>"""


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    print("⚽  World Cup Pool 2026 — Standings Tracker")
    print("=" * 48)

    no_fetch   = "--no-fetch" in sys.argv
    pos_args   = [a for a in sys.argv[1:] if not a.startswith("--")]
    excel_path = pos_args[0] if pos_args else EXCEL_FILE

    print(f"\n📥  Reading entries from: {excel_path}")
    if not os.path.exists(excel_path):
        print(f"  ❌  File not found: {excel_path}")
        sys.exit(1)
    entries = read_entries(excel_path)
    print(f"  ✓  {len(entries)} participants loaded")

    fetch_succeeded = False
    if no_fetch:
        print("\n📋  Skipping ESPN fetch (--no-fetch)")
        team_results, advanced, champion, knocked_out = {}, set(), set(), set()
        scorer_goals, round_matches = {}, {}
        group_standings = {}
        fetch_attempted = False
    else:
        print("\n🌐  Fetching live results from ESPN…")
        team_results, advanced, champion, knocked_out, scorer_goals, round_matches, fetch_succeeded = fetch_results()
        fetch_attempted = True
        print("\n📋  Fetching group standings from ESPN…")
        group_standings = fetch_group_standings()
        gs_count = len(group_standings)
        print(f"  ✓  {gs_count} group{'s' if gs_count != 1 else ''} loaded"
              if gs_count else "  ℹ  No group standings yet (pre-tournament)")

    pop = pick_popularity(entries)

    print("\n📊  Calculating standings…")
    standings = calculate_scores(entries, team_results, advanced, champion,
                                 knocked_out, scorer_goals)
    if standings:
        print(f"  Leader: {standings[0]['name']}  —  {standings[0]['total_points']:,} pts")

    print(f"\n✍️   Generating {OUTPUT_FILE}…")
    html = generate_html(
        standings, team_results, advanced, champion, knocked_out,
        entries, pop, scorer_goals, group_standings, round_matches,
        fetch_attempted=fetch_attempted,
        fetch_succeeded=fetch_succeeded,
    )
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    if scorer_goals:
        top = sorted(scorer_goals.items(), key=lambda x: -x[1])[:5]
        print(f"  Top scorers: {', '.join(f'{n} ({g})' for n, g in top)}")

    print(f"\n✅  Done! Open {OUTPUT_FILE} in your browser.")
    print("    Push to GitHub Pages to share with participants.")
    print("    GitHub Action re-runs automatically every 2 hours.\n")
