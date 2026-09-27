import os
import json
import time
import random
from datetime import date, datetime, timedelta
from pathlib import Path

import requests
import pandas as pd


# ============================================================
# FOTMOB FOOTBALL DATA SCRAPER
# ============================================================
#
# Collects:
#   - Top 5 European leagues
#   - Matches
#   - Match statistics
#   - Players
#   - Player match performance
#   - Teams
#   - Team information
#   - Shots / xG when available
#   - League standings
#
# Output:
#   CSV files
#   Excel workbook
#   SQLite database
#   Raw JSON cache
#
# Designed for Power BI / Star Schema
# ============================================================


# ============================================================
# 1. SETTINGS
# ============================================================

SEASON_NAME = "2026/2027"

# Current season start.
# Change these if you want another season.
START_DATE = date(2026, 8, 1)
END_DATE = date(2027, 6, 1)

# Top 5 leagues
LEAGUES = {
    47: {
        "name": "Premier League",
        "country": "England",
        "ccode3": "ENG"
    },

    87: {
        "name": "LaLiga",
        "country": "Spain",
        "ccode3": "ESP"
    },

    55: {
        "name": "Serie A",
        "country": "Italy",
        "ccode3": "ITA"
    },

    54: {
        "name": "Bundesliga",
        "country": "Germany",
        "ccode3": "GER"
    },

    53: {
        "name": "Ligue 1",
        "country": "France",
        "ccode3": "FRA"
    }
}


# ============================================================
# 2. FOLDERS
# ============================================================

OUTPUT_DIR = Path("fotmob_powerbi")

RAW_DIR = OUTPUT_DIR / "raw"

MATCH_CACHE_DIR = RAW_DIR / "matches"
TEAM_CACHE_DIR = RAW_DIR / "teams"
PLAYER_CACHE_DIR = RAW_DIR / "players"
LEAGUE_CACHE_DIR = RAW_DIR / "leagues"

ERROR_DIR = OUTPUT_DIR / "errors"

for folder in [
    OUTPUT_DIR,
    RAW_DIR,
    MATCH_CACHE_DIR,
    TEAM_CACHE_DIR,
    PLAYER_CACHE_DIR,
    LEAGUE_CACHE_DIR,
    ERROR_DIR
]:
    folder.mkdir(parents=True, exist_ok=True)


# ============================================================
# 3. API SETTINGS
# ============================================================

# Current community references document /api/data/ routes.
# Some older/current references use /api/ routes.
# We therefore try both.

BASE_URLS = [
    "https://www.fotmob.com/api/data",
    "https://www.fotmob.com/api"
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/153.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json,text/plain,*/*",
    "Referer": "https://www.fotmob.com/"
}


# ============================================================
# 4. REQUEST SETTINGS
# ============================================================

MAX_RETRIES = 5

# Keep requests slow enough to avoid hammering the service.
MIN_DELAY = 0.8
MAX_DELAY = 1.8

REQUEST_TIMEOUT = 30

session = requests.Session()
session.headers.update(HEADERS)


# ============================================================
# 5. GENERAL HELPERS
# ============================================================

def safe_int(value):
    """Convert value to int when possible."""
    try:
        if value is None:
            return None

        if isinstance(value, bool):
            return int(value)

        return int(float(value))

    except Exception:
        return None


def safe_float(value):
    """Convert value to float when possible."""
    try:
        if value is None:
            return None

        if isinstance(value, str):
            value = value.replace(",", "").strip()

        return float(value)

    except Exception:
        return None


def safe_string(value):
    """Convert value to clean string."""
    if value is None:
        return None

    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)

    return str(value)


def first_value(*values):
    """Return the first non-empty value."""
    for value in values:
        if value is not None and value != "":
            return value

    return None


def sleep_between_requests():
    time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))


def save_json(path, data):
    """Save JSON safely."""
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=2
            )

        return True

    except Exception as e:
        print(f"Could not save JSON: {e}")
        return False


def load_json(path):
    """Load cached JSON."""
    try:
        if not path.exists():
            return None

        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    except Exception:
        return None


# ============================================================
# 6. API REQUEST FUNCTION
# ============================================================

def api_get(endpoint, params=None, cache_file=None):
    """
    Fetch JSON from FotMob.

    Tries:
        /api/data/...
        /api/...

    Includes:
        retries
        rate limiting
        caching
        error logging
    """

    # --------------------------------------------------------
    # CACHE
    # --------------------------------------------------------

    if cache_file is not None:

        cached = load_json(cache_file)

        if cached is not None:
            return cached


    last_error = None


    # --------------------------------------------------------
    # TRY BOTH API BASE URLS
    # --------------------------------------------------------

    for base_url in BASE_URLS:

        url = f"{base_url}/{endpoint}"

        for attempt in range(1, MAX_RETRIES + 1):

            try:

                sleep_between_requests()

                response = session.get(
                    url,
                    params=params,
                    timeout=REQUEST_TIMEOUT
                )

                # ------------------------------------------------
                # RATE LIMIT
                # ------------------------------------------------

                if response.status_code == 429:

                    wait_time = min(
                        10 * attempt,
                        60
                    )

                    print(
                        f"  Rate limited. "
                        f"Waiting {wait_time}s..."
                    )

                    time.sleep(wait_time)

                    continue


                # ------------------------------------------------
                # SERVER ERRORS
                # ------------------------------------------------

                if response.status_code >= 500:

                    print(
                        f"  Server error "
                        f"{response.status_code}. "
                        f"Retry {attempt}/{MAX_RETRIES}"
                    )

                    time.sleep(3 * attempt)

                    continue


                # ------------------------------------------------
                # NOT FOUND
                # ------------------------------------------------

                if response.status_code == 404:

                    last_error = (
                        f"404 Not Found: {url}"
                    )

                    break


                # ------------------------------------------------
                # OTHER HTTP ERRORS
                # ------------------------------------------------

                response.raise_for_status()


                # ------------------------------------------------
                # JSON
                # ------------------------------------------------

                try:

                    data = response.json()

                except Exception:

                    text_preview = response.text[:500]

                    last_error = (
                        "Response was not JSON.\n"
                        f"URL: {url}\n"
                        f"Response: {text_preview}"
                    )

                    break


                # ------------------------------------------------
                # SAVE CACHE
                # ------------------------------------------------

                if cache_file is not None:
                    save_json(cache_file, data)


                return data


            except requests.RequestException as e:

                last_error = str(e)

                print(
                    f"  Request error: {e}"
                )

                if attempt < MAX_RETRIES:
                    time.sleep(2 * attempt)


    # --------------------------------------------------------
    # SAVE ERROR
    # --------------------------------------------------------

    error_file = ERROR_DIR / (
        f"error_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.txt"
    )

    try:

        with open(
            error_file,
            "w",
            encoding="utf-8"
        ) as f:

            f.write(
                f"Endpoint: {endpoint}\n"
                f"Params: {params}\n"
                f"Error: {last_error}\n"
            )

    except Exception:
        pass


    return None


# ============================================================
# 7. DATE HELPERS
# ============================================================

def daterange(start_date, end_date):

    current = start_date

    while current <= end_date:

        yield current

        current += timedelta(days=1)


# ============================================================
# 8. MATCH DISCOVERY
# ============================================================

def get_matches_for_date(target_date):

    date_string = target_date.strftime("%Y%m%d")

    cache_file = (
        RAW_DIR
        / "daily_matches"
        / f"{date_string}.json"
    )

    cache_file.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    data = api_get(
        "matches",
        params={
            "date": date_string
        },
        cache_file=cache_file
    )

    if data is None:
        return []


    # --------------------------------------------------------
    # FotMob normally returns:
    #
    # {
    #   "leagues": [...]
    # }
    #
    # But some wrappers/responses may expose a list.
    # --------------------------------------------------------

    if isinstance(data, dict):

        leagues = data.get("leagues", [])

    elif isinstance(data, list):

        leagues = data

    else:

        return []


    if not isinstance(leagues, list):
        return []


    results = []


    for league in leagues:

        if not isinstance(league, dict):
            continue


        league_id = first_value(
            league.get("primaryId"),
            league.get("leagueId"),
            league.get("id")
        )

        league_id = safe_int(league_id)


        # Only our five leagues
        if league_id not in LEAGUES:
            continue


        matches = league.get("matches", [])


        if not isinstance(matches, list):
            continue


        for match in matches:

            if not isinstance(match, dict):
                continue


            match_id = first_value(
                match.get("id"),
                match.get("matchId")
            )

            match_id = safe_int(match_id)


            if match_id is None:
                continue


            home = match.get("home", {})

            away = match.get("away", {})

            if not isinstance(home, dict):
                home = {}

            if not isinstance(away, dict):
                away = {}


            status = match.get(
                "status",
                {}
            )

            if not isinstance(status, dict):
                status = {}


            results.append({

                "Match_ID": match_id,

                "League_ID": league_id,

                "League_Name": LEAGUES[league_id]["name"],

                "Match_Date": target_date.isoformat(),

                "Home_Team_ID": safe_int(
                    home.get("id")
                ),

                "Home_Team_Name": home.get(
                    "name"
                ),

                "Away_Team_ID": safe_int(
                    away.get("id")
                ),

                "Away_Team_Name": away.get(
                    "name"
                ),

                "Home_Score": safe_int(
                    home.get("score")
                ),

                "Away_Score": safe_int(
                    away.get("score")
                ),

                "Status": safe_string(
                    status.get("reason")
                ),

                "Finished": status.get(
                    "finished"
                ),

                "Started": status.get(
                    "started"
                ),

                "Cancelled": status.get(
                    "cancelled"
                ),

                "UTC_Time": status.get(
                    "utcTime"
                ),

                "Score_String": status.get(
                    "scoreStr"
                )

            })


    return results


# ============================================================
# 9. GET ALL SEASON MATCHES
# ============================================================

def collect_all_matches():

    print("\n" + "=" * 70)
    print("STEP 1 - DISCOVERING MATCHES")
    print("=" * 70)

    all_matches = {}

    total_days = (
        END_DATE - START_DATE
    ).days + 1

    day_number = 0


    for current_date in daterange(
        START_DATE,
        END_DATE
    ):

        day_number += 1

        print(
            f"[{day_number}/{total_days}] "
            f"{current_date}"
        )


        matches = get_matches_for_date(
            current_date
        )


        for match in matches:

            match_id = match["Match_ID"]

            all_matches[match_id] = match


        print(
            f"    Found {len(matches)} top-5 matches | "
            f"Total unique: {len(all_matches)}"
        )


    df = pd.DataFrame(
        list(all_matches.values())
    )


    if not df.empty:

        df = df.sort_values(
            ["League_Name", "Match_Date", "Match_ID"]
        )


    print(
        f"\nTotal unique matches found: "
        f"{len(df)}"
    )


    return df


# ============================================================
# 10. MATCH DETAILS
# ============================================================

def get_match_details(match_id):

    cache_file = (
        MATCH_CACHE_DIR
        / f"{match_id}.json"
    )

    return api_get(
        "matchDetails",
        params={
            "matchId": match_id
        },
        cache_file=cache_file
    )


# ============================================================
# 11. GENERIC LIST EXTRACTION
# ============================================================

def find_lists(obj, key_name):

    """
    Recursively find every list associated with key_name.

    This makes the scraper much more resistant
    to changes in FotMob JSON nesting.
    """

    results = []


    if isinstance(obj, dict):

        for key, value in obj.items():

            if key.lower() == key_name.lower():

                if isinstance(value, list):
                    results.append(value)


            results.extend(
                find_lists(
                    value,
                    key_name
                )
            )


    elif isinstance(obj, list):

        for item in obj:

            results.extend(
                find_lists(
                    item,
                    key_name
                )
            )


    return results


# ============================================================
# 12. FIND PLAYER OBJECTS
# ============================================================

def extract_player_objects(obj):

    players = []


    def walk(value):

        if isinstance(value, dict):

            # ------------------------------------------------
            # Player-like object
            # ------------------------------------------------

            player_id = first_value(
                value.get("id"),
                value.get("playerId")
            )


            name = first_value(
                value.get("name"),
                value.get("playerName"),
                value.get("fullName")
            )


            if (
                player_id is not None
                and name is not None
            ):

                # Avoid treating teams as players
                position = first_value(
                    value.get("position"),
                    value.get("positionName")
                )


                # A player usually has one of these
                player_indicators = [
                    "rating",
                    "minutesPlayed",
                    "minutes_played",
                    "goals",
                    "assists",
                    "substitute",
                    "starter",
                    "position"
                ]


                if any(
                    key in value
                    for key in player_indicators
                ):

                    players.append(value)


            for child in value.values():
                walk(child)


        elif isinstance(value, list):

            for item in value:
                walk(item)


    walk(obj)


    # Remove duplicates
    unique = {}

    for player in players:

        pid = first_value(
            player.get("id"),
            player.get("playerId")
        )

        if pid is not None:

            unique[str(pid)] = player


    return list(unique.values())


# ============================================================
# 13. DIM MATCH
# ============================================================

def build_dim_match(
    matches_df,
    match_details
):

    rows = []


    for base_match in matches_df.to_dict(
        orient="records"
    ):

        match_id = base_match["Match_ID"]

        details = match_details.get(
            match_id
        )


        if not isinstance(details, dict):

            rows.append(base_match)

            continue


        general = details.get(
            "general",
            {}
        )

        header = details.get(
            "header",
            {}
        )


        if not isinstance(general, dict):
            general = {}

        if not isinstance(header, dict):
            header = {}


        home_team = general.get(
            "homeTeam",
            {}
        )

        away_team = general.get(
            "awayTeam",
            {}
        )


        if not isinstance(home_team, dict):
            home_team = {}

        if not isinstance(away_team, dict):
            away_team = {}


        row = dict(base_match)


        row["Home_Team_ID"] = first_value(
            safe_int(home_team.get("id")),
            row.get("Home_Team_ID")
        )


        row["Home_Team_Name"] = first_value(
            home_team.get("name"),
            row.get("Home_Team_Name")
        )


        row["Away_Team_ID"] = first_value(
            safe_int(away_team.get("id")),
            row.get("Away_Team_ID")
        )


        row["Away_Team_Name"] = first_value(
            away_team.get("name"),
            row.get("Away_Team_Name")
        )


        row["Referee"] = general.get(
            "referee"
        )


        row["Venue"] = general.get(
            "venue"
        )


        rows.append(row)


    return pd.DataFrame(rows)


# ============================================================
# 14. EXTRACT PLAYER MATCH DATA
# ============================================================

def build_fact_player_match(
    matches_df,
    match_details
):

    rows = []


    for match_id, details in match_details.items():

        if not isinstance(details, dict):
            continue


        base_match = matches_df[
            matches_df["Match_ID"] == match_id
        ]


        if base_match.empty:
            continue


        base = base_match.iloc[0].to_dict()


        # ----------------------------------------------------
        # Locate lineup
        # ----------------------------------------------------

        content = details.get(
            "content",
            {}
        )


        if not isinstance(content, dict):

            content = details


        lineup = content.get(
            "lineup",
            {}
        )


        if not isinstance(lineup, dict):
            continue


        # ----------------------------------------------------
        # FotMob lineup commonly contains:
        #
        # home / away
        # or homeTeam / awayTeam
        # ----------------------------------------------------

        sides = [
            ("home", lineup.get("home")),
            ("away", lineup.get("away")),
            ("homeTeam", lineup.get("homeTeam")),
            ("awayTeam", lineup.get("awayTeam"))
        ]


        for side_name, side_data in sides:

            if not isinstance(
                side_data,
                dict
            ):
                continue


            # Determine team
            team_id = first_value(
                side_data.get("teamId"),
                side_data.get("id")
            )

            team_name = first_value(
                side_data.get("teamName"),
                side_data.get("name")
            )


            team_id = safe_int(team_id)


            # ------------------------------------------------
            # Find players
            # ------------------------------------------------

            player_lists = []


            for key in [
                "players",
                "starters",
                "bench",
                "benchArr"
            ]:

                value = side_data.get(key)

                if isinstance(value, list):
                    player_lists.append(value)


                elif isinstance(value, dict):

                    for nested in value.values():

                        if isinstance(
                            nested,
                            list
                        ):

                            player_lists.append(
                                nested
                            )


            # If no obvious list, search recursively
            if not player_lists:

                player_lists = find_lists(
                    side_data,
                    "players"
                )


            for player_list in player_lists:

                for player in player_list:

                    if not isinstance(
                        player,
                        dict
                    ):
                        continue


                    player_id = first_value(
                        player.get("id"),
                        player.get("playerId")
                    )


                    if player_id is None:
                        continue


                    player_id = safe_int(
                        player_id
                    )


                    player_name = first_value(
                        player.get("name"),
                        player.get("playerName")
                    )


                    # ------------------------------------------------
                    # Keep all useful fields without assuming
                    # every field exists.
                    # ------------------------------------------------

                    row = {

                        "Match_ID": match_id,

                        "League_ID": base.get(
                            "League_ID"
                        ),

                        "Match_Date": base.get(
                            "Match_Date"
                        ),

                        "Team_ID": team_id,

                        "Team_Name": team_name,

                        "Opponent_Team_ID":
                            base.get(
                                "Away_Team_ID"
                            )
                            if side_name.lower().startswith("home")
                            else base.get(
                                "Home_Team_ID"
                            ),

                        "Player_ID": player_id,

                        "Player_Name": player_name,

                        "Position": first_value(
                            player.get("position"),
                            player.get("positionName")
                        ),

                        "Starter": first_value(
                            player.get("starter"),
                            player.get("isStarter")
                        ),

                        "Substitute": first_value(
                            player.get("substitute"),
                            player.get("isSubstitute")
                        ),

                        "Minutes_Played": first_value(
                            player.get("minutesPlayed"),
                            player.get("minutes_played")
                        ),

                        "Rating": first_value(
                            player.get("rating"),
                            player.get("ratingNum")
                        ),

                        "Goals": player.get(
                            "goals"
                        ),

                        "Assists": player.get(
                            "assists"
                        ),

                        "Expected_Goals": first_value(
                            player.get("expectedGoals"),
                            player.get("xG")
                        ),

                        "Expected_Assists": first_value(
                            player.get("expectedAssists"),
                            player.get("xA")
                        ),

                        "Shots": player.get(
                            "shots"
                        ),

                        "Shots_On_Target": first_value(
                            player.get("shotsOnTarget"),
                            player.get("shots_on_target")
                        ),

                        "Key_Passes": first_value(
                            player.get("keyPasses"),
                            player.get("key_passes")
                        ),

                        "Passes": player.get(
                            "passes"
                        ),

                        "Tackles": player.get(
                            "tackles"
                        ),

                        "Interceptions": player.get(
                            "interceptions"
                        ),

                        "Clearances": player.get(
                            "clearances"
                        ),

                        "Duels_Won": first_value(
                            player.get("duelsWon"),
                            player.get("duels_won")
                        ),

                        "Touches": player.get(
                            "touches"
                        ),

                        "Yellow_Cards": first_value(
                            player.get("yellowCards"),
                            player.get("yellow_cards")
                        ),

                        "Red_Cards": first_value(
                            player.get("redCards"),
                            player.get("red_cards")
                        )

                    }


                    # Keep extra stats as JSON
                    # instead of losing information.
                    row["Raw_Player_Data"] = json.dumps(
                        player,
                        ensure_ascii=False
                    )


                    rows.append(row)


    return pd.DataFrame(rows)


# ============================================================
# 15. EXTRACT MATCH STATISTICS
# ============================================================

def build_fact_match_statistics(
    matches_df,
    match_details
):

    rows = []


    for match_id, details in match_details.items():

        if not isinstance(details, dict):
            continue


        base_match = matches_df[
            matches_df["Match_ID"] == match_id
        ]


        if base_match.empty:
            continue


        base = base_match.iloc[0].to_dict()


        content = details.get(
            "content",
            {}
        )


        if not isinstance(content, dict):
            continue


        stats = content.get(
            "stats",
            {}
        )


        if not isinstance(stats, dict):
            continue


        periods = stats.get(
            "Periods",
            {}
        )


        if not isinstance(periods, dict):
            continue


        # ----------------------------------------------------
        # Usually "All" contains full-match statistics.
        # ----------------------------------------------------

        all_stats = periods.get(
            "All"
        )


        if isinstance(all_stats, dict):

            stat_groups = all_stats.get(
                "stats",
                []
            )

        elif isinstance(all_stats, list):

            stat_groups = all_stats

        else:

            stat_groups = []


        if not isinstance(
            stat_groups,
            list
        ):
            continue


        for group in stat_groups:

            if not isinstance(
                group,
                dict
            ):
                continue


            group_title = group.get(
                "title"
            )


            group_stats = group.get(
                "stats",
                []
            )


            if not isinstance(
                group_stats,
                list
            ):
                continue


            for stat in group_stats:

                if not isinstance(
                    stat,
                    dict
                ):
                    continue


                row = {

                    "Match_ID": match_id,

                    "League_ID": base.get(
                        "League_ID"
                    ),

                    "Match_Date": base.get(
                        "Match_Date"
                    ),

                    "Stat_Group": group_title,

                    "Stat_Title": first_value(
                        stat.get("title"),
                        stat.get("key")
                    ),

                    "Home_Value": first_value(
                        stat.get("home"),
                        stat.get("homeValue")
                    ),

                    "Away_Value": first_value(
                        stat.get("away"),
                        stat.get("awayValue")
                    ),

                    "Home_Display": stat.get(
                        "homeValue"
                    ),

                    "Away_Display": stat.get(
                        "awayValue"
                    )

                }


                rows.append(row)


    return pd.DataFrame(rows)


# ============================================================
# 16. EXTRACT SHOTS
# ============================================================

def build_fact_shots(
    matches_df,
    match_details
):

    rows = []


    for match_id, details in match_details.items():

        if not isinstance(details, dict):
            continue


        base_match = matches_df[
            matches_df["Match_ID"] == match_id
        ]


        if base_match.empty:
            continue


        base = base_match.iloc[0].to_dict()


        content = details.get(
            "content",
            {}
        )


        if not isinstance(content, dict):
            continue


        shotmap = content.get(
            "shotmap",
            {}
        )


        if not isinstance(
            shotmap,
            dict
        ):
            continue


        shots = shotmap.get(
            "shots",
            []
        )


        if not isinstance(
            shots,
            list
        ):
            continue


        for shot in shots:

            if not isinstance(
                shot,
                dict
            ):
                continue


            row = {

                "Match_ID": match_id,

                "League_ID": base.get(
                    "League_ID"
                ),

                "Match_Date": base.get(
                    "Match_Date"
                ),

                "Player_ID": safe_int(
                    first_value(
                        shot.get("playerId"),
                        shot.get("player", {}).get("id")
                        if isinstance(
                            shot.get("player"),
                            dict
                        )
                        else None
                    )
                ),

                "Player_Name": first_value(
                    shot.get("playerName"),
                    shot.get("player", {}).get("name")
                    if isinstance(
                        shot.get("player"),
                        dict
                    )
                    else None
                ),

                "Team_ID": safe_int(
                    first_value(
                        shot.get("teamId"),
                        shot.get("team", {}).get("id")
                        if isinstance(
                            shot.get("team"),
                            dict
                        )
                        else None
                    )
                ),

                "X": shot.get("x"),

                "Y": shot.get("y"),

                "Min": shot.get("min"),

                "Sec": shot.get("sec"),

                "Expected_Goal": first_value(
                    shot.get("expectedGoals"),
                    shot.get("xg")
                ),

                "Shot_Type": first_value(
                    shot.get("shotType"),
                    shot.get("type")
                ),

                "Outcome": first_value(
                    shot.get("eventType"),
                    shot.get("result")
                ),

                "Period": shot.get(
                    "period"
                ),

                "On_Target": shot.get(
                    "onTarget"
                ),

                "Blocked": shot.get(
                    "blocked"
                )

            }


            row["Raw_Shot_Data"] = json.dumps(
                shot,
                ensure_ascii=False
            )


            rows.append(row)


    return pd.DataFrame(rows)


# ============================================================
# 17. EXTRACT EVENTS
# ============================================================

def build_fact_events(
    matches_df,
    match_details
):

    rows = []


    for match_id, details in match_details.items():

        if not isinstance(details, dict):
            continue


        base_match = matches_df[
            matches_df["Match_ID"] == match_id
        ]


        if base_match.empty:
            continue


        base = base_match.iloc[0].to_dict()


        content = details.get(
            "content",
            {}
        )


        if not isinstance(content, dict):
            continue


        # Try several possible event locations

        possible_events = []


        match_facts = content.get(
            "matchFacts"
        )


        if isinstance(
            match_facts,
            dict
        ):

            events = match_facts.get(
                "events"
            )

            if isinstance(
                events,
                list
            ):

                possible_events.extend(
                    events
                )


        events_obj = content.get(
            "events"
        )


        if isinstance(
            events_obj,
            dict
        ):

            for key in [
                "events",
                "chronological"
            ]:

                events = events_obj.get(
                    key
                )

                if isinstance(
                    events,
                    list
                ):

                    possible_events.extend(
                        events
                    )


        # Remove duplicate objects
        seen = set()


        for event in possible_events:

            if not isinstance(
                event,
                dict
            ):
                continue


            raw = json.dumps(
                event,
                sort_keys=True,
                ensure_ascii=False
            )


            if raw in seen:
                continue


            seen.add(raw)


            rows.append({

                "Match_ID": match_id,

                "League_ID": base.get(
                    "League_ID"
                ),

                "Match_Date": base.get(
                    "Match_Date"
                ),

                "Event_Type": first_value(
                    event.get("type"),
                    event.get("eventType")
                ),

                "Time": first_value(
                    event.get("time"),
                    event.get("minute")
                ),

                "Player_ID": safe_int(
                    first_value(
                        event.get("playerId"),
                        event.get("player", {}).get("id")
                        if isinstance(
                            event.get("player"),
                            dict
                        )
                        else None
                    )
                ),

                "Player_Name": first_value(
                    event.get("playerName"),
                    event.get("player", {}).get("name")
                    if isinstance(
                        event.get("player"),
                        dict
                    )
                    else None
                ),

                "Team_ID": safe_int(
                    first_value(
                        event.get("teamId"),
                        event.get("team", {}).get("id")
                        if isinstance(
                            event.get("team"),
                            dict
                        )
                        else None
                    )
                ),

                "Description": first_value(
                    event.get("description"),
                    event.get("text")
                ),

                "Raw_Event_Data": raw

            })


    return pd.DataFrame(rows)


# ============================================================
# 18. COLLECT TEAM IDS
# ============================================================

def collect_team_ids(
    matches_df
):

    teams = {}


    for _, row in matches_df.iterrows():

        home_id = safe_int(
            row.get("Home_Team_ID")
        )

        away_id = safe_int(
            row.get("Away_Team_ID")
        )


        if home_id is not None:

            teams[home_id] = {
                "Team_ID": home_id,
                "Team_Name": row.get(
                    "Home_Team_Name"
                ),
                "League_ID": row.get(
                    "League_ID"
                ),
                "League_Name": row.get(
                    "League_Name"
                )
            }


        if away_id is not None:

            teams[away_id] = {
                "Team_ID": away_id,
                "Team_Name": row.get(
                    "Away_Team_Name"
                ),
                "League_ID": row.get(
                    "League_ID"
                ),
                "League_Name": row.get(
                    "League_Name"
                )
            }


    return teams


# ============================================================
# 19. TEAM DATA
# ============================================================

def get_team_data(team_id):

    cache_file = (
        TEAM_CACHE_DIR
        / f"{team_id}.json"
    )

    return api_get(
        "teams",
        params={
            "id": team_id
        },
        cache_file=cache_file
    )


# ============================================================
# 20. BUILD DIM TEAM
# ============================================================

def build_dim_team(
    team_index
):

    rows = []


    for team_id, basic in team_index.items():

        data = get_team_data(
            team_id
        )


        row = dict(basic)


        if isinstance(
            data,
            dict
        ):

            details = data.get(
                "details",
                {}
            )


            if not isinstance(
                details,
                dict
            ):
                details = {}


            row["Team_Name"] = first_value(
                details.get("name"),
                row.get("Team_Name")
            )


            row["Short_Name"] = first_value(
                details.get("shortName"),
                details.get("shortName2")
            )


            row["Country"] = first_value(
                details.get("country")
            )


            row["Stadium"] = first_value(
                details.get("stadium"),
                details.get("venue")
            )


            row["Manager"] = first_value(
                details.get("manager")
            )


            row["Team_Raw_JSON"] = json.dumps(
                data,
                ensure_ascii=False
            )


        rows.append(row)


    return pd.DataFrame(rows)


# ============================================================
# 21. COLLECT PLAYER IDS
# ============================================================

def collect_player_ids(
    player_match_df
):

    players = {}


    if player_match_df.empty:
        return players


    for _, row in player_match_df.iterrows():

        player_id = safe_int(
            row.get("Player_ID")
        )


        if player_id is None:
            continue


        if player_id not in players:

            players[player_id] = {

                "Player_ID": player_id,

                "Player_Name": row.get(
                    "Player_Name"
                ),

                "Team_ID": row.get(
                    "Team_ID"
                ),

                "Team_Name": row.get(
                    "Team_Name"
                )

            }


    return players


# ============================================================
# 22. PLAYER DATA
# ============================================================

def get_player_data(
    player_id
):

    cache_file = (
        PLAYER_CACHE_DIR
        / f"{player_id}.json"
    )

    return api_get(
        "playerData",
        params={
            "id": player_id
        },
        cache_file=cache_file
    )


# ============================================================
# 23. BUILD DIM PLAYER
# ============================================================

def build_dim_player(
    player_index
):

    rows = []


    for player_id, basic in player_index.items():

        data = get_player_data(
            player_id
        )


        row = dict(basic)


        if isinstance(
            data,
            dict
        ):

            # FotMob may expose player details
            # at different levels.

            details = data.get(
                "player",
                {}
            )


            if not isinstance(
                details,
                dict
            ):
                details = data


            row["Player_Name"] = first_value(
                details.get("name"),
                details.get("fullName"),
                row.get("Player_Name")
            )


            row["Position"] = first_value(
                details.get("position"),
                details.get("positionName")
            )


            row["Birth_Date"] = first_value(
                details.get("birthDate"),
                details.get("dateOfBirth")
            )


            row["Nationality"] = first_value(
                details.get("nationality"),
                details.get("country")
            )


            row["Height"] = details.get(
                "height"
            )


            row["Preferred_Foot"] = first_value(
                details.get("preferredFoot"),
                details.get("foot")
            )


            row["Player_Raw_JSON"] = json.dumps(
                data,
                ensure_ascii=False
            )


        rows.append(row)


    return pd.DataFrame(rows)


# ============================================================
# 24. PLAYER SEASON AGGREGATION
# ============================================================

def build_fact_player_season_stats(
    player_match_df
):

    if player_match_df.empty:

        return pd.DataFrame()


    df = player_match_df.copy()


    numeric_columns = [

        "Minutes_Played",
        "Rating",
        "Goals",
        "Assists",
        "Expected_Goals",
        "Expected_Assists",
        "Shots",
        "Shots_On_Target",
        "Key_Passes",
        "Passes",
        "Tackles",
        "Interceptions",
        "Clearances",
        "Duels_Won",
        "Touches",
        "Yellow_Cards",
        "Red_Cards"

    ]


    for column in numeric_columns:

        if column in df.columns:

            df[column] = pd.to_numeric(
                df[column],
                errors="coerce"
            )


    # One player/team/league combination
    group_columns = [
        "League_ID",
        "Team_ID",
        "Player_ID"
    ]


    aggregation = {

        "Match_ID": "nunique",

        "Minutes_Played": "sum",

        "Rating": "mean",

        "Goals": "sum",

        "Assists": "sum",

        "Expected_Goals": "sum",

        "Expected_Assists": "sum",

        "Shots": "sum",

        "Shots_On_Target": "sum",

        "Key_Passes": "sum",

        "Passes": "sum",

        "Tackles": "sum",

        "Interceptions": "sum",

        "Clearances": "sum",

        "Duels_Won": "sum",

        "Touches": "sum",

        "Yellow_Cards": "sum",

        "Red_Cards": "sum"

    }


    existing_aggregation = {

        key: value

        for key, value in aggregation.items()

        if key in df.columns

    }


    result = (

        df.groupby(
            group_columns,
            dropna=False
        )

        .agg(existing_aggregation)

        .reset_index()

    )


    result.rename(
        columns={
            "Match_ID":
                "Matches_Played",

            "Rating":
                "Average_Rating"
        },
        inplace=True
    )


    result["Season"] = SEASON_NAME


    return result


# ============================================================
# 25. TEAM SEASON STATISTICS
# ============================================================

def build_fact_team_season_stats(
    matches_df
):

    if matches_df.empty:

        return pd.DataFrame()


    rows = []


    for league_id, league_group in matches_df.groupby(
        "League_ID"
    ):

        team_stats = {}


        for _, match in league_group.iterrows():

            home_id = safe_int(
                match.get(
                    "Home_Team_ID"
                )
            )

            away_id = safe_int(
                match.get(
                    "Away_Team_ID"
                )
            )


            home_name = match.get(
                "Home_Team_Name"
            )

            away_name = match.get(
                "Away_Team_Name"
            )


            home_score = safe_int(
                match.get(
                    "Home_Score"
                )
            )

            away_score = safe_int(
                match.get(
                    "Away_Score"
                )
            )


            if (
                home_id is None
                or away_id is None
            ):
                continue


            # Initialize
            for team_id, team_name in [
                (home_id, home_name),
                (away_id, away_name)
            ]:

                if team_id not in team_stats:

                    team_stats[team_id] = {

                        "League_ID": league_id,

                        "Team_ID": team_id,

                        "Team_Name": team_name,

                        "Matches_Played": 0,

                        "Wins": 0,

                        "Draws": 0,

                        "Losses": 0,

                        "Goals_For": 0,

                        "Goals_Against": 0,

                        "Points": 0

                    }


            # Only completed matches
            if (
                home_score is None
                or away_score is None
            ):
                continue


            team_stats[home_id][
                "Matches_Played"
            ] += 1


            team_stats[away_id][
                "Matches_Played"
            ] += 1


            team_stats[home_id][
                "Goals_For"
            ] += home_score


            team_stats[home_id][
                "Goals_Against"
            ] += away_score


            team_stats[away_id][
                "Goals_For"
            ] += away_score


            team_stats[away_id][
                "Goals_Against"
            ] += home_score


            # Home result
            if home_score > away_score:

                team_stats[home_id][
                    "Wins"
                ] += 1

                team_stats[home_id][
                    "Points"
                ] += 3

                team_stats[away_id][
                    "Losses"
                ] += 1


            elif home_score < away_score:

                team_stats[away_id][
                    "Wins"
                ] += 1

                team_stats[away_id][
                    "Points"
                ] += 3

                team_stats[home_id][
                    "Losses"
                ] += 1


            else:

                team_stats[home_id][
                    "Draws"
                ] += 1

                team_stats[away_id][
                    "Draws"
                ] += 1

                team_stats[home_id][
                    "Points"
                ] += 1

                team_stats[away_id][
                    "Points"
                ] += 1


        for team in team_stats.values():

            team["Goal_Difference"] = (
                team["Goals_For"]
                -
                team["Goals_Against"]
            )

            team["Season"] = SEASON_NAME

            rows.append(team)


    return pd.DataFrame(rows)


# ============================================================
# 26. LEAGUE STANDINGS
# ============================================================

def get_league_data(
    league_id
):

    league = LEAGUES[league_id]

    cache_file = (
        LEAGUE_CACHE_DIR
        / f"{league_id}_{SEASON_NAME.replace('/', '-')}.json"
    )


    return api_get(

        "leagues",

        params={
            "id": league_id,
            "season": SEASON_NAME,
            "ccode3": league["ccode3"]
        },

        cache_file=cache_file
    )


# ============================================================
# 27. EXTRACT STANDINGS ROBUSTLY
# ============================================================

def extract_standing_rows(
    data
):

    if not isinstance(
        data,
        dict
    ):
        return []


    candidates = []


    def walk(obj):

        if isinstance(
            obj,
            dict
        ):

            for key, value in obj.items():

                key_lower = str(
                    key
                ).lower()


                if key_lower in [
                    "table",
                    "rows",
                    "standings"
                ]:

                    if isinstance(
                        value,
                        list
                    ):

                        candidates.append(
                            value
                        )


                    elif isinstance(
                        value,
                        dict
                    ):

                        walk(value)


                else:

                    walk(value)


        elif isinstance(
            obj,
            list
        ):

            for item in obj:
                walk(item)


    walk(data)


    # Choose list containing team-like objects
    for candidate in candidates:

        if not candidate:
            continue


        valid_count = 0


        for row in candidate:

            if not isinstance(
                row,
                dict
            ):
                continue


            if any(
                key in row
                for key in [
                    "name",
                    "team",
                    "id",
                    "pts",
                    "points"
                ]
            ):

                valid_count += 1


        if valid_count > 0:

            return candidate


    return []


# ============================================================
# 28. BUILD FACT STANDINGS
# ============================================================

def build_fact_standings():

    rows = []


    for league_id in LEAGUES:

        print(
            f"Getting standings: "
            f"{LEAGUES[league_id]['name']}"
        )


        data = get_league_data(
            league_id
        )


        if data is None:
            continue


        standing_rows = extract_standing_rows(
            data
        )


        for position, row in enumerate(
            standing_rows,
            start=1
        ):

            if not isinstance(
                row,
                dict
            ):
                continue


            team = row.get(
                "team",
                {}
            )


            if not isinstance(
                team,
                dict
            ):
                team = {}


            rows.append({

                "League_ID": league_id,

                "League_Name":
                    LEAGUES[league_id]["name"],

                "Season": SEASON_NAME,

                "Position": first_value(
                    row.get("idx"),
                    row.get("position"),
                    position
                ),

                "Team_ID": safe_int(
                    first_value(
                        row.get("id"),
                        team.get("id")
                    )
                ),

                "Team_Name": first_value(
                    row.get("name"),
                    team.get("name")
                ),

                "Played": first_value(
                    row.get("played"),
                    row.get("matches")
                ),

                "Wins": row.get(
                    "wins"
                ),

                "Draws": row.get(
                    "draws"
                ),

                "Losses": row.get(
                    "losses"
                ),

                "Goals_For": first_value(
                    row.get("scoresFor"),
                    row.get("goalsFor")
                ),

                "Goals_Against": first_value(
                    row.get("scoresAgainst"),
                    row.get("goalsAgainst")
                ),

                "Goal_Difference": first_value(
                    row.get("goalDifference"),
                    row.get("gd")
                ),

                "Points": first_value(
                    row.get("pts"),
                    row.get("points")
                )

            })


    return pd.DataFrame(rows)


# ============================================================
# 29. LEAGUE DIMENSION
# ============================================================

def build_dim_league():

    rows = []


    for league_id, league in LEAGUES.items():

        rows.append({

            "League_ID": league_id,

            "League_Name":
                league["name"],

            "Country":
                league["country"],

            "Country_Code":
                league["ccode3"],

            "Season":
                SEASON_NAME

        })


    return pd.DataFrame(rows)


# ============================================================
# 30. EXPORT CSV
# ============================================================

def export_csv(
    datasets
):

    print("\n" + "=" * 70)
    print("EXPORTING CSV FILES")
    print("=" * 70)


    for name, df in datasets.items():

        if df is None:
            continue


        if not isinstance(
            df,
            pd.DataFrame
        ):
            continue


        path = (
            OUTPUT_DIR
            / f"{name}.csv"
        )


        try:

            df.to_csv(
                path,
                index=False,
                encoding="utf-8-sig"
            )


            print(
                f"  {name}: "
                f"{len(df):,} rows"
            )


        except Exception as e:

            print(
                f"  ERROR exporting {name}: "
                f"{e}"
            )


# ============================================================
# 31. EXPORT EXCEL
# ============================================================

def export_excel(
    datasets
):

    print("\n" + "=" * 70)
    print("CREATING EXCEL WORKBOOK")
    print("=" * 70)


    path = (
        OUTPUT_DIR
        / "FotMob_Football_Data.xlsx"
    )


    with pd.ExcelWriter(
        path,
        engine="openpyxl"
    ) as writer:

        for name, df in datasets.items():

            if df is None:
                continue


            if not isinstance(
                df,
                pd.DataFrame
            ):
                continue


            # Excel has a maximum of 1,048,576 rows.
            if len(df) > 1_000_000:

                print(
                    f"  {name} is too large "
                    f"for one Excel sheet."
                )

                # First million
                df.iloc[
                    :1_000_000
                ].to_excel(
                    writer,
                    sheet_name=name[:31],
                    index=False
                )

            else:

                df.to_excel(
                    writer,
                    sheet_name=name[:31],
                    index=False
                )


            print(
                f"  {name}: "
                f"{len(df):,} rows"
            )


    print(
        f"\nExcel created:\n"
        f"{path}"
    )


# ============================================================
# 32. EXPORT SQLITE
# ============================================================

def export_sqlite(
    datasets
):

    import sqlite3


    path = (
        OUTPUT_DIR
        / "FotMob_Football_Data.db"
    )


    connection = sqlite3.connect(
        path
    )


    for name, df in datasets.items():

        if df is None:
            continue


        if not isinstance(
            df,
            pd.DataFrame
        ):
            continue


        df.to_sql(
            name,
            connection,
            if_exists="replace",
            index=False
        )


    connection.close()


    print(
        f"\nSQLite database created:\n"
        f"{path}"
    )


# ============================================================
# 33. MAIN PROGRAM
# ============================================================

def main():

    start_time = time.time()


    print("\n")
    print("=" * 70)
    print("FOTMOB FOOTBALL DATA COLLECTOR")
    print("=" * 70)

    print(
        f"Season: {SEASON_NAME}"
    )

    print(
        f"Date range: "
        f"{START_DATE} → {END_DATE}"
    )

    print(
        "Leagues:"
    )

    for league in LEAGUES.values():

        print(
            f"  - {league['name']}"
        )


    # ========================================================
    # STEP 1
    # MATCH DISCOVERY
    # ========================================================

    matches_df = collect_all_matches()


    if matches_df.empty:

        print(
            "\nNo matches were found."
        )

        print(
            "Check the API response or date range."
        )

        return


    # Save immediately
    matches_df.to_csv(
        OUTPUT_DIR / "DimMatch.csv",
        index=False,
        encoding="utf-8-sig"
    )


    # ========================================================
    # STEP 2
    # MATCH DETAILS
    # ========================================================

    print("\n" + "=" * 70)
    print("STEP 2 - DOWNLOADING MATCH DETAILS")
    print("=" * 70)


    match_details = {}


    total_matches = len(matches_df)


    for index, match_id in enumerate(
        matches_df["Match_ID"].unique(),
        start=1
    ):

        print(
            f"[{index}/{total_matches}] "
            f"Match {match_id}"
        )


        details = get_match_details(
            match_id
        )


        if details is not None:

            match_details[
                int(match_id)
            ] = details


    print(
        f"\nDownloaded details for "
        f"{len(match_details):,} matches."
    )


    # ========================================================
    # STEP 3
    # BUILD MATCH TABLE
    # ========================================================

    dim_match = build_dim_match(
        matches_df,
        match_details
    )


    # ========================================================
    # STEP 4
    # PLAYER MATCH DATA
    # ========================================================

    print("\n" + "=" * 70)
    print("STEP 3 - PLAYER MATCH STATISTICS")
    print("=" * 70)


    fact_player_match = (
        build_fact_player_match(
            matches_df,
            match_details
        )
    )


    print(
        f"Player-match rows: "
        f"{len(fact_player_match):,}"
    )


    # ========================================================
    # STEP 5
    # MATCH STATISTICS
    # ========================================================

    print("\n" + "=" * 70)
    print("STEP 4 - MATCH STATISTICS")
    print("=" * 70)


    fact_match_statistics = (
        build_fact_match_statistics(
            matches_df,
            match_details
        )
    )


    print(
        f"Match-stat rows: "
        f"{len(fact_match_statistics):,}"
    )


    # ========================================================
    # STEP 6
    # SHOTS
    # ========================================================

    print("\n" + "=" * 70)
    print("STEP 5 - SHOT DATA")
    print("=" * 70)


    fact_shots = build_fact_shots(
        matches_df,
        match_details
    )


    print(
        f"Shot rows: "
        f"{len(fact_shots):,}"
    )


    # ========================================================
    # STEP 7
    # EVENTS
    # ========================================================

    print("\n" + "=" * 70)
    print("STEP 6 - MATCH EVENTS")
    print("=" * 70)


    fact_events = build_fact_events(
        matches_df,
        match_details
    )


    print(
        f"Event rows: "
        f"{len(fact_events):,}"
    )


    # ========================================================
    # STEP 8
    # TEAMS
    # ========================================================

    print("\n" + "=" * 70)
    print("STEP 7 - TEAM DATA")
    print("=" * 70)


    team_index = collect_team_ids(
        matches_df
    )


    print(
        f"Unique teams: "
        f"{len(team_index):,}"
    )


    dim_team = build_dim_team(
        team_index
    )


    # ========================================================
    # STEP 9
    # PLAYERS
    # ========================================================

    print("\n" + "=" * 70)
    print("STEP 8 - PLAYER DATA")
    print("=" * 70)


    player_index = collect_player_ids(
        fact_player_match
    )


    print(
        f"Unique players: "
        f"{len(player_index):,}"
    )


    dim_player = build_dim_player(
        player_index
    )


    # ========================================================
    # STEP 10
    # PLAYER SEASON STATS
    # ========================================================

    print("\n" + "=" * 70)
    print("STEP 9 - PLAYER SEASON STATISTICS")
    print("=" * 70)


    fact_player_season_stats = (
        build_fact_player_season_stats(
            fact_player_match
        )
    )


    # ========================================================
    # STEP 11
    # TEAM SEASON STATS
    # ========================================================

    print("\n" + "=" * 70)
    print("STEP 10 - TEAM SEASON STATISTICS")
    print("=" * 70)


    fact_team_season_stats = (
        build_fact_team_season_stats(
            matches_df
        )
    )


    # ========================================================
    # STEP 12
    # STANDINGS
    # ========================================================

    print("\n" + "=" * 70)
    print("STEP 11 - LEAGUE STANDINGS")
    print("=" * 70)


    fact_standings = build_fact_standings()


    # ========================================================
    # STEP 13
    # LEAGUE DIMENSION
    # ========================================================

    dim_league = build_dim_league()


    # ========================================================
    # STEP 14
    # DATASET COLLECTION
    # ========================================================

    datasets = {

        "DimLeague":
            dim_league,

        "DimTeam":
            dim_team,

        "DimPlayer":
            dim_player,

        "DimMatch":
            dim_match,

        "FactStandings":
            fact_standings,

        "FactTeamSeasonStats":
            fact_team_season_stats,

        "FactPlayerSeasonStats":
            fact_player_season_stats,

        "FactPlayerMatch":
            fact_player_match,

        "FactMatchStatistics":
            fact_match_statistics,

        "FactEvents":
            fact_events,

        "FactShots":
            fact_shots

    }


    # ========================================================
    # STEP 15
    # EXPORT
    # ========================================================

    export_csv(
        datasets
    )


    export_excel(
        datasets
    )


    export_sqlite(
        datasets
    )


    # ========================================================
    # FINAL SUMMARY
    # ========================================================

    elapsed = (
        time.time()
        - start_time
    )


    print("\n")
    print("=" * 70)
    print("SCRAPING COMPLETED")
    print("=" * 70)


    for name, df in datasets.items():

        if isinstance(
            df,
            pd.DataFrame
        ):

            print(
                f"{name:<30} "
                f"{len(df):>10,} rows"
            )


    print("\nOutput folder:")
    print(
        OUTPUT_DIR.resolve()
    )


    print(
        f"\nTime taken: "
        f"{elapsed / 60:.2f} minutes"
    )


    print("\nFiles created:")

    print(
        "  FotMob_Football_Data.xlsx"
    )

    print(
        "  FotMob_Football_Data.db"
    )

    print(
        "  DimLeague.csv"
    )

    print(
        "  DimTeam.csv"
    )

    print(
        "  DimPlayer.csv"
    )

    print(
        "  DimMatch.csv"
    )

    print(
        "  FactStandings.csv"
    )

    print(
        "  FactTeamSeasonStats.csv"
    )

    print(
        "  FactPlayerSeasonStats.csv"
    )

    print(
        "  FactPlayerMatch.csv"
    )

    print(
        "  FactMatchStatistics.csv"
    )

    print(
        "  FactEvents.csv"
    )

    print(
        "  FactShots.csv"
    )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()