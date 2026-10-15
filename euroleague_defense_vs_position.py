"""
EUROLEAGUE - DEFENSE vs POSITION TABLE  (Google Colab + automatic web page)
=============================================================
What this script does
  1. Downloads every player box score of the chosen seasons (euroleague_api).
  2. Attaches a position (Guard / Forward / Center) to every player using your
     players_stats.xlsx file (matched by player ID, then by name).
  3. Optionally guesses the position of players that are NOT in your file.
  4. Builds the "defense vs position" table in terms of PIR only (the Euroleague
     efficiency rating): in each game, the average PIR of the opposing Guards /
     Forwards / Centers who played; then the average of those over all games.
     Lower = better defense against that position.
  5. Saves everything to an Excel workbook + CSVs and offers them for download.
  6. Writes a self-contained web page (index.html) with the same tables.

How to use in Colab
  - Paste the whole script into ONE cell and run it.
  - When asked, upload players_stats.xlsx.
  - Re-run it any time: only games played since the last run are downloaded.
  - Change the settings in section "CONFIG" below if you want other seasons.

Automatic updates (GitHub)
  - The same script runs unattended: the DVP_* environment variables below replace the
    Colab upload dialog. See README.md and .github/workflows/update.yml.
"""

# ============================================================================
# 0. SETUP  (installs missing packages; harmless if already installed)
# ============================================================================
import sys
import subprocess


def _ensure(import_name, pip_name=None):
    try:
        __import__(import_name)
    except ImportError:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-q", pip_name or import_name]
        )


_ensure("euroleague_api", "euroleague-api")
_ensure("openpyxl")
_ensure("sklearn", "scikit-learn")

import glob
import os
import re
import time
import unicodedata
from json.decoder import JSONDecodeError

import numpy as np
import pandas as pd
import requests
from tqdm.auto import tqdm
from euroleague_api.boxscore_data import BoxScoreData

# ============================================================================
# 1. CONFIG  - the only part you normally need to edit
# ============================================================================
COMPETITION = "E"                 # "E" = Euroleague, "U" = EuroCup
SEASONS = ([int(x) for x in os.environ["DVP_SEASONS"].split(",")]
           if os.environ.get("DVP_SEASONS") else [2026])   # 2026 = season 2026-27 (ongoing)

POSITIONS_FILE = os.environ.get("DVP_POSITIONS_FILE")   # None -> Colab asks you to upload it.
                                  # Or a path to the .xlsx / .csv, e.g. "/content/players_stats.xlsx"

# Players that are not in your file: guess their position from their stats?
INFER_MISSING_POSITIONS = True
MIN_MINUTES_TO_INFER = 20         # total minutes needed before we trust a guess

OUTPUT_DIR = os.environ.get("DVP_OUTPUT_DIR", "euroleague_output")
WRITE_HTML = True                 # also write OUTPUT_DIR/index.html (the web page)
FORCE_REFRESH = False             # True = ignore the saved file and re-download everything
PAUSE_SECONDS = 0.15              # small pause between API calls (be polite)

# Players with this many minutes or fewer are ignored. 0 = only drop players who did
# not play at all. Raise it (e.g. 3) to also ignore garbage-time cameos.
MIN_MINUTES_PLAYED = 0

POSITIONS = ["Guard", "Forward", "Center"]   # the three labels in your file

# Quick lookups printed at the end: (part of the team name, position).
# Example: average PIR of Forwards playing AGAINST Panathinaikos.
LOOKUPS = [("Panathinaikos", "Forward")]

# Optional manual fixes for team names, applied last. Example:
# TEAM_NAME_OVERRIDES = {"Panathinaikos AKTOR Athens": "Panathinaikos"}
TEAM_NAME_OVERRIDES = {}

# ============================================================================
# 2. HELPERS
# ============================================================================
CACHE_DIR = os.environ.get("DVP_CACHE_DIR", os.path.join(OUTPUT_DIR, "cache"))

# The API's column names -> our short names. Several spellings are listed so the
# script keeps working if the API renames a column slightly.
COLUMN_ALIASES = {
    "player_id": ["Player_ID", "PlayerID", "Player_Id"],
    "player": ["Player"],
    "team": ["Team"],
    "minutes": ["Minutes"],
    "pts": ["Points"],
    "fgm2": ["FieldGoalsMade2"],
    "fga2": ["FieldGoalsAttempted2"],
    "fgm3": ["FieldGoalsMade3"],
    "fga3": ["FieldGoalsAttempted3"],
    "ftm": ["FreeThrowsMade"],
    "fta": ["FreeThrowsAttempted"],
    "oreb": ["OffensiveRebounds"],
    "dreb": ["DefensiveRebounds"],
    "reb": ["TotalRebounds"],
    "ast": ["Assistances", "Assists"],
    "stl": ["Steals"],
    "tov": ["Turnovers"],
    "blk": ["BlocksFavour", "Blocks"],
    "pir": ["Valuation"],
}
REQUIRED = ["player_id", "player", "team", "minutes", "pts"]


def _pick(df, names):
    lower = {c.lower(): c for c in df.columns}
    for n in names:
        if n.lower() in lower:
            return lower[n.lower()]
    return None


def minutes_to_float(v):
    """'25:13' -> 25.22 ; 'DNP' / empty -> 0.0"""
    if pd.isna(v):
        return 0.0
    s = str(v).strip()
    if ":" in s:
        try:
            m, sec = s.split(":")[:2]
            return int(m) + int(sec) / 60
        except ValueError:
            return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def name_key(s):
    """Order-independent, accent-free name: 'VALANCIUNAS, JONAS' and
    'Jonas Valanciunas' both become 'jonas valanciunas'."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return " ".join(sorted(re.findall(r"[a-z]+", s.lower())))


def show(df, n=None):
    """Pretty print in Colab, plain print elsewhere."""
    try:
        from IPython.display import display
        display(df if n is None else df.head(n))
    except ImportError:
        print((df if n is None else df.head(n)).to_string())


# ============================================================================
# 3. LOAD POSITIONS FROM YOUR EXCEL FILE
# ============================================================================
POSITION_COLUMNS = {"ID", "Name", "Surname", "Position"}


def _is_positions_file(path):
    """True if the Excel file has the columns of the players_stats.xlsx file."""
    try:
        head = (pd.read_csv(path, nrows=1) if str(path).lower().endswith(".csv")
                else pd.read_excel(path, nrows=1))
        return POSITION_COLUMNS.issubset(head.columns)
    except Exception:
        return False


def get_positions_path():
    if POSITIONS_FILE:
        return POSITIONS_FILE

    # 1) already uploaded earlier in this Colab session?
    for path in sorted(glob.glob("*players_stats*.xlsx")):
        if _is_positions_file(path):
            print(f"Using positions file already in Colab: {path}")
            return path

    # 2) otherwise ask for it
    try:
        from google.colab import files
    except ImportError:
        raise SystemExit("Set POSITIONS_FILE to the path of players_stats.xlsx")
    print("Upload your players_stats.xlsx (the file with ID, Name, Surname, Position) ...")
    uploaded = files.upload()
    for path in uploaded:
        if _is_positions_file(path):
            return path
    raise SystemExit(
        f"None of the uploaded files ({list(uploaded)}) has the columns "
        f"{sorted(POSITION_COLUMNS)}. Upload players_stats.xlsx - not "
        "defense_vs_position.xlsx, which is this script's own output - and run again.")


def load_positions(path):
    raw = pd.read_csv(path) if str(path).lower().endswith(".csv") else pd.read_excel(path)
    for col in ["ID", "Name", "Surname", "Position"]:
        if col not in raw.columns:
            raise KeyError(f"Column '{col}' not found in {path}. Found: {list(raw.columns)}")

    pos = raw.copy()
    pos["Position"] = pos["Position"].astype(str).str.strip().str.title()
    pos = pos[pos["Position"].isin(POSITIONS)].copy()          # drops blank positions
    # Some players have "-" as ID (new players): they become NaN and are matched by name
    pos["player_num"] = pd.to_numeric(pos["ID"], errors="coerce")
    pos["name_key"] = (pos["Name"].fillna("") + " " + pos["Surname"].fillna("")).map(name_key)

    by_id = (pos.dropna(subset=["player_num"]).drop_duplicates("player_num")
             .assign(player_num=lambda d: d["player_num"].astype(int))
             .set_index("player_num")["Position"])
    by_name = pos.drop_duplicates("name_key").set_index("name_key")["Position"]
    print(f"Positions file: {len(pos)} players with a position "
          f"({len(by_id)} with an ID, {len(by_name)} usable by name).")
    return by_id, by_name


# ============================================================================
# 4. DOWNLOAD BOX SCORES  (one season at a time, cached to disk)
# ============================================================================
def download_season(season):
    """Downloads the season's played games. Games already saved from an earlier run
    are kept, so re-running during the season only fetches the NEW games."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_file = os.path.join(CACHE_DIR, f"boxscores_{COMPETITION}{season}.csv")

    cached = None
    if os.path.exists(cache_file) and not FORCE_REFRESH:
        cached = pd.read_csv(cache_file, low_memory=False)
    done = set(cached["Gamecode"].astype(int)) if cached is not None else set()

    api = BoxScoreData(COMPETITION)
    schedule = api.get_gamecodes_season(season)

    # The library's own 'played' flag is unreliable, so we also require a final
    # score. Games with 0-0 have not been played yet.
    if {"homescore", "awayscore"}.issubset(schedule.columns):
        schedule = schedule[(schedule["homescore"] + schedule["awayscore"]) > 0]
    game_codes = sorted(int(c) for c in schedule["gameCode"].unique())
    todo = [c for c in game_codes if c not in done]
    print(f"Season {season}: {len(game_codes)} played games, "
          f"{len(done)} already saved, {len(todo)} to download")

    frames, failed = [], []
    for code in tqdm(todo, desc=f"Season {season}"):
        df = None
        for attempt in range(3):
            try:
                df = api.get_players_boxscore_stats(season, code)
                break
            except (JSONDecodeError, KeyError, ValueError):
                break                                   # no usable data: do not retry
            except requests.exceptions.RequestException:
                time.sleep(2 * (attempt + 1))           # network hiccup: wait and retry
        if df is None or df.empty:
            failed.append(code)
        else:
            frames.append(df)
        time.sleep(PAUSE_SECONDS)

    if failed:
        print(f"  WARNING: {len(failed)} game(s) skipped (no data): {failed}")

    parts = ([cached] if cached is not None else []) + frames
    if not parts:
        raise RuntimeError(f"No box scores available for season {season} yet")
    season_df = pd.concat(parts, ignore_index=True)
    season_df.to_csv(cache_file, index=False)
    return season_df


NAME_COLS = ["Season", "Gamecode", "home_name", "away_name", "Round"]


def get_team_names(season):
    """Full team names per game, taken from the season schedule.
    Returns columns: Season, Gamecode, home_name, away_name (empty if unavailable)."""
    try:
        sched = BoxScoreData(COMPETITION).get_gamecodes_season(season)
        home = _pick(sched, ["hometeam", "home_team"])
        away = _pick(sched, ["awayteam", "away_team"])
        if home is None or away is None:
            print(f"  note: no team-name columns in the season {season} schedule "
                  f"(columns: {list(sched.columns)})")
            return pd.DataFrame(columns=NAME_COLS)
        out = pd.DataFrame({"Season": season,
                            "Gamecode": sched["gameCode"].astype(int),
                            "home_name": sched[home].astype(str).str.strip(),
                            "away_name": sched[away].astype(str).str.strip(),
                            "Round": (pd.to_numeric(sched["Round"], errors="coerce")
                                      if "Round" in sched.columns else np.nan)})
        return out.drop_duplicates(["Season", "Gamecode"])
    except Exception as exc:                      # names are cosmetic: never stop the run
        print(f"  note: could not fetch team names for season {season}: {exc}")
        return pd.DataFrame(columns=NAME_COLS)


def apply_full_names(df, names):
    """Replace the short team labels by the full names from the schedule."""
    if names is not None and not names.empty:
        d = df.merge(names, on=["Season", "Gamecode"], how="left")
        home = d["Home"] == 1
        full = np.where(home, d["home_name"], d["away_name"])
        opp = np.where(home, d["away_name"], d["home_name"])
        ok = pd.notna(full) & pd.notna(opp)
        # Safety check: one short label must always map to the same full name
        chk = pd.DataFrame({"Season": d["Season"], "raw": d["team"], "full": full})[ok]
        if (chk.groupby(["Season", "raw"])["full"].nunique() > 1).any():
            print("  WARNING: team names from the schedule do not line up with the "
                  "box scores - keeping the original team labels.")
        else:
            d["team"] = np.where(ok, full, d["team"])
            d["opp_team"] = np.where(ok, opp, d["opp_team"])
        df = d.drop(columns=["home_name", "away_name"])
    if TEAM_NAME_OVERRIDES:
        df["team"] = df["team"].replace(TEAM_NAME_OVERRIDES)
        df["opp_team"] = df["opp_team"].replace(TEAM_NAME_OVERRIDES)
    return df


# ============================================================================
# 5. CLEAN: one tidy row per player per game, with team + opponent
# ============================================================================
def clean_boxscores(raw, names=None):
    cols = {key: _pick(raw, names) for key, names in COLUMN_ALIASES.items()}
    missing = [k for k in REQUIRED if cols[k] is None]
    if missing:
        raise KeyError(f"Box score columns not found for {missing}. "
                       f"Columns available: {list(raw.columns)}")

    df = pd.DataFrame({"Season": raw["Season"].astype(int),
                       "Gamecode": raw["Gamecode"].astype(int),
                       "Home": raw["Home"].astype(int)})
    df["player"] = raw[cols["player"]].astype(str).str.strip()
    df["team_raw"] = raw[cols["team"]].astype(str).str.strip()
    df["player_num"] = pd.to_numeric(
        raw[cols["player_id"]].astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    df["minutes"] = raw[cols["minutes"]].map(minutes_to_float)

    for key in COLUMN_ALIASES:
        if key in ("player_id", "player", "team", "minutes"):
            continue
        if cols[key] is None:
            print(f"  note: column for '{key}' not found, filled with 0")
            df[key] = 0.0
        else:
            df[key] = pd.to_numeric(raw[cols[key]], errors="coerce").fillna(0.0)

    if cols["reb"] is None:
        df["reb"] = df["oreb"] + df["dreb"]
    df["fgm"] = df["fgm2"] + df["fgm3"]
    df["fga"] = df["fga2"] + df["fga3"]

    # Keep real players who actually played; drop the 'Team' and 'Total' summary rows
    keep = (~df["player"].str.lower().isin(["team", "total"])
            & df["player_num"].notna() & (df["minutes"] > MIN_MINUTES_PLAYED))
    df = df[keep].copy()
    df["player_num"] = df["player_num"].astype(int)

    # Work out who played against whom inside each game
    team_of_side = (df.groupby(["Season", "Gamecode", "Home"])["team_raw"]
                    .agg(lambda s: s.mode().iat[0]).unstack("Home"))
    team_of_side = team_of_side.dropna()                      # need both sides of the game
    team_of_side = team_of_side.rename(columns={0: "away_team", 1: "home_team"}).reset_index()
    df = df.merge(team_of_side, on=["Season", "Gamecode"], how="inner")
    df["team"] = np.where(df["Home"] == 1, df["home_team"], df["away_team"])
    df["opp_team"] = np.where(df["Home"] == 1, df["away_team"], df["home_team"])
    df = df.drop(columns=["home_team", "away_team", "team_raw"])
    return apply_full_names(df, names)


# ============================================================================
# 6. ATTACH POSITIONS
# ============================================================================
def infer_positions(df, players):
    """Guess positions for players missing from the file.

    Plain-language explanation
    - A 'classifier' (here a RandomForest = many small decision trees that vote) learns,
      from players whose position we KNOW, what a guard / forward / center looks
      like statistically (e.g. centers: many rebounds & blocks, few assists).
    - It then applies that learning to players whose position we don't know.
    - 'Per 36 minutes' puts everyone on the same scale (what a player would do in
      36 minutes), so starters and bench players can be compared fairly.
    - 'Cross-validation' hides part of the known players, predicts them, and checks
      the answers. The accuracy it prints tells you how far to trust the guesses.
    """
    from sklearn.ensemble import RandomForestClassifier
    from sklearn.model_selection import StratifiedKFold, cross_val_score

    sums = df.groupby("player_num")[["minutes", "pts", "reb", "oreb", "ast", "stl",
                                     "blk", "tov", "fga2", "fga3", "fta"]].sum()
    sums = sums[sums["minutes"] >= MIN_MINUTES_TO_INFER]
    feats = sums.drop(columns="minutes").div(sums["minutes"], axis=0) * 36
    feats["three_share"] = (sums["fga3"] / (sums["fga2"] + sums["fga3"]).replace(0, np.nan)).fillna(0)

    known = players["Position"].dropna()
    train_ids = feats.index.intersection(known.index)
    unknown_ids = feats.index.difference(known.index)
    y = known.loc[train_ids]

    if len(unknown_ids) == 0:
        print("No players need a guessed position.")
        return players
    if len(train_ids) < 30 or y.nunique() < 2:
        print("Not enough players with known positions to train a guesser - skipping.")
        return players

    model = RandomForestClassifier(n_estimators=300, min_samples_leaf=3,
                                   class_weight="balanced", random_state=0)
    folds = max(2, min(5, int(y.value_counts().min())))
    acc = cross_val_score(model, feats.loc[train_ids], y,
                          cv=StratifiedKFold(folds, shuffle=True, random_state=0)).mean()
    model.fit(feats.loc[train_ids], y)
    guess = pd.Series(model.predict(feats.loc[unknown_ids]), index=unknown_ids)

    players.loc[guess.index, "Position"] = guess
    players.loc[guess.index, "PosSource"] = "inferred"
    print(f"Guessed positions for {len(guess)} players "
          f"(cross-validated accuracy on known players: {acc:.0%}).")
    return players


def attach_positions(df, by_id, by_name):
    players = (df.sort_values(["Season", "Gamecode"])
               .drop_duplicates("player_num", keep="last")[["player_num", "player"]]
               .set_index("player_num"))
    players["Position"] = players.index.map(by_id)
    players["PosSource"] = np.where(players["Position"].notna(), "file (ID)", None)

    # fallback: match by name (covers players with no ID in the file)
    need = players["Position"].isna()
    keys = players.loc[need, "player"].map(name_key)
    by_name_hit = keys.map(by_name)
    players.loc[need, "Position"] = by_name_hit.values
    players.loc[by_name_hit.index[by_name_hit.notna()], "PosSource"] = "file (name)"

    if INFER_MISSING_POSITIONS:
        players = infer_positions(df, players)

    df = df.merge(players[["Position", "PosSource"]], left_on="player_num",
                  right_index=True, how="left")
    df["Position"] = df["Position"].fillna("Unknown")
    df["PosSource"] = df["PosSource"].fillna("none")
    return df, players


# ============================================================================
# 7. DEFENSE vs POSITION  (PIR only)
# ============================================================================
# How the numbers are built
#   Step 1: in every game, take all opposing players of one position who played
#           (e.g. all forwards who played against Panathinaikos) and average
#           their PIR  ->  one number per game per position.
#   Step 2: average those game numbers over all games  ->  PIR_avg.
#           Every game counts once, whether the opponent used 1 or 4 forwards.
#           Games where the opponent used nobody of that position are skipped.
#
# Columns
#   PIR_avg       the result of steps 1-2. Lower = tougher defense vs that position.
#   PIR_per36     cross-check: total PIR / total minutes x 36. It weights players by
#                 minutes played, so a 3-minute cameo cannot distort it.
#   PIR_avg_vs_lg%  how far above (+) / below (-) the league average for that position
#                 the team is. Negative = defends that position better than average.
#   PIR_rank      1 = team allows the LOWEST PIR to that position (best defense).
#   Games         games in which the opponent used at least one player of that position.
#   Appearances   player-games behind the average. Small = not reliable yet.
def _pir_stats(pg, keys):
    """Summary of the per-game position averages, grouped by `keys`."""
    g = pg.groupby(keys)
    out = pd.DataFrame({
        "Games": g.size(),
        "Appearances": g["n_players"].sum(),
        "PIR_avg": g["pir_game_avg"].mean(),
        "PIR_per36": g["pir_total"].sum() / g["min_total"].sum() * 36,
    })
    return out.reset_index()


def build_dvp(df):
    # d = one row per player per game (players who played), positioned players only.
    # "Team" = the team that DEFENDED against that player, i.e. his opponent.
    d = df[df["Position"].isin(POSITIONS)].rename(columns={"opp_team": "Team"})

    # Step 1: one row per (game, defending team, position)
    pg = (d.groupby(["Season", "Gamecode", "Team", "Position"])
          .agg(pir_game_avg=("pir", "mean"), n_players=("pir", "size"),
               pir_total=("pir", "sum"), min_total=("minutes", "sum"))
          .reset_index())

    metrics = ["PIR_avg", "PIR_per36"]
    results = {}
    for label, by in [("by_season", ["Season", "Team"]), ("all_seasons", ["Team"])]:
        lg_keys = [b for b in by if b != "Team"] + ["Position"]
        team = _pir_stats(pg, by + ["Position"])
        league = _pir_stats(pg, lg_keys)[lg_keys + metrics]
        t = team.merge(league, on=lg_keys, suffixes=("", "_lg"))
        for c in metrics:
            t[f"{c}_vs_lg%"] = (t[c] / t[f"{c}_lg"] - 1) * 100
        t = t.drop(columns=[f"{c}_lg" for c in metrics])
        t["PIR_rank"] = t.groupby(lg_keys)["PIR_avg"].rank(method="min").astype("Int64")

        front = by + ["Position", "Games", "Appearances"]
        t = t[front + [c for c in t.columns if c not in front]]
        t = t.sort_values(by[:-1] + ["Position", "PIR_avg"]).reset_index(drop=True)
        for c in t.columns:
            if c.endswith("_vs_lg%"):
                t[c] = t[c].round(1)
            elif t[c].dtype == float:
                t[c] = t[c].round(2)
        results[label] = t

    # Overall figure per team (all three positions together, per player-game)
    tot = pg.groupby("Team")[["pir_total", "n_players"]].sum()
    results["team_overall"] = (tot["pir_total"] / tot["n_players"]).rename("All positions")

    # Same, richer: all positions together, with minutes-weighted PIR and vs-league %
    ov = pg.groupby("Team").agg(Appearances=("n_players", "sum"), pir_total=("pir_total", "sum"),
                                min_total=("min_total", "sum"))
    ov["PIR_avg"] = ov["pir_total"] / ov["Appearances"]
    ov["PIR_per36"] = ov["pir_total"] / ov["min_total"] * 36
    lg_avg = pg["pir_total"].sum() / pg["n_players"].sum()
    lg_36 = pg["pir_total"].sum() / pg["min_total"].sum() * 36
    ov["PIR_avg_vs_lg%"] = (ov["PIR_avg"] / lg_avg - 1) * 100
    ov["PIR_per36_vs_lg%"] = (ov["PIR_per36"] / lg_36 - 1) * 100
    results["overall"] = (ov.drop(columns=["pir_total", "min_total"]).round(2)
                          .rename_axis("Team").reset_index().rename(columns={"Team": "team"})
                          .rename(columns={"Appearances": "n", "PIR_avg": "avg", "PIR_per36": "per36",
                                           "PIR_avg_vs_lg%": "avg_vs", "PIR_per36_vs_lg%": "per36_vs"}))
    return results


def pivots(dvp):
    """Teams in rows, positions in columns."""
    dvp_all = dvp["all_seasons"]

    def piv(col):
        return dvp_all.pivot_table(index="Team", columns="Position", values=col)[POSITIONS]

    avg = piv("PIR_avg").join(dvp["team_overall"])
    avg = avg.sort_values("All positions")                    # best defense first
    per36 = piv("PIR_per36").reindex(avg.index)
    rel = piv("PIR_avg_vs_lg%").reindex(avg.index)
    apps = dvp_all.pivot_table(index="Team", columns="Position",
                               values="Appearances")[POSITIONS].reindex(avg.index)
    return avg.round(2), per36.round(2), rel.round(1), apps


def lookup(dvp_all, team_text, position):
    """Print the PIR line for e.g. ('Panathinaikos', 'Forward')."""
    hit = dvp_all[dvp_all["Team"].str.contains(team_text, case=False, na=False)
                  & (dvp_all["Position"] == position.title())]
    if hit.empty:
        print(f"No match for '{team_text}' / {position}. Team names in the data:")
        print(", ".join(sorted(dvp_all["Team"].unique())))
        return
    cols = ["Team", "Position", "Games", "Appearances", "PIR_avg", "PIR_per36",
            "PIR_avg_vs_lg%", "PIR_rank"]
    show(hit[cols])


# ============================================================================
# 8. SAVE
# ============================================================================
def _autowidth(ws):
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 40)
    ws.freeze_panes = "A2"


def save_all(df, players, dvp, piv):
    avg, per36, rel, apps = piv
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    xlsx = os.path.join(OUTPUT_DIR, "defense_vs_position_PIR.xlsx")

    coverage = (df.groupby("PosSource")["minutes"].sum().rename("minutes").reset_index()
                .assign(share_pct=lambda t: (t["minutes"] / t["minutes"].sum() * 100).round(1)))

    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        avg.to_excel(xw, sheet_name="PIR_avg_pivot")
        per36.to_excel(xw, sheet_name="PIR_per36_pivot")
        rel.to_excel(xw, sheet_name="PIR_vs_league_pct_pivot")
        apps.to_excel(xw, sheet_name="Appearances_pivot")
        dvp["all_seasons"].to_excel(xw, sheet_name="DvP_PIR_long", index=False)
        if df["Season"].nunique() > 1:
            dvp["by_season"].to_excel(xw, sheet_name="DvP_PIR_by_season", index=False)
        coverage.to_excel(xw, sheet_name="Position_coverage", index=False)
        for ws in xw.book.worksheets:
            _autowidth(ws)

    df.to_csv(os.path.join(OUTPUT_DIR, "player_boxscores_with_positions.csv"), index=False)
    players.reset_index().to_csv(os.path.join(OUTPUT_DIR, "players_positions_used.csv"), index=False)
    dvp["all_seasons"].to_csv(os.path.join(OUTPUT_DIR, "dvp_PIR.csv"), index=False)
    print(f"\nSaved to '{OUTPUT_DIR}/': defense_vs_position_PIR.xlsx + CSV files")
    return xlsx


# ============================================================================
# 8b. WEB PAGE  (one self-contained index.html: no server, works offline)
# ============================================================================
HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Defense vs Position (PIR)</title>
<style>
:root{
  color-scheme: light dark;
  --bg:#f9f9f7; --surface:#fcfcfb; --ink:#0b0b0b; --ink2:#52514e; --muted:#6f6d68;
  --line:#e1e0d9; --mid:#f0efec; --cool:#2a78d6; --warm:#e34948; --accent:#2a78d6;
}
@media (prefers-color-scheme: dark){
  :root{ --bg:#0d0d0d; --surface:#1a1a19; --ink:#f0efec; --ink2:#c3c2b7; --muted:#a09e96;
         --line:#2c2c2a; --mid:#383835; --cool:#3987e5; --warm:#e66767; --accent:#3987e5; }
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font:15px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
.wrap{max-width:980px;margin:0 auto;padding:24px 16px 48px}
h1{font-size:22px;margin:0 0 4px;letter-spacing:-.01em}
.sub{color:var(--ink2);margin:0 0 16px}
.how{color:var(--ink2);margin:0 0 18px;max-width:70ch}
.controls{display:flex;flex-wrap:wrap;gap:10px 16px;align-items:center;margin:0 0 12px}
.seg{display:inline-flex;border:1px solid var(--line);border-radius:8px;overflow:hidden;background:var(--surface)}
.seg button{border:0;background:transparent;color:var(--ink2);padding:8px 12px;font:inherit;cursor:pointer}
.seg button+button{border-left:1px solid var(--line)}
.seg button[aria-pressed="true"]{background:var(--accent);color:#fff}
.seg button:focus-visible,input:focus-visible,th button:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
input[type=search]{border:1px solid var(--line);background:var(--surface);color:var(--ink);
  border-radius:8px;padding:8px 10px;font:inherit;min-width:180px}
.legend{display:flex;align-items:center;gap:8px;color:var(--ink2);font-size:13px;margin:0 0 10px}
.bar{height:10px;width:160px;border-radius:5px;border:1px solid var(--line);
  background:linear-gradient(90deg,
    color-mix(in srgb,var(--cool) 55%,var(--mid)),var(--mid),color-mix(in srgb,var(--warm) 55%,var(--mid)))}
.tablewrap{overflow-x:auto;background:var(--surface);border:1px solid var(--line);border-radius:10px}
table{border-collapse:separate;border-spacing:0;width:100%;min-width:640px}
th,td{padding:8px 10px;text-align:center;border-bottom:1px solid var(--line)}
th{font-weight:600;color:var(--ink2);background:var(--surface);position:sticky;top:0}
th button{all:unset;cursor:pointer;display:inline-block;padding:2px 4px}
th:first-child,td:first-child{text-align:left;position:sticky;left:0;background:var(--surface);z-index:1;min-width:190px}
tr:last-child td{border-bottom:0}
td.v{font-variant-numeric:tabular-nums}
td.v b{display:block;font-size:16px;font-weight:600}
td.v small{display:block;color:var(--ink2);font-size:12px;white-space:nowrap}
td.thin{opacity:.6}
td.all{border-left:2px solid var(--line)}
summary{cursor:pointer;color:var(--ink2);font-weight:600;margin:0 0 10px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:12px;margin:0 0 20px}
.card{background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:12px 14px}
.card h3{margin:0 0 6px;font-size:14px;color:var(--ink2);font-weight:600}
.card ol{margin:0;padding-left:20px}
.card li{display:flex;justify-content:space-between;gap:8px}
.card li span:last-child{font-variant-numeric:tabular-nums;color:var(--ink2)}
.card p{margin:0;color:var(--muted);font-size:13px}
.note{color:var(--ink2);font-size:13px;max-width:78ch}
footer{margin-top:20px;color:var(--muted);font-size:13px}
</style>
</head>
<body>
<div class="wrap">
  <h1>Defense vs position, in PIR</h1>
  <p class="sub" id="meta"></p>
  <p class="how">Each cell shows the average PIR that opposing <b>guards, forwards and centers</b> produce
  <b>against</b> that team. <b>Red = more than the league average (softer defense), blue = less (tougher).</b>
  Click a column title to sort.</p>

  <details id="cardsbox" open>
    <summary>Softest and toughest teams by position</summary>
    <div class="cards" id="cards"></div>
  </details>

  <div class="controls">
    <div class="seg" role="group" aria-label="Metric" id="metric">
      <button data-m="avg" aria-pressed="true">Avg PIR</button>
      <button data-m="per36" aria-pressed="false">PIR per 36 min</button>
      <button data-m="vs" aria-pressed="false">vs league %</button>
    </div>
    <input type="search" id="q" placeholder="Filter teams" aria-label="Filter teams">
  </div>
  <div class="legend"><span>Tougher</span><span class="bar"></span><span>Softer</span>
    <span>· faded cells = few games so far (†)</span></div>
  <div class="tablewrap"><table id="tbl"></table></div>
  <noscript><p>This page needs JavaScript to draw the table.</p></noscript>

  <p class="note" id="explain"></p>
  <footer id="foot"></footer>
</div>
<script id="data" type="application/json">__DATA__</script>
<script>
const D = JSON.parse(document.getElementById("data").textContent);
const POS = D.positions, MIN_N = D.min_n;
const cell = {}, over = {};
D.rows.forEach(r => { (cell[r.team] = cell[r.team] || {})[r.pos] = r; });
D.overall.forEach(r => { over[r.team] = r; });
const TEAMS = Object.keys(over).sort();
const METRICS = {
  avg:  {val: r => r.avg,    vs: r => r.avg_vs,   fmt: v => v.toFixed(1)},
  per36:{val: r => r.per36,  vs: r => r.per36_vs, fmt: v => v.toFixed(1)},
  vs:   {val: r => r.avg_vs, vs: r => r.avg_vs,   fmt: v => (v > 0 ? "+" : "") + v.toFixed(1) + "%"}
};
const signed = v => v == null ? "–" : (v > 0 ? "+" : "") + v.toFixed(1) + "%";
let metric = "avg", sortKey = "All", sortDir = -1, query = "";

function entry(team, key) { return key === "All" ? over[team] : (cell[team] || {})[key]; }

function tint(vsPct) {
  if (vsPct == null) return "";
  const pole = vsPct > 0 ? "var(--warm)" : "var(--cool)";
  const k = Math.round(Math.min(Math.abs(vsPct) / 40, 1) * 55);
  return "color-mix(in srgb," + pole + " " + k + "%,var(--mid))";
}

function drawCards() {
  const host = document.getElementById("cards"); host.innerHTML = "";
  POS.forEach(pos => {
    const rows = D.rows.filter(r => r.pos === pos && r.n >= MIN_N && r.avg != null)
                       .sort((a, b) => b.avg - a.avg);
    const card = document.createElement("div"); card.className = "card";
    const h = document.createElement("h3"); h.textContent = pos + "s: softest / toughest"; card.appendChild(h);
    if (rows.length < 6) {
      const p = document.createElement("p");
      p.textContent = "Not enough games yet (needs " + MIN_N + "+ player-games per team).";
      card.appendChild(p);
    } else {
      [["Softest", rows.slice(0, 3)], ["Toughest", rows.slice(-3).reverse()]].forEach(([lab, list]) => {
        const t = document.createElement("h3"); t.textContent = lab; t.style.marginTop = "8px"; card.appendChild(t);
        const ol = document.createElement("ol");
        list.forEach(r => {
          const li = document.createElement("li");
          const a = document.createElement("span"); a.textContent = r.team;
          const b = document.createElement("span"); b.textContent = r.avg.toFixed(1) + " (" + signed(r.avg_vs) + ")";
          li.append(a, b); ol.appendChild(li);
        });
        card.appendChild(ol);
      });
    }
    host.appendChild(card);
  });
}

function drawTable() {
  const m = METRICS[metric], tbl = document.getElementById("tbl");
  tbl.innerHTML = "";
  const cols = [["Team", "Team"]].concat(POS.map(p => [p, p]), [["All positions", "All"]]);
  const thead = tbl.createTHead().insertRow();
  cols.forEach(([label, key]) => {
    const th = document.createElement("th");
    th.setAttribute("aria-sort", key === sortKey ? (sortDir > 0 ? "ascending" : "descending") : "none");
    const b = document.createElement("button");
    b.textContent = label + (key === sortKey ? (sortDir > 0 ? " ▲" : " ▼") : "");
    b.onclick = () => { if (sortKey === key) sortDir = -sortDir; else { sortKey = key; sortDir = key === "Team" ? 1 : -1; } drawTable(); };
    th.appendChild(b); thead.appendChild(th);
  });
  const sortVal = t => {
    if (sortKey === "Team") return t;
    const e = entry(t, sortKey); const v = e ? m.val(e) : null;
    return v == null ? null : v;
  };
  const list = TEAMS.filter(t => t.toLowerCase().includes(query.toLowerCase()));
  list.sort((a, b) => {
    const x = sortVal(a), y = sortVal(b);
    if (x == null && y == null) return 0; if (x == null) return 1; if (y == null) return -1;
    return (x < y ? -1 : x > y ? 1 : 0) * sortDir;
  });
  const tbody = tbl.createTBody();
  list.forEach(team => {
    const tr = tbody.insertRow();
    tr.insertCell().textContent = team;
    POS.concat(["All"]).forEach(key => {
      const td = tr.insertCell(); td.className = "v" + (key === "All" ? " all" : "");
      const e = entry(team, key);
      if (!e || m.val(e) == null) { td.textContent = "–"; return; }
      const thin = e.n < MIN_N; if (thin) td.classList.add("thin");
      td.style.background = tint(m.vs(e));
      const b = document.createElement("b"); b.textContent = m.fmt(m.val(e)) + (thin ? " †" : "");
      const s = document.createElement("small");
      s.textContent = metric === "vs" ? e.avg.toFixed(1) + " PIR · n " + e.n
                                      : "n " + e.n + " · " + signed(m.vs(e));
      td.append(b, s);
      td.title = team + " vs " + (key === "All" ? "all positions" : key + "s") +
        "\nAvg PIR " + (e.avg == null ? "–" : e.avg.toFixed(2)) +
        " | per 36 min " + (e.per36 == null ? "–" : e.per36.toFixed(2)) +
        "\nPlayer-games " + e.n + (e.games != null ? " in " + e.games + " games" : "") +
        (e.rank != null ? "\nRank " + e.rank + " (1 = allows least PIR)" : "");
    });
  });
}

document.getElementById("meta").textContent =
  D.meta.season_label + " · " + D.meta.games + " games" +
  (D.meta.latest_round ? " · through round " + D.meta.latest_round : "") +
  " · updated " + new Date(D.meta.updated_iso).toLocaleString([], {dateStyle: "medium", timeStyle: "short"});
document.getElementById("explain").textContent =
  "n = player-games behind the number; † = fewer than " + MIN_N + " (about 30+ is a comfortable base). " +
  "Avg PIR: in every game the PIR of the opposing players of that position who played is averaged, " +
  "then those game values are averaged, so each game counts once. PIR per 36 min weights players by " +
  "minutes played. “vs league %” compares with the league average for that position. " +
  "The opponent's own strength is not adjusted for.";
document.getElementById("foot").textContent =
  "Positions come from a fantasy player list" +
  (D.meta.inferred_pct > 0 ? "; " + D.meta.inferred_pct.toFixed(1) + "% of minutes are from players whose position was estimated from their stats" : "") +
  (D.meta.none_pct > 0 ? "; " + D.meta.none_pct.toFixed(1) + "% of minutes are from players without a position and are left out" : "") + ".";
document.querySelectorAll("#metric button").forEach(b => b.onclick = () => {
  metric = b.dataset.m;
  document.querySelectorAll("#metric button").forEach(x => x.setAttribute("aria-pressed", x === b));
  drawTable();
});
document.getElementById("q").oninput = e => { query = e.target.value; drawTable(); };
if (window.innerWidth < 640) document.getElementById("cardsbox").open = false;
drawCards(); drawTable();
</script>
</body>
</html>
"""


def write_html(df, dvp):
    """Write OUTPUT_DIR/index.html from the finished tables."""
    import json
    from datetime import datetime, timezone

    ren = {"Team": "team", "Position": "pos", "Games": "games", "Appearances": "n",
           "PIR_avg": "avg", "PIR_per36": "per36", "PIR_avg_vs_lg%": "avg_vs",
           "PIR_per36_vs_lg%": "per36_vs", "PIR_rank": "rank"}
    rows = json.loads(dvp["all_seasons"][list(ren)].rename(columns=ren).to_json(orient="records"))
    overall = json.loads(dvp["overall"].to_json(orient="records"))

    cov = df.groupby("PosSource")["minutes"].sum()
    total = cov.sum() or 1.0
    seasons = sorted(int(x) for x in df["Season"].unique())
    label = ", ".join(f"{x}-{str(x + 1)[-2:]}" for x in seasons)
    latest_round = None
    if "Round" in df.columns and df["Round"].notna().any():
        latest_round = int(df["Round"].max())

    payload = {
        "positions": POSITIONS, "min_n": 15, "rows": rows, "overall": overall,
        "meta": {
            "season_label": f"Euroleague {label}",
            "games": int(df[["Season", "Gamecode"]].drop_duplicates().shape[0]),
            "latest_round": latest_round,
            "updated_iso": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "inferred_pct": round(float(cov.get("inferred", 0.0)) / total * 100, 1),
            "none_pct": round(float(cov.get("none", 0.0)) / total * 100, 1),
        },
    }
    blob = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, "index.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(HTML_TEMPLATE.replace("__DATA__", blob))
    return path


# ============================================================================
# 9. RUN
# ============================================================================
def main():
    by_id, by_name = load_positions(get_positions_path())

    raw = pd.concat([download_season(s) for s in SEASONS], ignore_index=True)
    print(f"\nDownloaded rows: {len(raw):,}")

    names = pd.concat([get_team_names(s) for s in SEASONS], ignore_index=True)
    df = clean_boxscores(raw, names)
    print("Teams found: " + ", ".join(sorted(df["team"].unique())))
    print(f"Clean player-game rows: {len(df):,} "
          f"({df[['Season', 'Gamecode']].drop_duplicates().shape[0]} games)")

    df, players = attach_positions(df, by_id, by_name)

    cov = df.groupby("PosSource")["minutes"].sum()
    print("\nShare of minutes by position source:")
    print((cov / cov.sum() * 100).round(1).astype(str).add(" %").to_string())
    unmatched = (df[df["Position"] == "Unknown"].groupby("player")["minutes"].sum()
                 .sort_values(ascending=False).head(10))
    if len(unmatched):
        print("\nTop players left without a position (excluded from the table):")
        print(unmatched.round(0).to_string())

    dvp = build_dvp(df)
    piv = pivots(dvp)
    avg, per36, rel, apps = piv

    print("\n=== AVERAGE PIR OF OPPOSING PLAYERS BY POSITION (average of per-game averages) ===")
    print("Rows = the DEFENDING team. Lower = tougher defense. Best defense first.")
    show(avg)
    print("\n=== SAME, AS % ABOVE (+) / BELOW (-) LEAGUE AVERAGE FOR THAT POSITION ===")
    show(rel)
    print("\n=== CROSS-CHECK: PIR PER 36 MINUTES (weights players by minutes played) ===")
    show(per36)
    print("\n=== NUMBER OF PLAYER-GAMES BEHIND EACH AVERAGE (small = unreliable) ===")
    show(apps)

    for team_text, position in LOOKUPS:
        print(f"\n--- Average PIR of {position}s against '{team_text}' ---")
        lookup(dvp["all_seasons"], team_text, position)

    xlsx = save_all(df, players, dvp, piv)
    html_path = write_html(df, dvp) if WRITE_HTML else None
    if html_path:
        print(f"Web page written: {html_path}  (open it in any browser)")
    try:
        from google.colab import files
        files.download(xlsx)
        if html_path:
            files.download(html_path)
    except ImportError:
        pass
    return df, dvp


if __name__ == "__main__":
    df_boxscores, dvp_tables = main()
