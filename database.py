import json
import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "colisseum.db"


def _connect():
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def initialize_database():
    with _connect() as connection:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS players (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                surname TEXT NOT NULL,
                age INTEGER NOT NULL,
                nickname TEXT NOT NULL COLLATE NOCASE UNIQUE,
                password TEXT NOT NULL,
                phone TEXT NOT NULL,
                points INTEGER NOT NULL DEFAULT 0,
                wins INTEGER NOT NULL DEFAULT 0,
                losses INTEGER NOT NULL DEFAULT 0,
                draws INTEGER NOT NULL DEFAULT 0,
                scouting_json TEXT
            );

            CREATE TABLE IF NOT EXISTS tournaments (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                subtitle TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'En Curso',
                champion TEXT,
                bracket_json TEXT,
                metricas_json TEXT
            );

            CREATE TABLE IF NOT EXISTS registrations (
                tournament_id INTEGER NOT NULL,
                player_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                surname TEXT NOT NULL,
                phone TEXT NOT NULL,
                nickname TEXT NOT NULL,
                stats_json TEXT,
                PRIMARY KEY (tournament_id, player_id),
                FOREIGN KEY (tournament_id) REFERENCES tournaments(id) ON DELETE CASCADE,
                FOREIGN KEY (player_id) REFERENCES players(id) ON DELETE CASCADE
            );
            """
        )
        tournament_columns = {row["name"] for row in connection.execute("PRAGMA table_info(tournaments)")}
        if "metricas_json" not in tournament_columns:
            connection.execute("ALTER TABLE tournaments ADD COLUMN metricas_json TEXT")
        registration_columns = {row["name"] for row in connection.execute("PRAGMA table_info(registrations)")}
        if "stats_json" not in registration_columns:
            connection.execute("ALTER TABLE registrations ADD COLUMN stats_json TEXT")


def load_state():
    initialize_database()
    with _connect() as connection:
        players = []
        for row in connection.execute("SELECT * FROM players ORDER BY id"):
            player = dict(row)
            player["scouting"] = json.loads(player["scouting_json"]) if player["scouting_json"] else None
            player.pop("scouting_json", None)
            players.append(player)

        tournaments = []
        registrations = {}
        brackets = {}
        for row in connection.execute("SELECT * FROM tournaments ORDER BY id"):
            tournament = {
                "id": row["id"],
                "titulo": row["title"],
                "subtitulo": row["subtitle"],
                "metricas": json.loads(row["metricas_json"]) if row["metricas_json"] else [],
            }
            tournaments.append(tournament)
            key = f"{row['title']} ({row['subtitle']})" if row["subtitle"] else row["title"]
            registrations[key] = []
            if row["bracket_json"]:
                brackets[key] = json.loads(row["bracket_json"])

        rows = connection.execute(
            "SELECT t.title, t.subtitle, r.name, r.surname, r.phone, r.nickname, r.stats_json "
            "FROM registrations r JOIN tournaments t ON t.id = r.tournament_id "
            "ORDER BY r.tournament_id, r.player_id"
        )
        for row in rows:
            key = f"{row['title']} ({row['subtitle']})" if row["subtitle"] else row["title"]
            registrations.setdefault(key, []).append({
                "name": row["name"],
                "surname": row["surname"],
                "phone": row["phone"],
                "nickname": row["nickname"],
                "estadisticas": json.loads(row["stats_json"]) if row["stats_json"] else {},
            })

        return players, tournaments, registrations, brackets


def save_state(players, tournaments, registrations, brackets):
    initialize_database()
    with _connect() as connection:
        connection.execute("DELETE FROM registrations")
        connection.execute("DELETE FROM tournaments")
        connection.execute("DELETE FROM players")

        for player in players:
            scouting = player.get("scouting")
            connection.execute(
                "INSERT INTO players "
                "(id, name, surname, age, nickname, password, phone, points, wins, losses, draws, scouting_json) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    player["id"], player.get("name", ""), player.get("surname", ""),
                    player.get("age", 0), player["nickname"], player.get("password", ""),
                    player.get("phone", ""), player.get("points", 0), player.get("wins", 0),
                    player.get("losses", 0), player.get("draws", 0),
                    json.dumps(scouting) if scouting is not None else None,
                ),
            )

        for tournament in tournaments:
            title = tournament["titulo"]
            subtitle = tournament.get("subtitulo", "")
            key = f"{title} ({subtitle})" if subtitle else title
            bracket = brackets.get(key)
            status = bracket.get("estado", "En Curso") if bracket else "En Curso"
            champion = bracket.get("campeon") if bracket else None
            connection.execute(
                "INSERT INTO tournaments (id, title, subtitle, status, champion, bracket_json, metricas_json) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (tournament["id"], title, subtitle, status, champion, json.dumps(bracket) if bracket else None, json.dumps(tournament.get("metricas", []))),
            )
            tournament_id = connection.execute("SELECT last_insert_rowid()").fetchone()[0]
            for participant in registrations.get(key, []):
                player = next(
                    (item for item in players if item["nickname"].casefold() == participant["nickname"].casefold()),
                    None,
                )
                if player:
                    connection.execute(
                        "INSERT INTO registrations (tournament_id, player_id, name, surname, phone, nickname, stats_json) VALUES (?, ?, ?, ?, ?, ?, ?)",
                        (tournament_id, player["id"], participant["name"], participant["surname"], participant["phone"], participant["nickname"], json.dumps(participant.get("estadisticas", {}))),
                    )
