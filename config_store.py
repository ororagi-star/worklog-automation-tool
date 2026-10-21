from __future__ import annotations

import json
import sqlite3
from copy import deepcopy
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

from worklog_config import validate_config


class ConfigConflict(ValueError):
    pass


class SessionConfigStore:
    database_url = ""
    session_only = True

    def __init__(self, state):
        self.state = state

    def initialize(self, defaults):
        if "temporary_worklog_settings" not in self.state:
            self.state["temporary_worklog_settings"] = [(1, validate_config(defaults), datetime.now(timezone.utc).isoformat())]

    def load(self):
        revision, config, updated = self.state["temporary_worklog_settings"][-1]
        return deepcopy(config), revision, updated

    def save(self, config, expected_revision):
        config = validate_config(config)
        _, revision, _ = self.load()
        if revision != expected_revision:
            raise ConfigConflict("설정이 변경되었습니다. 다시 불러와 주세요.")
        history = list(self.state["temporary_worklog_settings"])
        history.append((revision + 1, config, datetime.now(timezone.utc).isoformat()))
        self.state["temporary_worklog_settings"] = history[-30:]
        return revision + 1

    def history(self):
        return [(revision, updated) for revision, _, updated in reversed(self.state["temporary_worklog_settings"])]

    def restore(self, revision, expected_revision):
        config = next((config for version, config, _ in self.state["temporary_worklog_settings"] if version == revision), None)
        if config is None:
            raise ValueError("복원할 설정을 찾지 못했습니다.")
        return self.save(config, expected_revision)


class ConfigStore:
    def __init__(self, local_path: Path, database_url: str = ""):
        self.local_path = Path(local_path)
        self.database_url = database_url
        if database_url and not database_url.startswith(("postgres://", "postgresql://")):
            raise ValueError("DATABASE_URL에는 PostgreSQL 접속 주소를 설정해 주세요.")

    @contextmanager
    def connection(self):
        if self.database_url:
            import psycopg

            connection = psycopg.connect(self.database_url, connect_timeout=10)
        else:
            self.local_path.parent.mkdir(parents=True, exist_ok=True)
            connection = sqlite3.connect(self.local_path, timeout=10)
        try:
            with connection:
                yield connection
        finally:
            connection.close()

    def execute(self, connection, query, params=()):
        if self.database_url:
            query = query.replace("?", "%s")
        return connection.execute(query, params)

    def initialize(self, defaults: dict):
        payload = json.dumps(validate_config(defaults), ensure_ascii=False)
        now = datetime.now(timezone.utc).isoformat()
        with self.connection() as connection:
            self.execute(connection, "CREATE TABLE IF NOT EXISTS worklog_settings (id INTEGER PRIMARY KEY, revision INTEGER NOT NULL, payload TEXT NOT NULL, updated_at TEXT NOT NULL)")
            self.execute(connection, "CREATE TABLE IF NOT EXISTS worklog_settings_history (revision INTEGER PRIMARY KEY, payload TEXT NOT NULL, updated_at TEXT NOT NULL)")
            self.execute(connection, "INSERT INTO worklog_settings (id, revision, payload, updated_at) VALUES (1, 1, ?, ?) ON CONFLICT (id) DO NOTHING", (payload, now))
            row = self.execute(connection, "SELECT revision, payload, updated_at FROM worklog_settings WHERE id = 1").fetchone()
            self.execute(connection, "INSERT INTO worklog_settings_history (revision, payload, updated_at) VALUES (?, ?, ?) ON CONFLICT (revision) DO NOTHING", row)

    def load(self):
        with self.connection() as connection:
            revision, payload, updated_at = self.execute(connection, "SELECT revision, payload, updated_at FROM worklog_settings WHERE id = 1").fetchone()
        return validate_config(json.loads(payload)), revision, updated_at

    def save(self, config: dict, expected_revision: int):
        config = validate_config(config)
        payload = json.dumps(config, ensure_ascii=False)
        now = datetime.now(timezone.utc).isoformat()
        with self.connection() as connection:
            cursor = self.execute(connection, "UPDATE worklog_settings SET revision = ?, payload = ?, updated_at = ? WHERE id = 1 AND revision = ?", (expected_revision + 1, payload, now, expected_revision))
            if cursor.rowcount != 1:
                raise ConfigConflict("다른 창에서 설정이 변경되었습니다. 새로 불러온 뒤 다시 수정해 주세요.")
            self.execute(connection, "INSERT INTO worklog_settings_history (revision, payload, updated_at) VALUES (?, ?, ?)", (expected_revision + 1, payload, now))
        return expected_revision + 1

    def history(self):
        with self.connection() as connection:
            return self.execute(connection, "SELECT revision, updated_at FROM worklog_settings_history ORDER BY revision DESC LIMIT 30").fetchall()

    def restore(self, revision: int, expected_revision: int):
        with self.connection() as connection:
            row = self.execute(connection, "SELECT payload FROM worklog_settings_history WHERE revision = ?", (revision,)).fetchone()
        if not row:
            raise ValueError("복원할 설정을 찾지 못했습니다.")
        return self.save(json.loads(row[0]), expected_revision)
