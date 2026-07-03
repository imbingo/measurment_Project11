import sqlite3
import tempfile
import unittest
from pathlib import Path

import metrology_data_platform_v2_7 as app


class DatabaseMigrationSafetyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.old_db_file = app.DB_FILE
        self.old_backup_dir = app.BACKUP_DIR
        self.db_path = Path(self.tmp.name) / "legacy.db"
        app.DB_FILE = str(self.db_path)
        app.BACKUP_DIR = Path(self.tmp.name) / "backup"

    def tearDown(self):
        app.DB_FILE = self.old_db_file
        app.BACKUP_DIR = self.old_backup_dir
        self.tmp.cleanup()

    def _create_legacy_db_with_data(self):
        conn = sqlite3.connect(str(self.db_path))
        conn.execute(
            """
            CREATE TABLE production_config (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                production_code TEXT UNIQUE NOT NULL,
                production_name TEXT,
                product_model TEXT,
                process_version TEXT,
                description TEXT,
                status TEXT DEFAULT 'enabled',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT
            )
            """
        )
        conn.execute(
            "INSERT INTO production_config (production_code, production_name, updated_at) VALUES (?, ?, ?)",
            ("LEGACY_PROD", "Legacy production", "2026-07-03 08:00:00"),
        )
        conn.execute(
            """
            CREATE TABLE collect_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                production_code TEXT,
                status TEXT,
                message TEXT,
                created_at TEXT
            )
            """
        )
        conn.execute(
            "INSERT INTO collect_log (production_code, status, message, created_at) VALUES (?, ?, ?, ?)",
            ("LEGACY_PROD", "SUCCESS", "legacy log", "2026-07-03 08:01:00"),
        )
        conn.commit()
        conn.close()

    def test_init_db_backs_up_existing_db_before_schema_migration_and_preserves_data(self):
        self._create_legacy_db_with_data()
        app.init_db()

        backups = list(app.BACKUP_DIR.glob("legacy_*_to_V2.7.1_*.db"))
        self.assertEqual(len(backups), 1)

        conn = app.get_conn()
        prod = conn.execute("SELECT * FROM production_config WHERE production_code='LEGACY_PROD'").fetchone()
        schema = conn.execute("SELECT value FROM schema_meta WHERE key='schema_version'").fetchone()
        backup_meta = conn.execute("SELECT value FROM schema_meta WHERE key='last_migration_backup'").fetchone()
        conn.close()
        self.assertIsNotNone(prod)
        self.assertEqual(prod["production_name"], "Legacy production")
        self.assertEqual(schema["value"], app.SCHEMA_VERSION)
        self.assertTrue(Path(backup_meta["value"]).exists())

    def test_collect_log_clear_and_prune_are_disabled(self):
        app.init_db()
        conn = app.get_conn()
        conn.execute(
            "INSERT INTO collect_log (production_code, status, message, created_at) VALUES (?, ?, ?, ?)",
            ("PROD", "SUCCESS", "keep me", app.now_str()),
        )
        conn.commit()
        conn.close()

        with self.assertRaises(RuntimeError):
            app.clear_collect_logs()
        with self.assertRaises(RuntimeError):
            app.prune_collect_logs(1)

        conn = app.get_conn()
        count = conn.execute("SELECT COUNT(*) AS c FROM collect_log").fetchone()["c"]
        conn.close()
        self.assertEqual(count, 1)


if __name__ == "__main__":
    unittest.main()
