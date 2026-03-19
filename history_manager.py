"""
HistoryManager for WinSet - logs all setting changes to a persistent database.
"""

import sqlite3
import os
from datetime import datetime
from typing import List, Tuple, Any, Optional


class HistoryManager:
    """Handles logging and retrieving setting changes from a local SQLite DB."""

    def __init__(self, db_path: Optional[str] = None):
        if db_path is None:
            app_data_path = os.getenv('LOCALAPPDATA')
            if not app_data_path:
                app_data_path = os.path.expanduser("~")  # Fallback for safety
            db_dir = os.path.join(app_data_path, 'WinSet')
            os.makedirs(db_dir, exist_ok=True)
            self.db_path = os.path.join(db_dir, 'history.db')
        else:
            self.db_path = db_path  # Allow override for testing

        self._conn = sqlite3.connect(self.db_path, check_same_thread=False)
        self._create_table()

    def _create_table(self):
        """Create the history table if it doesn't exist."""
        cursor = self._conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS changes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                setting_id TEXT NOT NULL,
                setting_name TEXT NOT NULL,
                old_value TEXT,
                new_value TEXT,
                value_type TEXT NOT NULL,
                hive TEXT NOT NULL,
                key_path TEXT NOT NULL,
                value_name TEXT NOT NULL
            )
        """)
        self._conn.commit()

    def log_change(self, setting: Any, old_value: Any, new_value: Any):
        """Log a single setting change to the database."""
        cursor = self._conn.cursor()
        old_val_str = str(old_value) if old_value is not None else 'N/A'
        new_val_str = str(new_value) if new_value is not None else 'N/A'

        cursor.execute("""
            INSERT INTO changes (timestamp, setting_id, setting_name, old_value, new_value, value_type, hive, key_path, value_name)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.now().isoformat(sep=' ', timespec='seconds'),
            setting.id, setting.name, old_val_str, new_val_str,
            setting.value_type, setting.hive, setting.key_path, setting.value_name
        ))
        self._conn.commit()

    def get_history(self) -> List[Tuple[Any, ...]]:
        """Retrieve all changes, sorted by most recent first."""
        cursor = self._conn.cursor()
        cursor.execute("SELECT id, timestamp, setting_name, old_value, new_value FROM changes ORDER BY timestamp DESC")
        return cursor.fetchall()

    def get_change_details(self, change_id: int) -> Optional[Tuple[Any, ...]]:
        """Get the details needed to revert a specific change by its ID."""
        cursor = self._conn.cursor()
        cursor.execute("SELECT hive, key_path, value_name, value_type, old_value FROM changes WHERE id = ?", (change_id,))
        return cursor.fetchone()

    def close(self):
        """Close the database connection."""
        self._conn.close()