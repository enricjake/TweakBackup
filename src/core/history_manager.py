"""
HistoryManager for WinSet - logs all setting changes to a persistent database.
"""

import sqlite3  # Used to interface with the local SQLite database that stores change history
import os  # Used to resolve file system paths (e.g., LOCALAPPDATA) and check file existence
import threading  # Provides thread-safe locking so multiple threads can safely read/write the DB
from datetime import (
    datetime,
    timedelta,
)  # datetime for timestamps; timedelta for computing date offsets during pruning
from typing import (
    List,
    Tuple,
    Any,
    Optional,
    Dict,
)  # Type hints for clearer function signatures


class HistoryManager:
    """Handles logging and retrieving setting changes from a local SQLite DB."""

    # Maximum number of history entries to keep in the database. When the row count
    # exceeds this value, the oldest entries are pruned. This prevents the database
    # from growing indefinitely on machines that make frequent Registry changes.
    MAX_HISTORY_ENTRIES = 10000

    # Default number of days to retain history entries. Entries older than this are
    # removed when prune_old_history() is called without an explicit days_to_keep value.
    DEFAULT_RETENTION_DAYS = 90

    def __init__(self, db_path: Optional[str] = None):
        # db_path: Optional override for the database file location. When None (the
        #   typical case), the DB is stored under %LOCALAPPDATA%/WinSet/history.db so
        #   that each Windows user gets their own history. Passing a custom path is
        #   useful for unit tests or portable deployments.
        if db_path is None:
            # Resolve the user's LOCALAPPDATA folder (e.g., C:\Users\<user>\AppData\Local)
            app_data_path = os.getenv("LOCALAPPDATA")
            if not app_data_path:
                # Fallback to the user home directory if LOCALAPPDATA is unset
                app_data_path = os.path.expanduser("~")  # Fallback for safety
            # Build the WinSet data directory path and create it if missing
            db_dir = os.path.join(app_data_path, "WinSet")
            os.makedirs(db_dir, exist_ok=True)
            # Final path to the SQLite database file
            self.db_path = os.path.join(db_dir, "history.db")
        else:
            self.db_path = db_path  # Allow override for testing

        # _conn: The SQLite connection object shared across all operations.
        #   check_same_thread=False lets us use this connection from any thread
        #   (guarded by _lock below).
        self._conn = sqlite3.connect(self.db_path, check_same_thread=False)

        if os.name == "nt":
            try:
                import win32security
                import win32api
                import ntsecuritycon

                sid = win32security.LookupAccountName(None, win32api.GetUserName())[0]
                sd = win32security.SECURITY_DESCRIPTOR()
                sd.SetSecurityDescriptorOwner(sid, False)
                dacl = win32security.ACL()
                dacl.AddAccessAllowedAce(
                    win32security.ACL_REVISION, ntsecuritycon.FILE_ALL_ACCESS, sid
                )
                sd.SetSecurityDescriptorDacl(1, dacl, 0)
                win32security.SetFileSecurity(
                    self.db_path, win32security.DACL_SECURITY_INFORMATION, sd
                )
            except Exception:
                pass

        # _lock: A reentrant lock that serializes all database access so that
        #   concurrent threads (e.g., a background sync and the UI) never corrupt
        #   the database.
        self._lock = threading.RLock()

        # Enable Write-Ahead Logging so that readers do not block writers and
        # vice versa, improving concurrency on the SQLite database.
        self._conn.execute("PRAGMA journal_mode=WAL")  # Better concurrency

        # Create the 'changes' table if it does not yet exist.
        self._create_table()
        # Create performance indexes on commonly-queried columns.
        self._create_indexes()

    def __enter__(self):
        """Support context manager usage."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Ensure connection is closed when exiting context."""
        # exc_type, exc_val, exc_tb: Standard context-manager exception parameters
        #   indicating whether an exception occurred inside the with-block.
        self.close()

    def _create_table(self):
        """Create the history table if it doesn't exist."""
        with self._lock:
            # cursor: SQLite cursor used to execute SQL statements against the connection.
            cursor = self._conn.cursor()
            cursor.execute(
                """
            CREATE TABLE IF NOT EXISTS changes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,  -- Unique row ID, auto-generated
                timestamp TEXT NOT NULL,               -- ISO-8601 date/time the change occurred
                setting_id TEXT NOT NULL,              -- Identifier of the WinSet setting that was changed
                setting_name TEXT NOT NULL,            -- Human-readable name of the setting
                old_value TEXT,                        -- Previous value before the change (serialized as string)
                new_value TEXT,                        -- New value after the change (serialized as string)
                value_type TEXT NOT NULL,              -- Registry value type (e.g., REG_DWORD, REG_SZ)
                hive TEXT NOT NULL,                    -- Registry hive (e.g., HKLM, HKCU)
                key_path TEXT NOT NULL,                -- Full Registry key path (e.g., SOFTWARE\\Microsoft\\Windows)
                value_name TEXT NOT NULL,              -- Name of the Registry value within the key
                success INTEGER DEFAULT 1,             -- 1 if the write succeeded, 0 if it failed
                reverted INTEGER DEFAULT 0             -- 1 if this change was later reverted
            )
            """
            )

            # Migration: ensure the 'success' column exists for databases that were
            # created before this column was introduced.
            try:
                cursor.execute("SELECT success FROM changes LIMIT 1")
            except sqlite3.OperationalError:
                # Column missing -- add it with a default of 1 (assumed successful).
                cursor.execute(
                    "ALTER TABLE changes ADD COLUMN success INTEGER DEFAULT 1"
                )
            self._conn.commit()

    def _create_indexes(self):
        """Create indexes for better query performance."""
        with self._lock:
            # cursor: Reused cursor for executing index-creation statements.
            cursor = self._conn.cursor()

            # Index on timestamp (descending) speeds up queries that fetch the most
            # recent changes (e.g., get_history).
            cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_timestamp
            ON changes(timestamp DESC)
        """)

            # Index on setting_id speeds up queries that filter history by a specific
            # WinSet setting (e.g., get_history_by_setting).
            cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_setting_id
            ON changes(setting_id)
        """)

            # Index on success speeds up queries that filter to only successful
            # or failed changes.
            cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_success
            ON changes(success)
        """)
            self._conn.commit()

    def log_change(
        self, setting: Any, old_value: Any, new_value: Any, success: bool = True
    ):
        """
        Log a single setting change to the database.

        Args:
            setting: The Setting object being changed (must expose id, name, value_type, hive, key_path, value_name attributes)
            old_value: Previous value of the setting before the change
            new_value: New value applied to the setting
            success: Whether the Registry write operation succeeded
        """
        # Sanitize values to prevent SQL injection (though parameterized queries handle this)
        # old_val_str / new_val_str: String representations of the old and new values,
        #   converted safely via _sanitize_value so that None, binary data, dicts, and
        #   lists are handled gracefully.
        old_val_str = self._sanitize_value(old_value)
        new_val_str = self._sanitize_value(new_value)

        # Limit string lengths to 1000 characters to prevent excessive storage usage
        # for very large values (e.g., multi-KB Registry blobs).
        old_val_str = old_val_str[:1000] if old_val_str else "N/A"
        new_val_str = new_val_str[:1000] if new_val_str else "N/A"

        with self._lock:
            # cursor: Cursor for executing the INSERT statement.
            cursor = self._conn.cursor()
            cursor.execute(
                """
            INSERT INTO changes (
                timestamp, setting_id, setting_name, old_value, new_value,
                value_type, hive, key_path, value_name, success
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
                (
                    # timestamp: Current date/time formatted as "YYYY-MM-DD HH:MM:SS" for readability.
                    datetime.now().isoformat(sep=" ", timespec="seconds"),
                    # setting.id truncated to 100 chars to match column expectations.
                    getattr(setting, "id", "unknown")[:100],
                    # setting.name truncated to 200 chars.
                    getattr(setting, "name", "unknown")[:200],
                    # old / new values already sanitized and truncated above.
                    old_val_str,
                    new_val_str,
                    # Registry value type string (e.g., "REG_DWORD"), truncated to 50 chars.
                    getattr(setting, "value_type", "unknown")[:50],
                    # Registry hive string (e.g., "HKEY_LOCAL_MACHINE"), truncated to 50 chars.
                    getattr(setting, "hive", "unknown")[:50],
                    # Full Registry key path, truncated to 512 chars.
                    getattr(setting, "key_path", "unknown")[:512],
                    # Name of the value within the Registry key, truncated to 255 chars.
                    getattr(setting, "value_name", "unknown")[:255],
                    # Store 1 for success, 0 for failure.
                    1 if success else 0,
                ),
            )
            self._conn.commit()

        # After logging, check if the table has grown beyond the limit and prune if necessary.
        self._prune_if_needed()

    def _sanitize_value(self, value: Any) -> str:
        """Safely convert a value to string for storage."""
        # value: The raw Python value (could be None, int, float, bool, bytes, dict, list, or str).
        if value is None:
            return "N/A"
        if isinstance(value, (int, float, bool)):
            return str(value)
        if isinstance(value, (bytes, bytearray)):
            # Represent binary blobs as a placeholder to avoid storing raw bytes as text.
            return f"<binary data: {len(value)} bytes>"
        if isinstance(value, dict):
            # Limit dictionary representation to first 10 keys, each value capped at 100 chars,
            # and the overall string at 500 chars to avoid massive output.
            return str({k: str(v)[:100] for k, v in list(value.items())[:10]})[:500]
        if isinstance(value, list):
            # Limit list representation to first 10 items, each capped at 100 chars,
            # overall string at 500 chars.
            return str([str(v)[:100] for v in value[:10]])[:500]
        # For any other type, convert to string and truncate to 1000 characters.
        return str(value)[:1000]

    def _prune_if_needed(self):
        """Check if we need to prune old entries and do so."""
        with self._lock:
            # cursor: Used to count current rows and delete excess ones.
            cursor = self._conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM changes")
            # count: Total number of rows currently in the changes table.
            count = cursor.fetchone()[0]

        if count > self.MAX_HISTORY_ENTRIES:
            # to_delete: Number of oldest rows to remove so the table drops to 80% of
            #   MAX_HISTORY_ENTRIES, providing headroom before the next prune.
            to_delete = count - int(self.MAX_HISTORY_ENTRIES * 0.8)
            cursor.execute(
                "DELETE FROM changes WHERE id IN (SELECT id FROM changes ORDER BY timestamp ASC LIMIT ?)",
                (to_delete,),
            )
            self._conn.commit()
            # Lightweight index/statistics refresh without long VACUUM blocking.
            self._conn.execute("PRAGMA optimize")

    def prune_old_history(self, days_to_keep: int = None):
        """
        Delete history entries older than specified days.

        Args:
            days_to_keep: Number of days to keep (default: DEFAULT_RETENTION_DAYS)
        """
        # days_to_keep: Retention window in days. Entries with a timestamp older than
        #   (now - days_to_keep) are permanently deleted.
        if days_to_keep is None:
            days_to_keep = self.DEFAULT_RETENTION_DAYS

        with self._lock:
            # cursor: Used to execute the DELETE statement.
            cursor = self._conn.cursor()
            # cutoff_date: ISO-formatted timestamp representing the oldest acceptable entry.
            cutoff_date = (datetime.now() - timedelta(days=days_to_keep)).isoformat()
            cursor.execute("DELETE FROM changes WHERE timestamp < ?", (cutoff_date,))
            # deleted: Number of rows actually removed by the DELETE statement.
            deleted = cursor.rowcount
            self._conn.commit()
            if deleted > 0:
                # Run PRAGMA optimize to refresh index statistics after deletion.
                self._conn.execute("PRAGMA optimize")
            return deleted

    def get_history(
        self, limit: int = 100, success_only: bool = False
    ) -> List[Tuple[Any, ...]]:
        """
        Retrieve recent changes, sorted by most recent first.

        Args:
            limit: Maximum number of entries to return
            success_only: Only return successful changes (filters out failed writes)

        Returns:
            List of history entries, each as (id, timestamp, setting_name, old_value, new_value, success)
        """
        # limit: Caps the result set to avoid loading thousands of rows into memory.
        # success_only: When True, filters the query to only rows where success=1.
        with self._lock:
            # cursor: Used to SELECT from the changes table.
            cursor = self._conn.cursor()
            if success_only:
                cursor.execute(
                    "SELECT id, timestamp, setting_name, old_value, new_value, success "
                    "FROM changes WHERE success = 1 ORDER BY timestamp DESC LIMIT ?",
                    (limit,),
                )
            else:
                cursor.execute(
                    "SELECT id, timestamp, setting_name, old_value, new_value, success "
                    "FROM changes ORDER BY timestamp DESC LIMIT ?",
                    (limit,),
                )
            # Return all matching rows as a list of tuples.
            return cursor.fetchall()

    def get_history_by_setting(
        self, setting_id: str, limit: int = 50
    ) -> List[Tuple[Any, ...]]:
        """
        Retrieve history for a specific setting.

        Args:
            setting_id: The ID of the setting to look up (truncated to 100 chars for safety)
            limit: Maximum number of entries to return

        Returns:
            List of history entries for that setting, each as (id, timestamp, old_value, new_value, success)
        """
        # setting_id: The unique WinSet setting identifier used to filter rows.
        # limit: Caps the result set.
        with self._lock:
            # cursor: Used to execute the filtered SELECT.
            cursor = self._conn.cursor()
            cursor.execute(
                "SELECT id, timestamp, old_value, new_value, success "
                "FROM changes WHERE setting_id = ? ORDER BY timestamp DESC LIMIT ?",
                (setting_id[:100], limit),
            )
            return cursor.fetchall()

    def get_change_details(self, change_id: int) -> Optional[Tuple[Any, ...]]:
        """
        Get the details needed to revert a specific change by its ID.

        Args:
            change_id: The ID (primary key) of the change row to retrieve

        Returns:
            Tuple of (hive, key_path, value_name, value_type, old_value) or None if not found
        """
        # change_id: The auto-incremented primary key of the change record.
        with self._lock:
            # cursor: Used to SELECT the columns required for reverting a Registry change.
            cursor = self._conn.cursor()
            cursor.execute(
                "SELECT hive, key_path, value_name, value_type, old_value FROM changes WHERE id = ?",
                (change_id,),
            )
            # Returns a single tuple or None if no row matches the given ID.
            return cursor.fetchone()

    def get_statistics(self) -> Dict[str, Any]:
        """
        Get statistics about the history database.

        Returns:
            Dictionary with keys: total_entries, successful_changes, failed_changes,
            unique_settings, first_entry, last_entry, database_size_bytes, database_size_mb
        """
        with self._lock:
            # cursor: Used to run aggregate queries for statistics.
            cursor = self._conn.cursor()

            # total: Count of all rows in the changes table.
            cursor.execute("SELECT COUNT(*) FROM changes")
            total = cursor.fetchone()[0]

            # successful: Count of rows where the Registry write succeeded.
            cursor.execute("SELECT COUNT(*) FROM changes WHERE success = 1")
            successful = cursor.fetchone()[0]

            # unique_settings: Number of distinct setting_id values ever recorded.
            cursor.execute("SELECT COUNT(DISTINCT setting_id) FROM changes")
            unique_settings = cursor.fetchone()[0]

            # first / last: Timestamps of the earliest and most recent history entries.
            cursor.execute("SELECT MIN(timestamp), MAX(timestamp) FROM changes")
            first, last = cursor.fetchone()

        # db_size: Size of the SQLite database file in bytes (0 if the file does not exist).
        db_size = os.path.getsize(self.db_path) if os.path.exists(self.db_path) else 0

        # Return a dictionary summarizing the database contents and file size.
        return {
            "total_entries": total,
            "successful_changes": successful,
            "failed_changes": total - successful,
            "unique_settings": unique_settings,
            "first_entry": first,
            "last_entry": last,
            "database_size_bytes": db_size,
            "database_size_mb": round(db_size / (1024 * 1024), 2),
        }

    def export_history(self, filepath: str, format: str = "json") -> bool:
        """
        Export history to a file for backup or analysis.

        Args:
            filepath: Destination file path for the export
            format: Export format -- 'json' or 'csv'

        Returns:
            True if export succeeded, False otherwise
        """
        # csv / json: Imported locally to avoid loading them at module level when unused.
        import csv
        import json

        with self._lock:
            # cursor: Used to fetch all rows for export.
            cursor = self._conn.cursor()
            cursor.execute(
                "SELECT id, timestamp, setting_name, old_value, new_value, success FROM changes "
                "ORDER BY timestamp DESC"
            )
            # rows: Complete list of history records to be exported.
            rows = cursor.fetchall()

        try:
            if format == "json":
                # data: List of dictionaries, one per history row, suitable for JSON serialization.
                data = [
                    {
                        "id": row[0],
                        "timestamp": row[1],
                        "setting_name": row[2],
                        "old_value": row[3],
                        "new_value": row[4],
                        "success": bool(row[5]),
                    }
                    for row in rows
                ]
                with open(filepath, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
            elif format == "csv":
                with open(filepath, "w", newline="", encoding="utf-8") as f:
                    # writer: CSV writer that outputs a header row followed by one row per history entry.
                    writer = csv.writer(f)
                    writer.writerow(
                        [
                            "ID",
                            "Timestamp",
                            "Setting",
                            "Old Value",
                            "New Value",
                            "Success",
                        ]
                    )
                    writer.writerows(rows)
            else:
                # Unsupported format requested.
                return False
            return True
        except Exception as e:
            # Print the error and return False so callers can detect the failure.
            print(f"Export failed: {e}")
            return False

        # src/history/history_manager.py

    def clear_history(self) -> bool:
        """
        Clear all history.

        Returns:
            True if clear was successful, False otherwise
        """
        try:
            with self._lock:
                # cursor: Used to delete every row from the changes table.
                cursor = self._conn.cursor()
                cursor.execute("DELETE FROM changes")
                self._conn.commit()
            return True
        except Exception:
            # Return False on any database error without raising.
            return False

    def revert_change(self, change_id: int) -> bool:
        """
        Revert a specific change by ID.

        Args:
            change_id: The ID of the change row to revert

        Returns:
            True if revert was successful, False otherwise
        """
        try:
            # Fetch the hive, key_path, value_name, value_type, and old_value needed to
            # reconstruct the previous Registry state.
            details = self.get_change_details(change_id)
            if not details:
                return False

            # details: Tuple of (hive, key_path, value_name, value_type, old_value)
            hive, key_path, value_name, value_type, old_value = details

            # Import RegistryHandler here instead of at module level to avoid a circular
            # import between history_manager and registry_handler.
            from .registry_handler import RegistryHandler

            # handler: RegistryHandler instance used to write the old value back to the Windows Registry.
            handler = RegistryHandler()

            # converted_value: The old_value string converted back to its native Python type
            #   (int for DWORD/QWORD, str for SZ types, etc.) so the Registry API accepts it.
            converted_value = self._convert_string_to_value(old_value, value_type)

            # Write the original value back to the Registry to undo the change.
            success = handler.write_value(
                hive, key_path, value_name, value_type, converted_value
            )

            if success:
                with self._lock:
                    # cursor: Used to mark the change row as reverted so it is not reverted again.
                    cursor = self._conn.cursor()
                    cursor.execute(
                        "UPDATE changes SET reverted = 1 WHERE id = ?", (change_id,)
                    )
                    self._conn.commit()

            return success
        except Exception:
            # Swallow all exceptions and return False so callers get a clean boolean result.
            return False

    def _convert_string_to_value(self, value_str: str, value_type: str) -> Any:
        """Convert string value back to proper type."""
        # value_str: The string stored in the database (e.g., "42" or "N/A").
        # value_type: The Registry value type string (e.g., "REG_DWORD", "REG_SZ").
        if value_str == "N/A":
            # "N/A" is the sentinel we use for None values; return None so the Registry
            # handler knows there was no previous value.
            return None

        try:
            if value_type == "REG_DWORD":
                # DWORD is a 32-bit integer in the Windows Registry.
                return int(value_str)
            elif value_type == "REG_QWORD":
                # QWORD is a 64-bit integer in the Windows Registry.
                return int(value_str)
            elif value_type in ["REG_SZ", "REG_EXPAND_SZ"]:
                # REG_SZ and REG_EXPAND_SZ are string types; no conversion needed.
                return value_str
            else:
                # For any other type (REG_BINARY, REG_MULTI_SZ, etc.), return the raw string.
                return value_str
        except (ValueError, TypeError):
            # If conversion fails, return the raw string and let the handler deal with it.
            return value_str

    def get_changes_by_setting(self, setting_id: str) -> List[Tuple[Any, ...]]:
        """
        Get all changes for a specific setting.

        Args:
            setting_id: The setting ID to filter by (truncated to 100 chars)

        Returns:
            List of changes, each as (id, timestamp, old_value, new_value, reverted)
        """
        # setting_id: Unique WinSet identifier used to filter the changes table.
        with self._lock:
            # cursor: Used to SELECT changes matching the given setting_id.
            cursor = self._conn.cursor()
            cursor.execute(
                "SELECT id, timestamp, old_value, new_value, reverted "
                "FROM changes WHERE setting_id = ? ORDER BY timestamp DESC",
                (setting_id[:100],),
            )
            return cursor.fetchall()

    def close(self):
        """Close the database connection."""
        with self._lock:
            # Guard against calling close() more than once.
            if self._conn:
                # Close the underlying SQLite connection and release file handles.
                self._conn.close()
                # Set to None so subsequent calls to close() are safe no-ops.
                self._conn = None
