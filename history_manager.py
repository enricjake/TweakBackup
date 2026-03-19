"""
HistoryManager for WinSet - logs all setting changes to a persistent database.
"""

import sqlite3
import os
import logging
from datetime import datetime
from typing import List, Tuple, Any, Optional, Dict

# Configure logging
logger = logging.getLogger(__name__)


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
        logger.info(f"History database initialized at: {self.db_path}")

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
                value_name TEXT NOT NULL,
                reverted INTEGER DEFAULT 0
            )
        """)
        self._conn.commit()
        logger.debug("History table verified/created")

    def log_change(self, setting: Any, old_value: Any, new_value: Any):
        """Log a single setting change to the database."""
        cursor = self._conn.cursor()
        old_val_str = str(old_value) if old_value is not None else 'N/A'
        new_val_str = str(new_value) if new_value is not None else 'N/A'

        cursor.execute("""
            INSERT INTO changes (timestamp, setting_id, setting_name, old_value, new_value, value_type, hive, key_path, value_name, reverted)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.now().isoformat(sep=' ', timespec='seconds'),
            setting.id, setting.name, old_val_str, new_val_str,
            setting.value_type, setting.hive, setting.key_path, setting.value_name,
            0  # Not reverted
        ))
        self._conn.commit()
        logger.info(f"Logged change: {setting.name} ({setting.id}) from {old_val_str} to {new_val_str}")

    def get_history(
        self, 
        limit: Optional[int] = None,
        filter_reverted: bool = False
    ) -> List[Tuple[Any, ...]]:
        """Retrieve all changes, sorted by most recent first.
        
        Args:
            limit: Maximum number of records to return.
            filter_reverted: If True, exclude reverted changes.
        
        Returns:
            List of change records.
        """
        cursor = self._conn.cursor()
        query = "SELECT id, timestamp, setting_name, old_value, new_value, reverted FROM changes"
        
        if filter_reverted:
            query += " WHERE reverted = 0"
            
        query += " ORDER BY timestamp DESC"
        
        if limit:
            query += f" LIMIT {limit}"
            
        cursor.execute(query)
        results = cursor.fetchall()
        logger.debug(f"Retrieved {len(results)} history records")
        return results

    def get_change_details(self, change_id: int) -> Optional[Tuple[Any, ...]]:
        """Get the details needed to revert a specific change by its ID.
        
        Args:
            change_id: The ID of the change to retrieve.
            
        Returns:
            Tuple containing (hive, key_path, value_name, value_type, old_value) or None if not found.
        """
        cursor = self._conn.cursor()
        cursor.execute(
            "SELECT hive, key_path, value_name, value_type, old_value FROM changes WHERE id = ?", 
            (change_id,)
        )
        result = cursor.fetchone()
        if result:
            logger.debug(f"Retrieved details for change ID {change_id}")
        else:
            logger.warning(f"Change ID {change_id} not found")
        return result

    def revert_change(self, change_id: int) -> bool:
        """Revert a specific change by restoring the old value.
        
        Args:
            change_id: The ID of the change to revert.
            
        Returns:
            True if successfully reverted, False otherwise.
        """
        details = self.get_change_details(change_id)
        if not details:
            logger.error(f"Cannot revert change {change_id}: details not found")
            return False
        
        try:
            hive, key_path, value_name, value_type, old_value_str = details
            
            # Import here to avoid circular dependencies
            from src.core.registry_handler import RegistryHandler
            handler = RegistryHandler()
            
            # Convert old value back to appropriate type
            old_value = self._convert_value(old_value_str, value_type)
            
            # Write the old value back to the registry
            success = handler.write_value(
                hive=hive,
                key_path=key_path,
                value_name=value_name,
                value_type=value_type,
                value=old_value
            )
            
            if success:
                # Mark the change as reverted in the database
                cursor = self._conn.cursor()
                cursor.execute(
                    "UPDATE changes SET reverted = 1 WHERE id = ?",
                    (change_id,)
                )
                self._conn.commit()
                logger.info(f"Successfully reverted change {change_id}: {value_name}")
            else:
                logger.error(f"Failed to revert change {change_id}: registry write failed")
                
            return success
            
        except Exception as e:
            logger.error(f"Error reverting change {change_id}: {e}")
            return False

    def _convert_value(self, value_str: str, value_type: str) -> Any:
        """Convert string value back to appropriate type for registry write.
        
        Args:
            value_str: String representation of the value.
            value_type: Registry value type.
            
        Returns:
            Value in appropriate type.
        """
        if value_str == 'N/A':
            return None
            
        try:
            if value_type == "REG_DWORD":
                return int(value_str)
            elif value_type == "REG_QWORD":
                return int(value_str)
            elif value_type in ["REG_SZ", "REG_EXPAND_SZ"]:
                return str(value_str)
            elif value_type == "REG_BINARY":
                # Simple hex string conversion for binary data
                if value_str.startswith('0x'):
                    return bytes.fromhex(value_str[2:])
                return value_str.encode('utf-8')
            elif value_type == "REG_MULTI_SZ":
                # Parse multi-string (simplified)
                return value_str.split('|') if '|' in value_str else [value_str]
            else:
                return value_str
        except Exception as e:
            logger.warning(f"Failed to convert value '{value_str}' to type {value_type}: {e}")
            return value_str

    def revert_all_changes(self, filter_reverted: bool = True) -> Dict[int, bool]:
        """Revert all changes in the history.
        
        Args:
            filter_reverted: If True, only revert non-reverted changes.
            
        Returns:
            Dictionary mapping change IDs to success status.
        """
        results = {}
        changes = self.get_history(filter_reverted=filter_reverted)
        
        for change in changes:
            change_id = change[0]
            results[change_id] = self.revert_change(change_id)
        
        logger.info(f"Reverted {sum(results.values())}/{len(results)} changes")
        return results

    def get_changes_by_setting(self, setting_id: str) -> List[Tuple[Any, ...]]:
        """Get all changes for a specific setting.
        
        Args:
            setting_id: The ID of the setting to filter by.
            
        Returns:
            List of change records for the specified setting.
        """
        cursor = self._conn.cursor()
        cursor.execute(
            "SELECT id, timestamp, old_value, new_value, reverted FROM changes WHERE setting_id = ? ORDER BY timestamp DESC",
            (setting_id,)
        )
        results = cursor.fetchall()
        logger.debug(f"Found {len(results)} changes for setting {setting_id}")
        return results

    def clear_history(self) -> bool:
        """Clear all history records.
        
        Returns:
            True if successfully cleared, False otherwise.
        """
        try:
            cursor = self._conn.cursor()
            cursor.execute("DELETE FROM changes")
            self._conn.commit()
            logger.info("Cleared all history records")
            return True
        except Exception as e:
            logger.error(f"Failed to clear history: {e}")
            return False

    def export_history(self, file_path: str) -> bool:
        """Export history to a JSON file.
        
        Args:
            file_path: Path to the output file.
            
        Returns:
            True if successfully exported, False otherwise.
        """
        try:
            import json
            from collections import OrderedDict
            
            cursor = self._conn.cursor()
            cursor.execute("SELECT * FROM changes ORDER BY timestamp DESC")
            rows = cursor.fetchall()
            
            # Get column names
            column_names = [description[0] for description in cursor.description]
            
            # Convert to list of dictionaries
            history_data = []
            for row in rows:
                history_data.append(OrderedDict(zip(column_names, row)))
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
            
            # Write to file
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(history_data, f, indent=2, default=str)
            
            logger.info(f"Exported history to {file_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to export history: {e}")
            return False

    def close(self):
        """Close the database connection."""
        self._conn.close()
        logger.debug("History database connection closed")
