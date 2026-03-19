"""
Tests for history manager
"""

import pytest
import sys
import os
import tempfile
import sqlite3

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from history_manager import HistoryManager
from src.models.setting import RegistrySetting, SettingCategory, SettingType


class TestHistoryManager:
    """Test HistoryManager class"""
    
    def setup_method(self):
        """Setup test database"""
        # Create a temporary database for testing
        self.db_file = tempfile.NamedTemporaryFile(delete=False, suffix='.db')
        self.db_file.close()
        self.history = HistoryManager(db_path=self.db_file.name)
    
    def teardown_method(self):
        """Cleanup test database"""
        self.history.close()
        os.unlink(self.db_file.name)
    
    def test_initialization(self):
        """Test history manager initialization"""
        assert self.history is not None
        assert os.path.exists(self.db_file.name)
    
    def test_log_and_retrieve_change(self):
        """Test logging and retrieving a change"""
        # Create a mock setting
        setting = RegistrySetting(
            id="test_setting",
            name="Test Setting",
            description="A test setting",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=1,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Test",
            value_name="TestValue",
            value_type="REG_DWORD"
        )
        
        # Log a change
        self.history.log_change(setting, old_value=0, new_value=1)
        
        # Retrieve history
        history = self.history.get_history()
        assert len(history) == 1
        
        change = history[0]
        assert change[2] == "Test Setting"  # setting_name
        assert change[3] == "0"  # old_value
        assert change[4] == "1"  # new_value
    
    def test_get_change_details(self):
        """Test retrieving change details"""
        setting = RegistrySetting(
            id="test_setting",
            name="Test Setting",
            description="A test setting",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=1,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Test",
            value_name="TestValue",
            value_type="REG_DWORD"
        )
        
        self.history.log_change(setting, old_value=0, new_value=1)
        
        # Get the change ID from history
        history = self.history.get_history()
        change_id = history[0][0]
        
        # Get details
        details = self.history.get_change_details(change_id)
        assert details is not None
        assert details[0] == "HKEY_CURRENT_USER"
        assert details[2] == "TestValue"
        assert details[4] == "0"  # old_value
    
    def test_revert_change(self, monkeypatch):
        """Test reverting a change"""
        setting = RegistrySetting(
            id="test_setting",
            name="Test Setting",
            description="A test setting",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=1,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Test",
            value_name="TestValue",
            value_type="REG_DWORD"
        )
        
        self.history.log_change(setting, old_value=0, new_value=1)
        
        # Mock registry handler
        def mock_write_value(hive, key_path, value_name, value_type, value):
            return True
        
        # Mock the import
        from unittest.mock import MagicMock
        mock_handler = MagicMock()
        mock_handler.write_value = mock_write_value
        
        monkeypatch.setattr('src.core.registry_handler.RegistryHandler', lambda: mock_handler)
        
        # Get change ID and revert
        history = self.history.get_history()
        change_id = history[0][0]
        
        success = self.history.revert_change(change_id)
        assert success == True
        
        # Check that the change is marked as reverted
        # Use the history manager's connection to avoid file locking issues
        cursor = self.history._conn.cursor()
        cursor.execute("SELECT reverted FROM changes WHERE id = ?", (change_id,))
        result = cursor.fetchone()
        assert result[0] == 1  # reverted flag should be set
    
    def test_get_changes_by_setting(self):
        """Test filtering changes by setting ID"""
        setting1 = RegistrySetting(
            id="setting1",
            name="Setting 1",
            description="First setting",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=1,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Test",
            value_name="Value1",
            value_type="REG_DWORD"
        )
        
        setting2 = RegistrySetting(
            id="setting2",
            name="Setting 2",
            description="Second setting",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=1,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Test",
            value_name="Value2",
            value_type="REG_DWORD"
        )
        
        # Log changes for both settings
        self.history.log_change(setting1, old_value=0, new_value=1)
        self.history.log_change(setting2, old_value=0, new_value=1)
        self.history.log_change(setting1, old_value=1, new_value=2)
        
        # Get changes for setting1 only
        setting1_changes = self.history.get_changes_by_setting("setting1")
        assert len(setting1_changes) == 2
        
        # Check that we have the right changes (old_value, new_value)
        # Note: changes are returned in reverse chronological order
        # The tuple is (id, timestamp, old_value, new_value, reverted)
        # Since both changes have the same timestamp, we need to check for both possibilities
        
        # Collect the actual values
        old_values = [str(change[2]) for change in setting1_changes]
        new_values = [str(change[3]) for change in setting1_changes]
        
        # We should have one change with old=0,new=1 and one with old=1,new=2
        assert "0" in old_values and "1" in old_values
        assert "1" in new_values and "2" in new_values
        assert old_values.count("0") == 1 and old_values.count("1") == 1
        assert new_values.count("1") == 1 and new_values.count("2") == 1
    
    def test_export_history(self):
        """Test exporting history to file"""
        setting = RegistrySetting(
            id="test_setting",
            name="Test Setting",
            description="A test setting",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=1,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Test",
            value_name="TestValue",
            value_type="REG_DWORD"
        )
        
        self.history.log_change(setting, old_value=0, new_value=1)
        
        # Export to temporary file
        export_file = tempfile.NamedTemporaryFile(delete=False, suffix='.json')
        export_file.close()
        
        success = self.history.export_history(export_file.name)
        assert success == True
        assert os.path.exists(export_file.name)
        
        # Cleanup
        os.unlink(export_file.name)
    
    def test_clear_history(self):
        """Test clearing history"""
        setting = RegistrySetting(
            id="test_setting",
            name="Test Setting",
            description="A test setting",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=1,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Test",
            value_name="TestValue",
            value_type="REG_DWORD"
        )
        
        # Add some changes
        self.history.log_change(setting, old_value=0, new_value=1)
        self.history.log_change(setting, old_value=1, new_value=2)
        
        # Verify changes exist
        history = self.history.get_history()
        assert len(history) == 2
        
        # Clear history
        success = self.history.clear_history()
        assert success == True
        
        # Verify history is empty
        history = self.history.get_history()
        assert len(history) == 0
