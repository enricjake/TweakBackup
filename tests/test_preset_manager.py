"""
Tests for preset manager
"""

import pytest
import sys
import os
import tempfile
import json
import shutil

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from presets.preset_manager import PresetManager
from models.setting import RegistrySetting, SettingCategory, SettingType


class TestPresetManager:
    """Test PresetManager class"""
    
    def setup_method(self):
        """Setup test environment"""
        # Create temporary directories
        self.test_dir = tempfile.mkdtemp()
        self.presets_dir = os.path.join(self.test_dir, "presets")
        os.makedirs(self.presets_dir, exist_ok=True)
        
        # Create a test preset file
        test_preset = {
            "name": "Test Preset",
            "description": "A test preset",
            "version": "1.0",
            "settings": {
                "test_setting": {
                    "value": 1,
                    "description": "A test setting"
                }
            }
        }
        
        with open(os.path.join(self.presets_dir, "test.json"), 'w') as f:
            json.dump(test_preset, f, indent=2)
        
        # Initialize preset manager
        self.manager = PresetManager(presets_dir=self.presets_dir)
    
    def teardown_method(self):
        """Cleanup test environment"""
        shutil.rmtree(self.test_dir)
    
    def test_initialization(self):
        """Test preset manager initialization"""
        assert self.manager is not None
        assert len(self.manager.get_preset_list()) > 0
    
    def test_get_preset_list(self):
        """Test getting list of available presets"""
        presets = self.manager.get_preset_list()
        assert "test" in presets
    
    def test_get_preset_info(self):
        """Test getting preset information"""
        info = self.manager.get_preset_info("test")
        assert info is not None
        assert info["name"] == "Test Preset"
        assert info["description"] == "A test preset"
        assert info["is_custom"] == False
    
    def test_load_preset(self):
        """Test loading a preset"""
        success, msg, profile = self.manager.load_preset("test")
        assert success == True
        assert profile is not None
        assert profile.name == "Test Preset"
    
    def test_create_custom_preset(self):
        """Test creating a custom preset"""
        # Create a setting for the preset
        setting = RegistrySetting(
            id="custom_setting",
            name="Custom Setting",
            description="A custom setting",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=42,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Custom",
            value_name="CustomValue",
            value_type="REG_DWORD"
        )
        
        # Create custom preset
        success, msg, preset_id = self.manager.create_custom_preset(
            name="My Custom Preset",
            description="A preset I created",
            settings=[setting]
        )
        
        assert success == True
        assert preset_id == "custom_my_custom_preset"
        assert os.path.exists(os.path.join(self.manager.custom_presets_dir, "my_custom_preset.json"))
        
        # Verify preset is in the list
        presets = self.manager.get_preset_list()
        assert "custom_my_custom_preset" in presets
    
    def test_delete_custom_preset(self):
        """Test deleting a custom preset"""
        # First create a custom preset
        setting = RegistrySetting(
            id="temp_setting",
            name="Temp Setting",
            description="A temporary setting",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=1,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Temp",
            value_name="TempValue",
            value_type="REG_DWORD"
        )
        
        self.manager.create_custom_preset(
            name="Temp Preset",
            settings=[setting]
        )
        
        # Verify it exists
        presets = self.manager.get_preset_list()
        assert "custom_temp_preset" in presets
        
        # Delete it
        success, msg = self.manager.delete_custom_preset("custom_temp_preset")
        assert success == True
        
        # Verify it's gone
        presets = self.manager.get_preset_list()
        assert "custom_temp_preset" not in presets
    
    def test_export_import_preset(self):
        """Test exporting and importing presets"""
        # Create a custom preset first
        setting = RegistrySetting(
            id="export_setting",
            name="Export Setting",
            description="A setting for export",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=99,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Export",
            value_name="ExportValue",
            value_type="REG_DWORD"
        )
        
        self.manager.create_custom_preset(
            name="Exportable Preset",
            settings=[setting]
        )
        
        # Export the preset
        export_path = os.path.join(self.test_dir, "exported_preset.json")
        success, msg = self.manager.export_preset("custom_exportable_preset", export_path)
        assert success == True
        assert os.path.exists(export_path)
        
        # Import it back with a different name
        # Create a new preset file with the desired name and minimal structure
        imported_preset_data = {
            "name": "Imported Preset",
            "description": "An imported preset",
            "version": "1.0",
            "settings": {
                "export_setting": {
                    "value": 99,
                    "description": "A setting for export"
                }
            }
        }
        
        with open(export_path, 'w') as f:
            json.dump(imported_preset_data, f, indent=2)
        
        # Now import it
        success, msg, new_preset_id = self.manager.import_preset(export_path, overwrite=True)
        assert success == True
        # The preset ID will be based on the new name "Imported Preset" -> "imported_preset"
        assert new_preset_id == "custom_imported_preset"
        
        # Refresh presets to include the newly imported one
        self.manager.refresh_presets()
        
        # Verify the imported preset exists
        presets = self.manager.get_preset_list()
        assert "custom_imported_preset" in presets
    
    def test_cannot_delete_builtin_preset(self):
        """Test that built-in presets cannot be deleted"""
        # Try to delete a built-in preset
        success, msg = self.manager.delete_custom_preset("test")
        assert success == False
        assert "Cannot delete built-in presets" in msg
    
    def test_cannot_update_builtin_preset(self):
        """Test that built-in presets cannot be updated"""
        # Try to update a built-in preset
        success, msg = self.manager.update_custom_preset(
            preset_id="test",
            name="New Name"
        )
        assert success == False
        assert "Cannot update built-in presets" in msg
