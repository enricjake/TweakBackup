"""
Tests for preset loading and validation.
"""

import pytest
import sys
import os
import json
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from src.presets.preset_manager import PresetManager


class TestPresetLoading:
    """Test preset loading functionality."""
    
    def test_load_valid_preset(self):
        """Test loading a valid preset file."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create a valid preset file
            preset_data = {
                "app": "WinSet",
                "name": "Valid Preset",
                "description": "This is a valid preset",
                "icon": "✅",
                "settings": {
                    "setting1": 1,
                    "setting2": "value"
                }
            }
            
            preset_path = os.path.join(temp_dir, "valid.preset.json")
            with open(preset_path, 'w') as f:
                json.dump(preset_data, f)
            
            manager = PresetManager(presets_dir=temp_dir)
            info = manager.get_preset_info("valid")
            
            assert info is not None
            assert info["name"] == "Valid Preset"
            assert info["setting_count"] == 2
    
    def test_load_invalid_preset_skipped(self):
        """Test that invalid preset files are skipped."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create invalid preset files
            invalid_files = [
                ("missing_name.preset.json", {"app": "WinSet", "description": "No name", "settings": {}}),
                ("missing_settings.preset.json", {"app": "WinSet", "name": "No Settings"}),
                ("wrong_type.preset.json", {"app": "WinSet", "name": "Wrong", "settings": "not a dict"}),
                ("missing_app.preset.json", {"name": "No App", "description": "Missing signature", "settings": {}}),
                ("empty.preset.json", {}),
            ]
            
            for filename, data in invalid_files:
                filepath = os.path.join(temp_dir, filename)
                with open(filepath, 'w') as f:
                    json.dump(data, f)
            
            # Also create a valid one to ensure loading works
            valid_data = {
                "app": "WinSet",
                "name": "Valid",
                "description": "Valid",
                "settings": {"a": 1}
            }
            with open(os.path.join(temp_dir, "valid.preset.json"), 'w') as f:
                json.dump(valid_data, f)
            
            manager = PresetManager(presets_dir=temp_dir)
            
            # Only valid preset should be loaded
            assert manager.get_preset_info("valid") is not None
            for filename, _ in invalid_files:
                preset_id = filename[:-12]
                assert manager.get_preset_info(preset_id) is None
    
    def test_load_malformed_json_skipped(self):
        """Test that malformed JSON files are skipped."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create a malformed JSON file
            filepath = os.path.join(temp_dir, "malformed.preset.json")
            with open(filepath, 'w') as f:
                f.write('{"app": "WinSet", "name": "test", "description": "malformed", "settings": {')
            
            manager = PresetManager(presets_dir=temp_dir)
            
            # Should not load
            assert manager.get_preset_info("malformed") is None
    
    def test_preset_with_too_many_settings_rejected(self):
        """Test that presets with too many settings are rejected."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create a preset with 300 settings (exceeds limit of 200)
            many_settings = {f"setting_{i}": i for i in range(300)}
            preset_data = {
                "app": "WinSet",
                "name": "Too Many",
                "description": "Has too many settings",
                "settings": many_settings
            }
            
            filepath = os.path.join(temp_dir, "toomany.preset.json")
            with open(filepath, 'w') as f:
                json.dump(preset_data, f)
            
            manager = PresetManager(presets_dir=temp_dir)
            
            # Should be rejected due to validation
            assert manager.get_preset_info("toomany") is None
    
    def test_preset_injection_attempt_blocked(self):
        """Test that injection attempts in preset data are blocked."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create preset with malicious setting ID
            malicious_data = {
                "app": "WinSet",
                "name": "Malicious",
                "description": "Attempts injection",
                "settings": {
                    "'; DROP TABLE settings; --": "malicious",
                    "../../../Windows/System32/config/SAM": "malicious"
                }
            }
            
            filepath = os.path.join(temp_dir, "malicious.preset.json")
            with open(filepath, 'w') as f:
                json.dump(malicious_data, f)
            
            manager = PresetManager(presets_dir=temp_dir)
            
            # Should either be rejected or sanitized
            info = manager.get_preset_info("malicious")
            if info is not None:
                settings = manager.get_preset_settings("malicious")
                # Setting IDs should be safe
                for setting_id in settings.keys():
                    assert ";" not in setting_id
                    assert "../" not in setting_id