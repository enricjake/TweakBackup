"""
Tests for WinSet model classes.
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from models.setting import RegistrySetting, SettingCategory, SettingType


class TestModels:
    """Tests for model classes"""

    def test_registry_setting_creation(self):
        """Test creating a registry setting"""
        setting = RegistrySetting(
            id="reg1",
            name="Registry Test",
            description="Test registry setting",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=1,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Microsoft\\Windows\\CurrentVersion",
            value_name="TestValue",
            value_type="REG_DWORD"
        )

        assert setting.hive == "HKEY_CURRENT_USER"
        assert setting.key_path == "Software\\Microsoft\\Windows\\CurrentVersion"
        assert setting.value_name == "TestValue"
        assert setting.value_type == "REG_DWORD"
        assert setting.is_expanded is False  # Default value

    def test_registry_setting_validate_dword(self):
        """Test DWORD validation"""
        setting = RegistrySetting(
            id="reg2",
            name="DWORD Test",
            description="",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value=0,
            default_value=0,
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Test",
            value_name="Val",
            value_type="REG_DWORD"
        )
        assert setting.validate(0) is True
        assert setting.validate(4294967295) is True
        assert setting.validate(-1) is False
        assert setting.validate("not_an_int") is False

    def test_registry_setting_validate_sz(self):
        """Test REG_SZ validation"""
        setting = RegistrySetting(
            id="reg3",
            name="SZ Test",
            description="",
            category=SettingCategory.SYSTEM,
            setting_type=SettingType.REGISTRY,
            value="",
            default_value="",
            hive="HKEY_CURRENT_USER",
            key_path="Software\\Test",
            value_name="Val",
            value_type="REG_SZ"
        )
        assert setting.validate("hello") is True
        assert setting.validate(123) is False

    def test_registry_setting_export(self):
        """Test that export() returns the expected keys"""
        setting = RegistrySetting(
            id="reg4",
            name="Export Test",
            description="desc",
            category=SettingCategory.APPEARANCE,
            setting_type=SettingType.REGISTRY,
            value=1,
            default_value=0,
            hive="HKEY_LOCAL_MACHINE",
            key_path="Software\\Test",
            value_name="ExportVal",
            value_type="REG_DWORD"
        )
        exported = setting.export()
        assert exported["hive"] == "HKEY_LOCAL_MACHINE"
        assert exported["value_name"] == "ExportVal"
        assert exported["value_type"] == "REG_DWORD"