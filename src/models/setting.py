"""
Setting models for WinSet
"""

from dataclasses import dataclass, field, KW_ONLY
from typing import Any, Optional, Dict
from enum import Enum
from datetime import datetime


class SettingType(Enum):
    """Type of setting"""
    REGISTRY = "registry"
    FILE = "file"
    SYSTEM = "system"
    POWER = "power"


class SettingCategory(Enum):
    """Category of setting"""
    APPEARANCE = "appearance"
    FILE_EXPLORER = "file_explorer"
    TASKBAR = "taskbar"
    POWER = "power"
    PRIVACY = "privacy"
    SYSTEM = "system"
    KEYBOARD = "keyboard"
    MOUSE = "mouse"
    NETWORK = "network"


@dataclass
class Setting:
    """Base class for all settings"""
    id: str
    name: str
    description: str
    category: SettingCategory
    setting_type: SettingType
    value: Any
    default_value: Any
    is_applied: bool = False
    requires_admin: bool = False
    requires_restart: bool = False
    last_modified: Optional[datetime] = None

    def validate(self, value: Any) -> bool:
        """Validate if a value is acceptable"""
        raise NotImplementedError(f"Validate not implemented for {self.__class__.__name__}")

    def apply(self) -> bool:
        """Apply this setting to the system"""
        raise NotImplementedError(f"Apply not implemented for {self.__class__.__name__}")

    def export(self) -> dict:
        """Convert to JSON-serializable dict"""
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "value": self.value,
            "category": self.category.value,
            "type": self.setting_type.value,
            "requires_admin": self.requires_admin,
            "requires_restart": self.requires_restart
        }


@dataclass
class RegistrySetting(Setting):
    """Setting stored in Windows Registry"""
    # KW_ONLY sentinel: all fields below this are keyword-only.
    # This is required because the base Setting class has optional fields
    # (e.g. last_modified) that would otherwise conflict with these
    # required (no-default) fields in the subclass.
    _: KW_ONLY
    hive: str  # HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, etc.
    key_path: str
    value_name: str
    value_type: str  # REG_SZ, REG_DWORD, etc.
    # Optional fields with defaults
    is_expanded: bool = False  # For REG_EXPAND_SZ

    def __post_init__(self):
        """Validate after initialization"""
        # Set the setting_type automatically if not already set
        if self.setting_type == SettingType.REGISTRY:
            pass  # Already correct
        else:
            # Force to registry type
            object.__setattr__(self, 'setting_type', SettingType.REGISTRY)

    def validate(self, value: Any) -> bool:
        """Validate based on registry type"""
        if self.value_type == "REG_DWORD":
            return isinstance(value, int) and 0 <= value <= 4294967295
        elif self.value_type == "REG_SZ" or self.value_type == "REG_EXPAND_SZ":
            return isinstance(value, str)
        elif self.value_type == "REG_BINARY":
            return isinstance(value, (bytes, bytearray))
        elif self.value_type == "REG_MULTI_SZ":
            return isinstance(value, list) and all(isinstance(s, str) for s in value)
        elif self.value_type == "REG_QWORD":
            return isinstance(value, int) and 0 <= value <= 18446744073709551615
        return True

    def apply(self) -> bool:
        """Apply to registry"""
        try:
            import winreg
            
            # Map hive string to winreg constant
            hive_map = {
                "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
                "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
                "HKEY_CLASSES_ROOT": winreg.HKEY_CLASSES_ROOT,
                "HKEY_CURRENT_CONFIG": winreg.HKEY_CURRENT_CONFIG,
                "HKEY_USERS": winreg.HKEY_USERS
            }
            
            # Map type string to winreg constant
            type_map = {
                "REG_SZ": winreg.REG_SZ,
                "REG_DWORD": winreg.REG_DWORD,
                "REG_BINARY": winreg.REG_BINARY,
                "REG_MULTI_SZ": winreg.REG_MULTI_SZ,
                "REG_EXPAND_SZ": winreg.REG_EXPAND_SZ,
                "REG_QWORD": winreg.REG_QWORD
            }
            
            hive_constant = hive_map.get(self.hive)
            type_constant = type_map.get(self.value_type)
            
            if not hive_constant or not type_constant:
                raise ValueError(f"Invalid hive or type: {self.hive}, {self.value_type}")
            
            # Open key with write access
            key = winreg.OpenKey(
                hive_constant,
                self.key_path,
                0,
                winreg.KEY_SET_VALUE
            )
            
            # Set value
            winreg.SetValueEx(key, self.value_name, 0, type_constant, self.value)
            winreg.CloseKey(key)
            
            self.is_applied = True
            return True
            
        except Exception as e:
            print(f"Failed to apply {self.name}: {e}")
            return False

    def export(self) -> dict:
        """Export registry setting to dict"""
        base_dict = super().export()
        base_dict.update({
            "hive": self.hive,
            "key_path": self.key_path,
            "value_name": self.value_name,
            "value_type": self.value_type,
            "is_expanded": self.is_expanded
        })
        return base_dict