"""
Setting Loader for WinSet - Loads settings from JSON with security validation.
"""

import json
import os
import re
import sys
from pathlib import Path
from typing import List, Dict, Optional, Any
from src.models.setting import RegistrySetting, SettingCategory, SettingType, Setting


def _safe_print(*args, **kwargs):
    """Print with Unicode-safe encoding."""
    try:
        print(*args, **kwargs)
    except UnicodeEncodeError:
        # Fallback: encode with replacement characters
        safe_args = [str(arg).encode('ascii', 'replace').decode('ascii') for arg in args]
        print(*safe_args, **kwargs)


class SettingLoader:
    """Loads settings from resources/settings.json with security validation"""

    CATEGORY_MAP = {
        "System Appearance": SettingCategory.APPEARANCE,
        "File Explorer Settings": SettingCategory.FILE_EXPLORER,
        "Taskbar & Start Menu": SettingCategory.TASKBAR,
        "Power Settings": SettingCategory.POWER,
        "Privacy Options": SettingCategory.PRIVACY,
        "Privacy Settings": SettingCategory.PRIVACY,
        "Keyboard & Mouse": SettingCategory.KEYBOARD,
        "Mouse & Keyboard Settings": SettingCategory.KEYBOARD,
        "Regional & Language Settings": SettingCategory.SYSTEM,
        "Accessibility Settings": SettingCategory.SYSTEM,
        "Gaming Settings": SettingCategory.SYSTEM,
        "System & Performance": SettingCategory.SYSTEM,
        "Network Settings": SettingCategory.NETWORK,
        "Advanced Settings": SettingCategory.SYSTEM
    }

    # Allowed registry hives for security
    ALLOWED_HIVES = ["HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE"]
    
    # Dangerous registry paths to block
    BLOCKED_KEY_PATTERNS = [
        r'\\\\Security\\\\',
        r'\\\\SAM\\\\',
        r'\\\\System\\\\CurrentControlSet\\\\Control\\\\SecurePipeServers',
        r'\\\\Software\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Policies\\\\System\\\\*',
        r'\\\\.\\.\\\\',  # Path traversal
    ]

    def __init__(self, resource_path: Optional[str] = None):
        """
        Initialize setting loader with optional custom resource path.
        
        Args:
            resource_path: Optional path to settings.json. If not provided,
                          uses the default location within the application.
                          
        Raises:
            ValueError: If the provided resource path is outside the application directory.
        """
        if resource_path is None:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            self.resource_path = os.path.join(base_dir, "resources", "settings.json")
        else:
            # Validate external path to prevent path traversal
            self.resource_path = self._validate_resource_path(resource_path)
        
        self.settings_by_category: Dict[SettingCategory, List[RegistrySetting]] = {}
        self.load_settings()

    def _validate_resource_path(self, path: str) -> str:
        """
        Validate that a resource path is safe to access.
        
        Args:
            path: The path to validate
            
        Returns:
            The resolved safe path
            
        Raises:
            ValueError: If the path is unsafe
        """
        try:
            safe_path = Path(path).resolve()
            base_dir = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))).resolve()
            
            # Check if path is within application directory
            if base_dir not in safe_path.parents and safe_path != base_dir:
                # Also allow user's Documents/WinSet folder for custom profiles
                user_docs = Path(os.path.expanduser("~")) / "Documents" / "WinSet"
                if user_docs.resolve() not in safe_path.parents and safe_path != user_docs.resolve():
                    raise ValueError(
                        f"Resource path '{path}' is outside allowed directories. "
                        f"Allowed: {base_dir} or {user_docs}"
                    )
            
            # Check file extension
            if safe_path.suffix.lower() != '.json':
                raise ValueError(f"Invalid file type. Only .json files are allowed.")
            
            # Check file size (prevent DoS)
            if safe_path.exists():
                size = safe_path.stat().st_size
                if size > 10 * 1024 * 1024:  # 10 MB limit
                    raise ValueError(f"File too large: {size} bytes (max 10MB)")
            
            return str(safe_path)
            
        except Exception as e:
            raise ValueError(f"Invalid resource path: {e}")

    def _validate_setting_data(self, s_data: dict, category_name: str) -> tuple[bool, str]:
        """
        Validate setting data before creating RegistrySetting.
        
        Args:
            s_data: Setting dictionary from JSON
            category_name: Category name for error reporting
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        # Check required fields
        required = ["name", "hive", "key", "value", "type"]
        for field in required:
            if field not in s_data:
                return False, f"Missing required field: {field}"
        
        # Validate field lengths
        name = s_data.get("name", "")
        if len(name) > 200:
            return False, f"Setting name too long: {len(name)} chars (max 200)"
        
        # Validate hive is allowed
        hive = s_data["hive"]
        if hive not in self.ALLOWED_HIVES:
            return False, f"Invalid hive: {hive}. Allowed: {self.ALLOWED_HIVES}"
        
        # Validate key path doesn't contain traversal or dangerous patterns
        key_path = s_data.get("key", "")
        if ".." in key_path:
            return False, f"Invalid key path - contains path traversal: {key_path}"
        
        if key_path.startswith("\\\\"):
            return False, f"Invalid key path - starts with network path: {key_path}"
        
        # Check against blocked patterns
        for pattern in self.BLOCKED_KEY_PATTERNS:
            if re.search(pattern, key_path, re.IGNORECASE):
                return False, f"Blocked registry path pattern: {pattern}"
        
        # Validate key path length
        if len(key_path) > 512:
            return False, f"Key path too long: {len(key_path)} chars (max 512)"
        
        # Validate value name
        value_name = s_data.get("value", "")
        if len(value_name) > 255:
            return False, f"Value name too long: {len(value_name)} chars (max 255)"
        
        # Validate value type
        valid_types = ["REG_SZ", "REG_DWORD", "REG_BINARY", "REG_MULTI_SZ", 
                       "REG_EXPAND_SZ", "REG_QWORD"]
        if s_data["type"] not in valid_types:
            return False, f"Invalid registry type: {s_data['type']}"
        
        # Validate value based on type (if default value provided)
        value = s_data.get("default_value")
        if value is not None:
            if s_data["type"] == "REG_DWORD":
                if not isinstance(value, int) or value < 0 or value > 4294967295:
                    return False, f"Invalid DWORD value: {value}"
            elif s_data["type"] in ["REG_SZ", "REG_EXPAND_SZ"]:
                if not isinstance(value, str) or len(value) > 32767:
                    return False, f"Invalid string value length"
        
        return True, ""

    def _validate_category_data(self, cat_data: dict) -> tuple[bool, str]:
        """
        Validate category data structure.
        
        Args:
            cat_data: Category dictionary from JSON
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        if "name" not in cat_data:
            return False, "Category missing 'name' field"
        
        if "settings" not in cat_data:
            return False, f"Category '{cat_data.get('name')}' missing 'settings' field"
        
        if not isinstance(cat_data["settings"], list):
            return False, f"Category '{cat_data.get('name')}' settings is not a list"
        
        # Limit number of settings per category
        if len(cat_data["settings"]) > 100:
            return False, f"Too many settings in category: {len(cat_data['settings'])} (max 100)"
        
        return True, ""

    def load_settings(self):
        """Load settings from JSON file with comprehensive validation."""
        if not os.path.exists(self.resource_path):
            _safe_print(f"Settings resource not found: {self.resource_path}")
            return

        try:
            with open(self.resource_path, 'r', encoding='utf-8') as f:
                # Read with size limit to prevent DoS
                content = f.read(50 * 1024 * 1024)  # 50 MB max
                data = json.loads(content)
        except json.JSONDecodeError as e:
            _safe_print(f"Invalid JSON in settings file: {e}")
            return
        except Exception as e:
            _safe_print(f"Error reading settings file: {e}")
            return

        if not isinstance(data, list):
            _safe_print("Settings file must contain a list of categories")
            return

        # Limit total categories
        if len(data) > 50:
            _safe_print("Too many categories in settings file")
            return

        for cat_data in data:
            # Validate category structure
            valid, error = self._validate_category_data(cat_data)
            if not valid:
                _safe_print(f"Skipping invalid category: {error}")
                continue

            cat_name = cat_data.get("name")
            enum_cat = self.CATEGORY_MAP.get(cat_name, SettingCategory.SYSTEM)
            
            if enum_cat not in self.settings_by_category:
                self.settings_by_category[enum_cat] = []

            # Limit settings per category
            settings_loaded = 0
            for s_data in cat_data.get("settings", []):
                # Validate setting data
                valid, error = self._validate_setting_data(s_data, cat_name)
                if not valid:
                    _safe_print(f"Skipping setting '{s_data.get('name', 'unknown')}' in category '{cat_name}': {error}")
                    continue

                try:
                    # Create safe setting ID from name
                    safe_id = re.sub(r'[^a-z0-9_]', '_', s_data.get("name", "unknown").lower())
                    
                    setting = RegistrySetting(
                        id=safe_id,
                        name=s_data.get("name", "Unknown Setting")[:200],  # Truncate if needed
                        description=s_data.get("description", "")[:500],   # Truncate description
                        category=enum_cat,
                        setting_type=SettingType.REGISTRY,
                        value=None,  # To be read from registry later
                        default_value=s_data.get("default_value"),
                        hive=s_data["hive"],
                        key_path=s_data.get("key", "").replace("\\\\", "\\")[:512],
                        value_name=s_data.get("value", "")[:255],
                        value_type=s_data["type"],
                        options=None
                    )
                    
                    # Store the values field if present
                    if "values" in s_data:
                        if isinstance(s_data["values"], dict):
                            # Limit size of values dictionary
                            if len(s_data["values"]) <= 50:
                                setting.values = s_data["values"]
                        elif isinstance(s_data["values"], str):
                            # Parse the string format: "0 = Left, 1 = Center"
                            options = {}
                            for pair in s_data["values"].split(','):
                                pair = pair.strip()
                                if '=' in pair:
                                    key, val = pair.split('=', 1)
                                    options[key.strip()] = val.strip()
                            if len(options) <= 50:
                                setting.values = options
                    elif "range" in s_data:
                        if isinstance(s_data["range"], list) and len(s_data["range"]) <= 100:
                            setting.values = s_data["range"]
                            setting.is_range = True
                    
                    # Store option hints if present
                    if "option_hints" in s_data:
                        if isinstance(s_data["option_hints"], dict):
                            # Limit size of option hints dictionary
                            if len(s_data["option_hints"]) <= 50:
                                setting.option_hints = s_data["option_hints"]
                
                    self.settings_by_category[enum_cat].append(setting)
                    settings_loaded += 1
                    
                except Exception as e:
                    _safe_print(f"Error loading setting {s_data.get('name')}: {e}")
                    continue
            
            if settings_loaded == 0:
                _safe_print(f"Warning: No valid settings loaded for category '{cat_name}'")

    def get_settings_for_category(self, category: SettingCategory) -> List[RegistrySetting]:
        """Get all settings for a specific category."""
        return self.settings_by_category.get(category, [])

    def get_categories(self) -> List[SettingCategory]:
        """Get list of available categories."""
        return list(self.settings_by_category.keys())