"""
Registry handler for WinSet - reads and writes Windows Registry values.
"""

import winreg
import re
import logging
from typing import Any, Optional, List

logger = logging.getLogger(__name__)


class RegistryHandler:
    """Handles reading and writing Windows Registry values securely."""

    # Maps hive name strings to winreg constants
    HIVE_MAP = {
        "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
        "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
        "HKEY_CLASSES_ROOT": winreg.HKEY_CLASSES_ROOT,
        "HKEY_CURRENT_CONFIG": winreg.HKEY_CURRENT_CONFIG,
        "HKEY_USERS": winreg.HKEY_USERS,
    }

    # Maps type name strings to winreg constants
    TYPE_MAP = {
        "REG_SZ": winreg.REG_SZ,
        "REG_DWORD": winreg.REG_DWORD,
        "REG_BINARY": winreg.REG_BINARY,
        "REG_MULTI_SZ": winreg.REG_MULTI_SZ,
        "REG_EXPAND_SZ": winreg.REG_EXPAND_SZ,
        "REG_QWORD": winreg.REG_QWORD,
    }

    # Sensitive registry paths to block
    BLOCKED_PATHS = [
        r'\\Security\\',
        r'\\SAM\\',
        r'\\System\\CurrentControlSet\\Control\\SecurePipeServers',
        r'\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon\\SpecialAccounts',
    ]

    def _validate_key_path(self, key_path: str) -> bool:
        """
        Validate registry key path to prevent security issues.
        
        Args:
            key_path: Registry key path to validate
            
        Returns:
            True if valid, raises ValueError if invalid
        """
        # Check length
        if len(key_path) > 512:
            raise ValueError(f"Key path too long: {len(key_path)} chars (max 512)")
        
        # Prevent path traversal
        if ".." in key_path:
            raise ValueError(f"Key path contains path traversal: {key_path}")
        
        # Block UNC paths
        if key_path.startswith("\\\\"):
            raise ValueError(f"Key path starts with network path: {key_path}")
        
        # Block sensitive paths
        for blocked in self.BLOCKED_PATHS:
            if re.search(blocked, key_path, re.IGNORECASE):
                raise ValueError(f"Access to blocked registry path: {key_path}")
        
        # Check for suspicious characters.
        # Allow alphanumerics, backslash, hyphens, underscores, spaces, curly
        # braces (used in GUID-based power plan paths), dots, dollar signs, and
        # @ signs which all appear in legitimate Windows registry paths.
        if not re.match(r'^[a-zA-Z0-9\\\-_.{}$@ ]+$', key_path):
            raise ValueError(f"Invalid characters in key path: {key_path}")
        
        return True

    def _validate_value_name(self, value_name: str) -> bool:
        """
        Validate registry value name.
        
        Args:
            value_name: Value name to validate
            
        Returns:
            True if valid, raises ValueError if invalid
        """
        if len(value_name) > 255:
            raise ValueError(f"Value name too long: {len(value_name)} chars (max 255)")
        
        # Value names can contain most characters, but block some
        if value_name.startswith("."):
            raise ValueError(f"Value name starts with dot: {value_name}")
        
        return True

    def _get_hive_constant(self, hive: str) -> int:
        """Convert a hive name string to the corresponding winreg constant.

        Args:
            hive: Registry hive name, e.g. 'HKEY_CURRENT_USER'.

        Returns:
            The winreg hive constant.

        Raises:
            ValueError: If the hive name is not recognised.
        """
        if hive not in self.HIVE_MAP:
            raise ValueError(
                f"Unknown registry hive '{hive}'. "
                f"Valid hives: {list(self.HIVE_MAP.keys())}"
            )
        return self.HIVE_MAP[hive]

    def _get_type_constant(self, value_type: str) -> int:
        """Convert a registry type name string to the corresponding winreg constant.

        Args:
            value_type: Registry type name, e.g. 'REG_DWORD'.

        Returns:
            The winreg type constant.

        Raises:
            ValueError: If the type name is not recognised.
        """
        if value_type not in self.TYPE_MAP:
            raise ValueError(
                f"Unknown registry type '{value_type}'. "
                f"Valid types: {list(self.TYPE_MAP.keys())}"
            )
        return self.TYPE_MAP[value_type]

    def read_value(
        self,
        hive: str,
        key_path: str,
        value_name: str,
    ) -> Optional[Any]:
        """Read a value from the Windows Registry.

        Args:
            hive: Registry hive name, e.g. 'HKEY_CURRENT_USER'.
            key_path: Path to the registry key.
            value_name: Name of the value to read.

        Returns:
            The stored value, or None if it could not be read.
        """
        try:
            # Validate inputs
            self._validate_key_path(key_path)
            self._validate_value_name(value_name)
            
            hive_constant = self._get_hive_constant(hive)
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, value_name)
            winreg.CloseKey(key)
            return value
        except FileNotFoundError:
            # Key or value simply doesn't exist on this system — not an error.
            logger.debug("Registry value not found: %s", value_name)
            return None
        except (OSError, ValueError) as e:
            logger.warning("Failed to read registry value '%s': %s", value_name, e)
            return None

    def write_value(
        self,
        hive: str,
        key_path: str,
        value_name: str,
        value_type: str,
        value: Any,
    ) -> bool:
        """Write a value to the Windows Registry.

        Args:
            hive: Registry hive name, e.g. 'HKEY_CURRENT_USER'.
            key_path: Path to the registry key (created if it does not exist).
            value_name: Name of the value to write.
            value_type: Registry type string, e.g. 'REG_DWORD'.
            value: The value to write.

        Returns:
            True on success, False on failure.
        """
        try:
            # Validate inputs
            self._validate_key_path(key_path)
            self._validate_value_name(value_name)
            
            # Validate value based on type
            if not self._validate_value(value_type, value):
                return False
            
            hive_constant = self._get_hive_constant(hive)
            type_constant = self._get_type_constant(value_type)

            # OpenKey with write access; create key if it doesn't exist
            key = winreg.OpenKey(
                hive_constant,
                key_path,
                0,
                winreg.KEY_SET_VALUE,
            )
            winreg.SetValueEx(key, value_name, 0, type_constant, value)
            winreg.CloseKey(key)
            return True
        except (FileNotFoundError, OSError) as e:
            # Key may not exist — try creating it
            try:
                self._validate_key_path(key_path)
                hive_constant = self._get_hive_constant(hive)
                type_constant = self._get_type_constant(value_type)
                key = winreg.CreateKey(hive_constant, key_path)
                winreg.SetValueEx(key, value_name, 0, type_constant, value)
                winreg.CloseKey(key)
                return True
            except Exception as inner_e:
                print(f"Failed to write registry value '{value_name}': {inner_e}")
                return False
        except ValueError as e:
            print(f"Invalid hive or type: {e}")
            return False

    def _validate_value(self, value_type: str, value: Any) -> bool:
        """
        Validate a registry value based on its type.
        
        Args:
            value_type: Registry type
            value: The value to validate
            
        Returns:
            True if valid, False otherwise
        """
        try:
            if value_type == "REG_DWORD":
                if not isinstance(value, int):
                    return False
                if value < 0 or value > 4294967295:
                    return False
            elif value_type in ["REG_SZ", "REG_EXPAND_SZ"]:
                if not isinstance(value, str):
                    return False
                if len(value) > 32767:
                    return False
            elif value_type == "REG_BINARY":
                if not isinstance(value, (bytes, bytearray)):
                    return False
                if len(value) > 1024 * 1024:  # 1MB max
                    return False
            elif value_type == "REG_MULTI_SZ":
                if not isinstance(value, list):
                    return False
                if len(value) > 100:  # Max strings
                    return False
                for s in value:
                    if len(s) > 1000:
                        return False
            elif value_type == "REG_QWORD":
                if not isinstance(value, int):
                    return False
                if value < 0 or value > 18446744073709551615:
                    return False
            return True
        except Exception:
            return False

    def delete_value(self, hive: str, key_path: str, value_name: str) -> bool:
        """Delete a value from the Windows Registry.

        Args:
            hive: Registry hive name.
            key_path: Path to the registry key.
            value_name: Name of the value to delete.

        Returns:
            True on success, False on failure.
        """
        try:
            # Validate inputs
            self._validate_key_path(key_path)
            self._validate_value_name(value_name)
            
            hive_constant = self._get_hive_constant(hive)
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, value_name)
            winreg.CloseKey(key)
            return True
        except (FileNotFoundError, OSError, ValueError) as e:
            print(f"Failed to delete registry value '{value_name}': {e}")
            return False

    def key_exists(self, hive: str, key_path: str) -> bool:
        """
        Check if a registry key exists.
        
        Args:
            hive: Registry hive name
            key_path: Path to the registry key
            
        Returns:
            True if the key exists, False otherwise
        """
        try:
            self._validate_key_path(key_path)
            hive_constant = self._get_hive_constant(hive)
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_READ)
            winreg.CloseKey(key)
            return True
        except (FileNotFoundError, OSError, ValueError):
            return False
        
        # src/core/registry_handler.py

    def read_multiple_values(self, hive: str, key_path: str, value_names: List[str]) -> dict:
        """
        Read multiple registry values from the same key.
        
        Args:
            hive: Registry hive name
            key_path: Path to the registry key
            value_names: List of value names to read
        
        Returns:
            Dictionary with value names as keys and their values as values (None if not found)
        """
        results = {}
        
        for value_name in value_names:
            results[value_name] = self.read_value(hive, key_path, value_name)
        
        return results
    
    def write_multiple_values(self, operations: List[tuple]) -> dict:
        """
        Write multiple registry values in bulk.
        
        Args:
            operations: List of tuples (hive, key_path, value_name, value_type, value)
        
        Returns:
            Dictionary with operation indices as keys and boolean success as values
        """
        results = {}
        
        for i, op in enumerate(operations):
            try:
                if len(op) == 5:
                    hive, key_path, value_name, value_type, value = op
                else:
                    # Handle different tuple lengths if needed
                    results[i] = False
                    continue
                    
                success = self.write_value(hive, key_path, value_name, value_type, value)
                results[i] = success
            except Exception:
                results[i] = False
        
        return results