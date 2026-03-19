"""
Registry handler for WinSet - reads and writes Windows Registry values.
"""

import winreg
import logging
from typing import Any, Optional, Dict, List, Tuple

# Configure logging
logger = logging.getLogger(__name__)


class RegistryHandler:
    """Handles reading and writing Windows Registry values."""

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
            hive_constant = self._get_hive_constant(hive)
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, value_name)
            winreg.CloseKey(key)
            logger.debug(f"Successfully read registry value: {hive}\\{key_path}\\{value_name}")
            return value
        except (FileNotFoundError, OSError, ValueError) as e:
            logger.warning(f"Failed to read registry value '{value_name}': {e}")
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
            logger.info(f"Successfully wrote registry value: {hive}\\{key_path}\\{value_name} = {value}")
            return True
        except (FileNotFoundError, OSError) as e:
            # Key may not exist — try creating it
            try:
                hive_constant = self._get_hive_constant(hive)
                type_constant = self._get_type_constant(value_type)
                key = winreg.CreateKey(hive_constant, key_path)
                winreg.SetValueEx(key, value_name, 0, type_constant, value)
                winreg.CloseKey(key)
                logger.info(f"Created registry key and wrote value: {hive}\\{key_path}\\{value_name} = {value}")
                return True
            except Exception as inner_e:
                logger.error(f"Failed to write registry value '{value_name}': {inner_e}")
                return False
        except ValueError as e:
            logger.error(f"Invalid hive or type: {e}")
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
            hive_constant = self._get_hive_constant(hive)
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, value_name)
            winreg.CloseKey(key)
            logger.info(f"Successfully deleted registry value: {hive}\\{key_path}\\{value_name}")
            return True
        except (FileNotFoundError, OSError, ValueError) as e:
            logger.warning(f"Failed to delete registry value '{value_name}': {e}")
            return False

    def delete_key(self, hive: str, key_path: str) -> bool:
        """Delete a registry key and all its subkeys/values.

        Args:
            hive: Registry hive name.
            key_path: Path to the registry key to delete.

        Returns:
            True on success, False on failure.
        """
        try:
            hive_constant = self._get_hive_constant(hive)
            winreg.DeleteKey(hive_constant, key_path)
            logger.info(f"Successfully deleted registry key: {hive}\\{key_path}")
            return True
        except (FileNotFoundError, OSError, ValueError) as e:
            logger.warning(f"Failed to delete registry key '{key_path}': {e}")
            return False

    def key_exists(self, hive: str, key_path: str) -> bool:
        """Check if a registry key exists.

        Args:
            hive: Registry hive name.
            key_path: Path to the registry key.

        Returns:
            True if the key exists, False otherwise.
        """
        try:
            hive_constant = self._get_hive_constant(hive)
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_READ)
            winreg.CloseKey(key)
            return True
        except (FileNotFoundError, OSError, ValueError):
            return False

    def value_exists(self, hive: str, key_path: str, value_name: str) -> bool:
        """Check if a registry value exists.

        Args:
            hive: Registry hive name.
            key_path: Path to the registry key.
            value_name: Name of the value to check.

        Returns:
            True if the value exists, False otherwise.
        """
        try:
            hive_constant = self._get_hive_constant(hive)
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, value_name)
            winreg.CloseKey(key)
            return True
        except (FileNotFoundError, OSError, ValueError):
            return False

    def read_multiple_values(
        self, 
        hive: str, 
        key_path: str, 
        value_names: List[str]
    ) -> Dict[str, Optional[Any]]:
        """Read multiple values from the same registry key.

        Args:
            hive: Registry hive name.
            key_path: Path to the registry key.
            value_names: List of value names to read.

        Returns:
            Dictionary mapping value names to their values (or None if not found).
        """
        results = {}
        try:
            hive_constant = self._get_hive_constant(hive)
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_READ)
            
            for value_name in value_names:
                try:
                    value, _ = winreg.QueryValueEx(key, value_name)
                    results[value_name] = value
                    logger.debug(f"Read registry value: {hive}\\{key_path}\\{value_name}")
                except (FileNotFoundError, OSError):
                    results[value_name] = None
                    logger.warning(f"Registry value not found: {hive}\\{key_path}\\{value_name}")
            
            winreg.CloseKey(key)
            return results
        except (FileNotFoundError, OSError, ValueError) as e:
            logger.error(f"Failed to read multiple registry values: {e}")
            return {name: None for name in value_names}

    def write_multiple_values(
        self, 
        operations: List[Tuple[str, str, str, str, Any]]
    ) -> Dict[str, bool]:
        """Write multiple registry values in a batch operation.

        Args:
            operations: List of tuples containing (hive, key_path, value_name, value_type, value).

        Returns:
            Dictionary mapping operation indices to success status.
        """
        results = {}
        for i, (hive, key_path, value_name, value_type, value) in enumerate(operations):
            try:
                success = self.write_value(hive, key_path, value_name, value_type, value)
                results[i] = success
            except Exception as e:
                logger.error(f"Batch operation {i} failed: {e}")
                results[i] = False
        return results
