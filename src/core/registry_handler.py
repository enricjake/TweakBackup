"""
Registry handler for WinSet - reads and writes Windows Registry values.
"""

import winreg  # Python's built-in module for low-level Windows Registry access via the Windows API
import re  # Regular expressions, used for validating key paths against blocked patterns and character sets
import logging  # Standard logging module for debug/warning/error output
from typing import (
    Any,
    Optional,
    List,
)  # Type hints: Any = any value, Optional = value or None, List = typed list

# Module-level logger scoped to this file, used throughout the class for debug and warning messages
logger = logging.getLogger(__name__)


class RegistryHandler:
    """Handles reading and writing Windows Registry values securely."""

    # HIVE_MAP: Maps human-readable Windows Registry hive name strings to their
    # corresponding winreg integer constants. The Windows Registry is divided into
    # top-level "hives" that each serve a different scope:
    #   - HKEY_CURRENT_USER (HKCU): Settings for the currently logged-in user
    #   - HKEY_LOCAL_MACHINE (HKLM): System-wide settings shared across all users
    #   - HKEY_CLASSES_ROOT (HKCR): File type associations and COM object registrations
    #   - HKEY_CURRENT_CONFIG (HKCC): Hardware profile settings for the current boot
    #   - HKEY_USERS (HKU): Default user profile and all loaded user profiles
    # This map allows WinSet configuration files to reference hives by string name
    # while the handler translates them to the numeric constants that winreg requires.
    HIVE_MAP = {
        "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
        "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
        "HKEY_CLASSES_ROOT": winreg.HKEY_CLASSES_ROOT,
        "HKEY_CURRENT_CONFIG": winreg.HKEY_CURRENT_CONFIG,
        "HKEY_USERS": winreg.HKEY_USERS,
    }

    # TYPE_MAP: Maps human-readable registry value type strings to their
    # corresponding winreg integer constants. Each type defines how the data
    # is stored and interpreted:
    #   - REG_SZ: A null-terminated Unicode string (the most common type)
    #   - REG_DWORD: A 32-bit unsigned integer (0 to 4,294,967,295)
    #   - REG_BINARY: Raw binary data of arbitrary length
    #   - REG_MULTI_SZ: An array of null-terminated strings, stored as a list in Python
    #   - REG_EXPAND_SZ: A string containing unexpanded environment variable references
    #     (e.g. "%SystemRoot%") that are expanded at runtime
    #   - REG_QWORD: A 64-bit unsigned integer (0 to 18,446,744,073,709,551,615)
    # This map lets WinSet config files specify types by name while the handler
    # converts them to the numeric constants winreg needs for API calls.
    TYPE_MAP = {
        "REG_SZ": winreg.REG_SZ,
        "REG_DWORD": winreg.REG_DWORD,
        "REG_BINARY": winreg.REG_BINARY,
        "REG_MULTI_SZ": winreg.REG_MULTI_SZ,
        "REG_EXPAND_SZ": winreg.REG_EXPAND_SZ,
        "REG_QWORD": winreg.REG_QWORD,
    }

    # BLOCKED_PATHS: A list of regex patterns matching registry paths that WinSet
    # must never read or modify. These are security-critical areas of the registry:
    #   - \Security\: Contains security policy data, audit settings, and SAM pointers
    #   - \SAM\: The Security Accounts Manager database with user/password hashes
    #   - \System\CurrentControlSet\Control\SecurePipeServers: Named pipe security
    #     descriptors that control access to system services
    #   - \Software\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts:
    #     Controls which user accounts are hidden from the login screen
    # Any registry operation targeting a path matching these patterns will be
    # rejected with a ValueError to prevent accidental or malicious damage.
    BLOCKED_PATHS = [
        r"\\Security\\",
        r"\\SAM\\",
        r"\\System\\CurrentControlSet\\Control\\SecurePipeServers",
        r"\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon\\SpecialAccounts",
    ]

    def _validate_key_path(self, key_path: str) -> bool:
        """
        Validate registry key path to prevent security issues.

        Args:
            key_path: Registry key path to validate (e.g. "Software\\Microsoft\\Windows\\CurrentVersion")

        Returns:
            True if valid, raises ValueError if invalid
        """
        # Check length — Windows Registry key paths have a maximum length of 512 characters;
        # anything longer is rejected to prevent buffer-related issues or abuse
        if len(key_path) > 512:
            raise ValueError(f"Key path too long: {len(key_path)} chars (max 512)")

        # Prevent path traversal — ".." could be used to navigate outside intended
        # registry scope in certain edge cases, so it is blocked outright
        if ".." in key_path:
            raise ValueError(f"Key path contains path traversal: {key_path}")

        # Block UNC paths — paths starting with "\\" refer to network resources
        # (e.g. remote registries), which WinSet should never access for safety
        if key_path.startswith("\\\\"):
            raise ValueError(f"Key path starts with network path: {key_path}")

        # Block sensitive paths — iterate through BLOCKED_PATHS and use regex matching
        # (case-insensitive) to reject any key path that contains a blocked substring.
        # This prevents WinSet from modifying security-critical registry areas.
        for blocked in self.BLOCKED_PATHS:
            if re.search(blocked, key_path, re.IGNORECASE):
                raise ValueError(f"Access to blocked registry path: {key_path}")

        # Check for suspicious characters.
        # Allow alphanumerics, backslash, hyphens, underscores, spaces, curly
        # braces (used in GUID-based power plan paths), dots, dollar signs, and
        # @ signs which all appear in legitimate Windows registry paths.
        # The regex anchors to start (^) and end ($) of the string so the entire
        # path must consist of allowed characters.
        if not re.match(r"^[a-zA-Z0-9\\\-_.{} ]+$", key_path):
            raise ValueError(f"Invalid characters in key path: {key_path}")

        return True

    def _validate_value_name(self, value_name: str) -> bool:
        """
        Validate registry value name.

        Args:
            value_name: Value name to validate (e.g. "MySetting" or "" for default value)

        Returns:
            True if valid, raises ValueError if invalid
        """
        # Windows Registry value names are limited to 255 characters maximum;
        # reject anything longer to stay within API constraints
        if len(value_name) > 255:
            raise ValueError(f"Value name too long: {len(value_name)} chars (max 255)")

        # Value names that start with a dot are reserved by the system and could
        # cause unexpected behavior or indicate a malformed name, so block them
        if value_name.startswith("."):
            raise ValueError(f"Value name starts with dot: {value_name}")

        return True

    def _get_hive_constant(self, hive: str) -> int:
        """Convert a hive name string to the corresponding winreg constant.

        Args:
            hive: Registry hive name, e.g. 'HKEY_CURRENT_USER'. Must be one of
                  the keys defined in HIVE_MAP.

        Returns:
            The winreg hive constant (an integer) that the Windows API requires
            for registry operations like OpenKey, CreateKey, etc.

        Raises:
            ValueError: If the hive name is not recognised (not a key in HIVE_MAP).
        """
        # Look up the hive name in HIVE_MAP; if not found, raise a descriptive
        # error listing all valid hive names so the user can correct their config
        if hive not in self.HIVE_MAP:
            raise ValueError(
                f"Unknown registry hive '{hive}'. "
                f"Valid hives: {list(self.HIVE_MAP.keys())}"
            )
        return self.HIVE_MAP[hive]

    def _get_type_constant(self, value_type: str) -> int:
        """Convert a registry type name string to the corresponding winreg constant.

        Args:
            value_type: Registry type name, e.g. 'REG_DWORD'. Must be one of
                        the keys defined in TYPE_MAP.

        Returns:
            The winreg type constant (an integer) that the Windows API requires
            for SetValueEx and QueryValueEx calls.

        Raises:
            ValueError: If the type name is not recognised (not a key in TYPE_MAP).
        """
        # Look up the type name in TYPE_MAP; if not found, raise a descriptive
        # error listing all valid type names so the user can correct their config
        if value_type not in self.TYPE_MAP:
            raise ValueError(
                f"Unknown registry type '{value_type}'. "
                f"Valid types: {list(self.TYPE_MAP.keys())}"
            )
        return self.TYPE_MAP[value_type]

    def read_value(
        self,
        hive: str,  # The top-level hive name (e.g. "HKEY_CURRENT_USER")
        key_path: str,  # The subkey path within the hive (e.g. "Software\\MyApp")
        value_name: str,  # The name of the specific value to read (e.g. "InstallPath")
    ) -> Optional[Any]:
        """Read a value from the Windows Registry.

        Args:
            hive: Registry hive name, e.g. 'HKEY_CURRENT_USER'.
            key_path: Path to the registry key.
            value_name: Name of the value to read.

        Returns:
            The stored value (type depends on registry value type), or None if
            the key/value does not exist or the read operation fails.
        """
        try:
            # Step 1: Validate the key path and value name against security rules
            # (length limits, blocked paths, traversal attempts, allowed characters)
            self._validate_key_path(key_path)
            self._validate_value_name(value_name)

            # Step 2: Convert the hive name string to its winreg integer constant
            hive_constant = self._get_hive_constant(hive)

            # Step 3: Open the registry key with read-only access (KEY_READ).
            # Parameters: (hive_constant, sub_key, reserved=0, access_rights)
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_READ)

            # Step 4: Query the value data. QueryValueEx returns a tuple of
            # (value, type_code); we only need the value, so we discard the type.
            value, _ = winreg.QueryValueEx(key, value_name)

            # Step 5: Close the key handle to release the Windows resource
            winreg.CloseKey(key)
            return value
        except FileNotFoundError:
            # Key or value simply doesn't exist on this system — not an error,
            # just return None so callers can use a default value gracefully
            logger.debug("Registry value not found: %s", value_name)
            return None
        except (OSError, ValueError) as e:
            # Catch OS-level errors (e.g. access denied) and validation errors;
            # log a warning and return None rather than crashing the application
            logger.warning("Failed to read registry value '%s': %s", value_name, e)
            return None

    def write_value(
        self,
        hive: str,  # The top-level hive name (e.g. "HKEY_LOCAL_MACHINE")
        key_path: str,  # The subkey path (e.g. "Software\\MyApp\\Settings")
        value_name: str,  # The name of the value to write (e.g. "Version")
        value_type: str,  # The registry type string (e.g. "REG_SZ", "REG_DWORD")
        value: Any,  # The actual data to store (type must match value_type)
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
            # Step 1: Validate key path and value name for security
            self._validate_key_path(key_path)
            self._validate_value_name(value_name)

            # Step 2: Validate that the value matches the expected type constraints
            # (e.g. REG_DWORD must be an int in range 0-4294967295).
            # If validation fails, return False immediately without touching the registry.
            if not self._validate_value(value_type, value):
                return False

            # Step 3: Resolve the hive name and type name to their winreg constants
            hive_constant = self._get_hive_constant(hive)
            type_constant = self._get_type_constant(value_type)

            # Step 4: Open the existing key with write access (KEY_SET_VALUE).
            # This will raise FileNotFoundError if the key doesn't exist,
            # which is caught below to trigger key creation.
            # Parameters: (hive_constant, sub_key, reserved=0, access_rights)
            key = winreg.OpenKey(
                hive_constant,
                key_path,
                0,
                winreg.KEY_SET_VALUE,
            )

            # Step 5: Set the value. Parameters:
            #   (key_handle, value_name, reserved=0, type_constant, data)
            winreg.SetValueEx(key, value_name, 0, type_constant, value)

            # Step 6: Close the key handle to release the Windows resource
            winreg.CloseKey(key)
            return True
        except (FileNotFoundError, OSError) as e:
            # The key may not exist yet — attempt to create it as a fallback
            try:
                # Re-validate the key path (defense-in-depth in case state changed)
                self._validate_key_path(key_path)

                # Re-resolve constants in case this is a fresh execution path
                hive_constant = self._get_hive_constant(hive)
                type_constant = self._get_type_constant(value_type)

                # CreateKey creates the key (and any missing parent keys) if it
                # doesn't already exist; returns an open handle to the key
                key = winreg.CreateKey(hive_constant, key_path)

                # Write the value into the newly created key
                winreg.SetValueEx(key, value_name, 0, type_constant, value)

                # Close the key handle
                winreg.CloseKey(key)
                return True
            except Exception as inner_e:
                # If key creation also fails, report the error and return False
                print(f"Failed to write registry value '{value_name}': {inner_e}")
                return False
        except ValueError as e:
            # Raised by _get_hive_constant or _get_type_constant if hive/type is invalid
            print(f"Invalid hive or type: {e}")
            return False

    def _validate_value(self, value_type: str, value: Any) -> bool:
        """
        Validate a registry value based on its type.

        Args:
            value_type: Registry type string (e.g. "REG_DWORD", "REG_SZ")
            value: The value to validate — its Python type and range are checked
                   against the constraints of the specified registry type

        Returns:
            True if the value is valid for the given type, False otherwise
        """
        try:
            if value_type == "REG_DWORD":
                # REG_DWORD is a 32-bit unsigned integer: must be a Python int
                # within the range [0, 4294967295] (0xFFFFFFFF)
                if not isinstance(value, int):
                    return False
                if value < 0 or value > 4294967295:
                    return False
            elif value_type in ["REG_SZ", "REG_EXPAND_SZ"]:
                # REG_SZ and REG_EXPAND_SZ are string types: must be a Python str
                # with a maximum length of 32767 characters (Windows API limit)
                if not isinstance(value, str):
                    return False
                if len(value) > 32767:
                    return False
            elif value_type == "REG_BINARY":
                # REG_BINARY stores raw bytes: must be a bytes or bytearray object
                # with a maximum size of 1 MB (1048576 bytes) to prevent abuse
                if not isinstance(value, (bytes, bytearray)):
                    return False
                if len(value) > 1024 * 1024:  # 1MB max
                    return False
            elif value_type == "REG_MULTI_SZ":
                # REG_MULTI_SZ is an array of strings: must be a Python list
                # containing at most 100 strings, each at most 1000 characters
                if not isinstance(value, list):
                    return False
                if len(value) > 100:  # Max strings
                    return False
                for s in value:
                    if len(s) > 1000:
                        return False
            elif value_type == "REG_QWORD":
                # REG_QWORD is a 64-bit unsigned integer: must be a Python int
                # within the range [0, 18446744073709551615] (0xFFFFFFFFFFFFFFFF)
                if not isinstance(value, int):
                    return False
                if value < 0 or value > 18446744073709551615:
                    return False
            # If none of the type-specific checks failed, the value is valid
            return True
        except Exception:
            # Catch any unexpected errors during validation and treat as invalid
            # rather than letting them propagate up and crash the caller
            return False

    def delete_value(self, hive: str, key_path: str, value_name: str) -> bool:
        """Delete a value from the Windows Registry.

        Args:
            hive: Registry hive name (e.g. "HKEY_CURRENT_USER").
            key_path: Path to the registry key containing the value.
            value_name: Name of the value to delete.

        Returns:
            True on success, False on failure.
        """
        try:
            # Step 1: Validate key path and value name for security
            self._validate_key_path(key_path)
            self._validate_value_name(value_name)

            # Step 2: Resolve the hive name to its winreg integer constant
            hive_constant = self._get_hive_constant(hive)

            # Step 3: Open the key with write access (KEY_SET_VALUE) since
            # deleting a value requires modification rights
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_SET_VALUE)

            # Step 4: Delete the named value from the key
            winreg.DeleteValue(key, value_name)

            # Step 5: Close the key handle to release the Windows resource
            winreg.CloseKey(key)
            return True
        except (FileNotFoundError, OSError, ValueError) as e:
            # FileNotFoundError: key or value doesn't exist
            # OSError: access denied or other OS-level failure
            # ValueError: invalid hive name or failed validation
            # All cases: log the error and return False
            print(f"Failed to delete registry value '{value_name}': {e}")
            return False

    def key_exists(self, hive: str, key_path: str) -> bool:
        """
        Check if a registry key exists.

        Args:
            hive: Registry hive name (e.g. "HKEY_LOCAL_MACHINE")
            key_path: Path to the registry key to check (e.g. "Software\\MyApp")

        Returns:
            True if the key exists, False otherwise
        """
        try:
            # Step 1: Validate the key path for security
            self._validate_key_path(key_path)

            # Step 2: Resolve the hive name to its winreg constant
            hive_constant = self._get_hive_constant(hive)

            # Step 3: Attempt to open the key with read-only access.
            # If the key exists, OpenKey succeeds; if not, it raises FileNotFoundError.
            # Parameters: (hive_constant, sub_key, reserved=0, access_rights)
            key = winreg.OpenKey(hive_constant, key_path, 0, winreg.KEY_READ)

            # Step 4: Close the key handle immediately — we only needed to
            # confirm the key exists, not read any values from it
            winreg.CloseKey(key)
            return True
        except (FileNotFoundError, OSError, ValueError):
            # Any of these exceptions means the key doesn't exist or isn't accessible
            return False

    def read_multiple_values(
        self, hive: str, key_path: str, value_names: List[str]
    ) -> dict:
        """
        Read multiple registry values from the same key.

        Args:
            hive: Registry hive name (e.g. "HKEY_CURRENT_USER")
            key_path: Path to the registry key containing all the values
            value_names: List of value names to read (e.g. ["Version", "InstallPath"])

        Returns:
            Dictionary mapping each value name to its stored value (or None if
            that particular value was not found or could not be read).
            Example: {"Version": "1.0", "InstallPath": None}
        """
        # results accumulates the output: one entry per requested value name
        results = {}

        # Iterate over each value name and delegate to read_value, which handles
        # validation, key opening, querying, and error handling for each individual read.
        # This is a convenience method that avoids the caller needing to loop themselves.
        for value_name in value_names:
            results[value_name] = self.read_value(hive, key_path, value_name)

        return results

    def write_multiple_values(self, operations: List[tuple]) -> dict:
        """
        Write multiple registry values in bulk.

        Args:
            operations: List of tuples, each containing:
                (hive, key_path, value_name, value_type, value)
                Example: [
                    ("HKEY_CURRENT_USER", "Software\\MyApp", "Version", "REG_SZ", "2.0"),
                    ("HKEY_CURRENT_USER", "Software\\MyApp", "Enabled", "REG_DWORD", 1),
                ]

        Returns:
            Dictionary mapping each operation index (0-based) to a boolean indicating
            whether that write succeeded. Example: {0: True, 1: False}
        """
        # results tracks success/failure of each write operation by its index
        results = {}

        # enumerate gives us both the index (i) and the operation tuple (op)
        for i, op in enumerate(operations):
            try:
                # Each operation tuple must have exactly 5 elements:
                # (hive, key_path, value_name, value_type, value)
                if len(op) == 5:
                    # Unpack the tuple into named variables for clarity
                    hive, key_path, value_name, value_type, value = op
                else:
                    # If the tuple has a different length, it's malformed;
                    # mark this operation as failed and skip to the next one
                    results[i] = False
                    continue

                # Delegate to write_value which handles validation, key creation,
                # writing, and error handling for each individual write
                success = self.write_value(
                    hive, key_path, value_name, value_type, value
                )
                results[i] = success
            except Exception:
                # Catch any unexpected errors (e.g. tuple unpacking failures,
                # type errors) and mark this operation as failed without
                # interrupting the remaining operations in the batch
                results[i] = False

        return results
