"""
Preset Manager for WinSet - manages preset configurations.
"""

import json  # Used for reading/writing preset files in JSON format
import os  # Used for filesystem operations (path checks, directory creation, file listing)
import re  # Used for regex validation of preset IDs and setting IDs
import tempfile  # Used to resolve the system temp directory for path validation
from pathlib import Path  # Used for secure path resolution and parent-directory checks
from typing import (
    List,
    Dict,
    Optional,
    Any,
    Tuple,
)  # Type hints for IDE support and documentation
from datetime import datetime  # Used to timestamp newly created presets


class PresetManager:
    """
    Manages preset configurations for WinSet.
    Handles loading, validating, and applying presets.
    """

    # ---------------------------------------------------------------------------
    # Initialization
    # ---------------------------------------------------------------------------

    def __init__(self, presets_dir: Optional[str] = None):
        """
        Initialize the preset manager.

        Args:
            presets_dir: Optional custom presets directory path. When provided
                         (e.g. during testing), only this single directory is
                         scanned for preset files.  When *None*, the manager
                         searches the default built-in, APPDATA, Documents, and
                         home directories.
        """

        # Dictionary mapping preset IDs (str) to their full preset data
        # dictionaries.  Each data dict contains keys like 'name',
        # 'description', 'settings', etc.
        self.presets: Dict[str, Dict[str, Any]] = {}

        # Dictionary mapping each preset ID to the filesystem directory from
        # which it was loaded.  Used to determine whether a preset is built-in
        # or user-created and to locate the file for deletion.
        self.preset_sources: Dict[str, str] = {}  # preset_id -> directory

        if presets_dir:
            # Custom/testing mode: validate the supplied path and use it as the
            # sole directory to scan.
            self.presets_dirs = [self._validate_preset_path(presets_dir)]
        else:
            # Default mode: build a list of directories to scan in priority
            # order.  Later directories can override presets from earlier ones
            # when IDs collide.
            self.presets_dirs = []

            # ---- 1. Built-in presets directory (shipped with the application) ----
            # Navigate three levels up from this file to reach the project root,
            # then look for a sibling "presets" folder.
            base_dir = os.path.dirname(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            )
            self.builtin_dir = os.path.join(base_dir, "presets")
            if os.path.exists(self.builtin_dir):
                self.presets_dirs.append(self.builtin_dir)

            # ---- 2. User-writable presets directory (persistent in LOCALAPPDATA) ----
            # This is the preferred location for presets created via the GUI.
            app_data_path = os.getenv(
                "LOCALAPPDATA"
            )  # e.g. C:\Users\<user>\AppData\Local
            if not app_data_path:
                # Fallback to the user's home directory if LOCALAPPDATA is unset
                app_data_path = os.path.expanduser("~")

            user_dir = os.path.join(app_data_path, "WinSet", "presets")
            os.makedirs(user_dir, exist_ok=True)  # Create if it doesn't exist
            self.presets_dirs.append(user_dir)

            # ---- 3. Documents/WinSet/presets directory ----
            # A more discoverable location for manually managed preset files.
            self.home_presets_dir = os.path.join(
                os.path.expanduser("~"), "Documents", "WinSet", "presets"
            )
            if os.path.exists(self.home_presets_dir):
                self.presets_dirs.append(self.home_presets_dir)

            # ---- 4. User home directory root ----
            # Scan the home folder itself for loose *.preset.json files.
            self.home_dir = os.path.expanduser("~")
            self.presets_dirs.append(self.home_dir)

            # The directory used when the user creates a new preset through the
            # application (GUI or CLI).  Defaults to the APPDATA user directory.
            self.writable_presets_dir = user_dir

        # Immediately load and validate all preset files found in the configured
        # directories so the manager is ready to use after construction.
        self._load_presets()

    # ---------------------------------------------------------------------------
    # Built-in / directory helpers
    # ---------------------------------------------------------------------------

    def is_builtin(self, preset_id: str) -> bool:
        """
        Check whether a preset was shipped with the application (built-in)
        rather than created by the user.

        Args:
            preset_id: The unique identifier of the preset to check.

        Returns:
            True if the preset's source directory matches the built-in directory.
        """
        source = self.preset_sources.get(preset_id)
        if not source:
            return False
        # Compare the source directory against the builtin_dir attribute.
        # getattr with a default handles the case where builtin_dir was never set
        # (custom presets_dir mode).
        return source == getattr(self, "builtin_dir", "")

    def ensure_user_presets_dir(self) -> bool:
        """
        Ensure that the Documents/WinSet/presets directory exists on disk,
        creating it (and any parents) if necessary.

        Returns:
            True if the directory already existed, False if it had to be created
            or if creation failed.
        """
        path = os.path.join(os.path.expanduser("~"), "Documents", "WinSet", "presets")
        if os.path.exists(path):
            return True  # Directory was already present
        try:
            os.makedirs(path, exist_ok=True)
            return False  # Directory was just created
        except Exception:
            return False  # Creation failed (permissions, etc.)

    def user_presets_dir_exists(self) -> bool:
        """
        Check whether the Documents/WinSet/presets directory currently exists
        without attempting to create it.

        Returns:
            True if the directory exists, False otherwise.
        """
        path = os.path.join(os.path.expanduser("~"), "Documents", "WinSet", "presets")
        return os.path.exists(path)

    # ---------------------------------------------------------------------------
    # Path validation
    # ---------------------------------------------------------------------------

    def _validate_preset_path(self, path: str) -> str:
        """
        Validate a preset directory path for security and resolve it to an
        absolute path.  Only paths inside allowed base directories (project
        root, user home, Documents/WinSet, or the system temp directory) are
        accepted to prevent directory-traversal and injection attacks.

        Args:
            path: The directory path string to validate.

        Returns:
            The resolved absolute path as a string.

        Raises:
            ValueError: If the path contains dangerous patterns or points
                        outside allowed directories.
        """
        try:
            # ---- Dangerous-string check ----
            # Reject paths containing characters or keywords that could
            # indicate command injection or path traversal attempts.
            # IMPORTANT: Check the ORIGINAL input string, not the resolved path,
            # to catch ".." and other traversal attempts that might resolve to
            # allowed directories.
            dangerous_strings = [
                ";",
                "--",
                "..",
                '"',
                "'",
                "*",
                "DROP",
                "DELETE",
                "INSERT",
                "UPDATE",
            ]
            original_path_lower = path.lower()  # Use original input for checking
            for s in dangerous_strings:
                if s in original_path_lower:
                    raise ValueError(f"Dangerous string detected in path: {s}")

            # ---- Additional traversal check ----
            # Ensure the path doesn't try to escape via ".." after normalization
            # by comparing resolved path components against original intent
            import os

            normalized = os.path.normpath(path)
            if ".." in normalized:
                raise ValueError(f"Path traversal detected in path: {path}")

            # Resolve the path to an absolute, symlink-resolved form
            safe_path = Path(path).resolve()

            # ---- Allowed-directory check ----
            # The path is permitted if it resides under the system temp dir,
            # the user home dir, the project root, or Documents/WinSet/presets.

            # System temporary directory (used for tests)
            temp_dir = Path(tempfile.gettempdir()).resolve()
            is_temp_path = temp_dir in safe_path.parents or safe_path == temp_dir

            # User home directory
            home_dir = Path(os.path.expanduser("~")).resolve()
            is_home_path = home_dir in safe_path.parents or safe_path == home_dir

            if not is_temp_path and not is_home_path:
                # Project root (three levels up from this source file)
                base_dir = Path(
                    os.path.dirname(
                        os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                    )
                ).resolve()
                # Documents/WinSet/presets
                user_winset = (
                    Path(os.path.expanduser("~")) / "Documents" / "WinSet" / "presets"
                )
                user_winset = user_winset.resolve()

                # Reject if the path is outside both the project root and the
                # user WinSet presets directory.
                if base_dir not in safe_path.parents and safe_path != base_dir:
                    if (
                        user_winset not in safe_path.parents
                        and safe_path != user_winset
                    ):
                        raise ValueError(
                            f"Preset path '{path}' is outside allowed directories"
                        )

            # Create the directory (and parents) if it doesn't exist yet
            os.makedirs(safe_path, exist_ok=True)

            return str(safe_path)

        except Exception as e:
            # Re-raise ValueError as-is; wrap anything else
            if isinstance(e, ValueError):
                raise e
            raise ValueError(f"Invalid preset path: {e}")

    # ---------------------------------------------------------------------------
    # Preset loading and validation
    # ---------------------------------------------------------------------------

    def _load_presets(self):
        """
        Load all valid preset files (*.preset.json) from every configured
        directory.  Later directories in the list override earlier ones when
        preset IDs conflict, allowing user presets to shadow built-in presets.
        """

        # Reset the in-memory caches before reloading
        self.presets = {}

        # Iterate over each configured scan directory in order
        for directory in self.presets_dirs:
            if not os.path.exists(directory):
                continue  # Skip directories that no longer exist

            for filename in os.listdir(directory):
                # Only consider files with the .preset.json extension
                if filename.endswith(".preset.json"):
                    preset_path = os.path.join(directory, filename)

                    # Skip directories that happen to end with the extension
                    if not os.path.isfile(preset_path):
                        continue

                    try:
                        # Read and parse the JSON preset file
                        with open(preset_path, "r", encoding="utf-8") as f:
                            preset_data = json.load(f)

                        # Run structural and security validation on the data
                        if self._validate_preset_data(preset_data):
                            # Derive the preset ID by stripping the 12-char
                            # suffix ".preset.json" from the filename
                            preset_id = filename[:-12]
                            # Store the validated data and track its source dir.
                            # Later directories override earlier ones (user > builtin).
                            self.presets[preset_id] = preset_data
                            self.preset_sources[preset_id] = directory
                        else:
                            print(
                                f"DEBUG: Skipping invalid preset: {filename} (Validation failed)"
                            )

                    except Exception as e:
                        # Log parse errors but don't crash the whole load
                        print(
                            f"DEBUG: Error loading preset {filename} from {directory}: {e}"
                        )

    def _validate_preset_data(self, data: Dict[str, Any]) -> bool:
        """
        Validate the structure and contents of a parsed preset dictionary.

        Ensures the preset has required fields, a properly typed 'settings'
        dict, reasonable size limits, and safe setting IDs (no injection).

        Args:
            data: The parsed preset data dictionary to validate.

        Returns:
            True if the preset is valid and safe to load, False otherwise.
        """

        # Fields that must be present in every preset file
        required_fields = ["name", "description", "settings"]

        # Check for the WinSet application signature.  User-created presets
        # may omit this tag, so we relax the check: if 'app' is missing or
        # not 'WinSet', we still accept the file as long as it has the
        # essential 'name' and 'settings' keys.
        if data.get("app") != "WinSet":
            if not all(field in data for field in ["name", "settings"]):
                return False
            print(
                f"DEBUG: Loading preset '{data.get('name')}' without explicit app signature."
            )

        # Verify every required field exists in the dictionary
        for field in required_fields:
            if field not in data:
                return False

        # The 'settings' value must be a dictionary (mapping of setting IDs to values)
        if not isinstance(data["settings"], dict):
            return False

        # Hard upper bound on the number of settings to prevent abuse
        if len(data["settings"]) > 200:
            return False

        # Per-setting validation
        for setting_id, setting_value in data["settings"].items():
            # Reject setting IDs that contain unsafe characters (injection risk)
            if not self._is_safe_setting_id(setting_id):
                return False

            # Cap individual setting ID length
            if len(setting_id) > 100:
                return False

            # Setting values must be one of the supported basic types
            if not isinstance(setting_value, (str, int, bool, float, dict)):
                return False

            # If the value is a nested dict, limit its size
            if isinstance(setting_value, dict):
                if len(setting_value) > 20:
                    return False

        return True

    def _is_safe_setting_id(self, setting_id: str) -> bool:
        """
        Check whether a setting ID contains only safe characters —
        alphanumeric, underscores, and hyphens.  Rejects anything that could
        be used for injection or path traversal.

        Args:
            setting_id: The setting identifier string to check.

        Returns:
            True if the ID matches the safe pattern, False otherwise.
        """
        import re  # Imported locally to keep the module-level namespace minimal

        return bool(re.match(r"^[a-zA-Z0-9_-]+$", setting_id))

    # ---------------------------------------------------------------------------
    # Preset retrieval
    # ---------------------------------------------------------------------------

    def get_preset_info(self, preset_id: str) -> Optional[Dict[str, Any]]:
        """
        Retrieve metadata about a preset without returning the full settings
        dictionary.  Useful for displaying preset cards/lists in the UI.

        Args:
            preset_id: The unique identifier of the preset.

        Returns:
            A dictionary with keys 'id', 'name', 'description', 'icon',
            'version', 'setting_count', and 'tags', or None if not found.
        """
        if preset_id not in self.presets:
            return None

        # The full preset data dictionary
        data = self.presets[preset_id]

        # Return a lightweight metadata dictionary (no raw settings payload)
        return {
            "id": preset_id,
            "name": data.get("name", "Unknown"),
            "description": data.get("description", ""),
            "icon": data.get("icon", "⚙️"),
            "version": data.get("version", "1.0"),
            "setting_count": len(data.get("settings", {})),
            "tags": data.get("tags", []),
        }

    def get_preset_settings(self, preset_id: str) -> Dict[str, Any]:
        """
        Get the full settings dictionary for a preset — the actual
        setting_id -> value mappings that will be applied to the system.

        Args:
            preset_id: The unique identifier of the preset.

        Returns:
            A dictionary of {setting_id: value} pairs, or an empty dict if
            the preset is not found.
        """
        if preset_id not in self.presets:
            return {}

        return self.presets[preset_id].get("settings", {})

    def list_presets(self) -> List[Dict[str, Any]]:
        """
        List all available presets, sorted alphabetically by display name.

        Returns:
            A list of preset metadata dictionaries (see get_preset_info).
        """
        presets_list = []
        for preset_id in self.presets:
            info = self.get_preset_info(preset_id)
            if info:
                presets_list.append(info)

        # Sort the list alphabetically by the 'name' field for consistent UI ordering
        presets_list.sort(key=lambda x: x["name"])
        return presets_list

    # ---------------------------------------------------------------------------
    # Preset application
    # ---------------------------------------------------------------------------

    def apply_preset(
        self, preset_id: str, apply_function
    ) -> Tuple[bool, Dict[str, bool]]:
        """
        Apply every setting in a preset by calling the provided callback for
        each one.  This decouples the preset data from the actual system
        modification logic (registry writes, service toggles, etc.).

        Args:
            preset_id:      The ID of the preset to apply.
            apply_function: A callable that accepts (setting_id: str, value: Any)
                            and returns True on success or False on failure.

        Returns:
            A tuple of (overall_success, per_setting_results) where
            overall_success is True only if *every* setting applied successfully,
            and per_setting_results maps each setting_id to its success boolean.
        """
        if preset_id not in self.presets:
            return False, {}  # Preset doesn't exist

        # Fetch the setting_id -> value mapping for this preset
        settings = self.get_preset_settings(preset_id)

        # Accumulates per-setting success/failure results
        results: Dict[str, bool] = {}

        # Track whether all settings applied without error
        all_success = True

        for setting_id, value in settings.items():
            try:
                # Invoke the caller-provided apply callback
                success = apply_function(setting_id, value)
                results[setting_id] = success
                if not success:
                    all_success = False
            except Exception as e:
                # Treat exceptions as failures but continue applying the rest
                results[setting_id] = False
                all_success = False
                print(f"Error applying setting {setting_id}: {e}")

        return all_success, results

    # ---------------------------------------------------------------------------
    # Preset creation and deletion
    # ---------------------------------------------------------------------------

    def create_preset(
        self,
        preset_id: str,
        name: str,
        description: str,
        settings: Dict[str, Any],
        icon: str = "⚙️",
    ) -> bool:
        """
        Create a new preset file on disk and reload the preset cache.

        Args:
            preset_id:   Unique identifier for the preset (lowercase alphanumeric,
                         hyphens, and underscores only).
            name:        Human-readable display name (truncated to 100 chars).
            description: Short description of what the preset configures
                         (truncated to 500 chars).
            settings:    Dictionary of {setting_id: value} pairs to store.
            icon:        Optional emoji icon (only the first 2 characters are kept).

        Returns:
            True if the preset file was written successfully, False on validation
            failure or I/O error.
        """
        # Validate the preset ID format
        if not preset_id or not re.match(r"^[a-z0-9_-]+$", preset_id, re.IGNORECASE):
            return False

        # Enforce a hard limit on the number of settings
        if len(settings) > 200:
            return False

        # Build the preset data dictionary with metadata
        preset_data = {
            "app": "WinSet",  # Application signature
            "name": name[:100],  # Truncate to max length
            "description": description[:500],  # Truncate to max length
            "icon": icon[:2] if icon else "⚙️",  # Keep first 2 chars (emoji)
            "version": "1.0",  # Schema version
            "created": datetime.now().isoformat(),  # ISO 8601 creation timestamp
            "settings": settings,  # The actual configuration values
        }

        # Write to the writable presets directory, falling back to the last
        # configured directory if writable_presets_dir is not set.
        target_dir = getattr(self, "writable_presets_dir", self.presets_dirs[-1])
        preset_path = os.path.join(target_dir, f"{preset_id}.preset.json")

        try:
            # Serialize the preset data as pretty-printed JSON
            with open(preset_path, "w", encoding="utf-8") as f:
                json.dump(preset_data, f, indent=2, ensure_ascii=False)

            # Refresh the in-memory preset cache so the new preset is available
            self._load_presets()
            return True

        except Exception as e:
            print(f"Error creating preset: {e}")
            return False

    def delete_preset(self, preset_id: str) -> bool:
        """
        Delete a preset file from disk and remove it from the in-memory cache.

        Built-in presets that live in read-only directories may fail to delete;
        those failures are silently ignored.

        Args:
            preset_id: The ID of the preset to delete.

        Returns:
            True if the preset file was found and deleted, False otherwise.
        """
        if preset_id not in self.presets:
            return False  # Preset not loaded

        # Track whether we actually deleted a file from any directory
        deleted = False

        # Search every configured directory for a matching preset file
        for directory in self.presets_dirs:
            preset_path = os.path.join(directory, f"{preset_id}.preset.json")
            if os.path.exists(preset_path):
                try:
                    os.remove(preset_path)
                    deleted = True
                except Exception:
                    # Silently ignore failures (e.g., read-only built-in dir)
                    pass

        # If at least one file was removed, drop it from the in-memory caches
        if deleted:
            if preset_id in self.presets:
                del self.presets[preset_id]
            return True
        return False

    # ---------------------------------------------------------------------------
    # Convenience accessors
    # ---------------------------------------------------------------------------

    def get_preset_list(self) -> List[str]:
        """
        Get a flat list of all loaded preset IDs.

        Returns:
            List of preset ID strings (keys of self.presets).
        """
        return list(self.presets.keys())

    def get_preset_usage(self, preset_id: str) -> Optional[int]:
        """
        Get the number of times a preset has been applied (usage count).

        Intended to integrate with HistoryManager for analytics; currently
        returns a placeholder value.

        Args:
            preset_id: The ID of the preset.

        Returns:
            The number of times the preset was applied, or None if not found.
        """
        # Placeholder — future integration with HistoryManager
        return 0
