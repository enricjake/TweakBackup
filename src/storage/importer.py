import json  # Used to parse JSON profile files from disk and serialize data for checksum verification
import os  # Used to check whether a profile file exists on disk before attempting to read it
import hashlib  # Used to compute SHA-256 checksums for profile integrity verification
import base64  # Used to decode base64-encoded binary registry values stored in profiles
from typing import (
    Dict,
    Any,
    Tuple,
)  # Type hints: Dict for result maps, Any for flexible value types, Tuple for multi-value returns

from src.models.profile import (
    Profile,
)  # Data model representing a collection of settings (a configuration profile)
from src.models.setting import (
    RegistrySetting,
    PowerSetting,
    ServiceSetting,
    SettingCategory,
    SettingType,
)  # Setting models and enums for different Windows configuration categories
from src.core.registry_handler import (
    RegistryHandler,
)  # Low-level handler that reads/writes values to the Windows Registry


class ProfileImporter:
    """Imports and applies settings from a JSON profile file."""

    def __init__(self):
        # RegistryHandler: interface to the Windows Registry; used during apply_profile() to write registry values
        self.registry = RegistryHandler()
        # Lazy-import SettingLoader to avoid circular imports at module load time
        from src.core.setting_loader import SettingLoader

        # SettingLoader: provides access to the master list of known settings (categories, defaults, metadata)
        self.loader = SettingLoader()
        # _master_settings: dict mapping lowercase setting names and IDs to their canonical Setting objects.
        # Used during profile loading to "hydrate" simplified preset entries that only store an ID/name and value,
        # filling in the rest of the metadata (hive, key_path, value_type, etc.) from the master definitions.
        self._master_settings = {}
        # Flatten settings from categories for easy lookup
        for cat in (
            self.loader.get_categories()
        ):  # cat: a SettingCategory enum value representing a group of related settings
            for s in self.loader.get_settings_for_category(
                cat
            ):  # s: a Setting object (e.g., RegistrySetting) within that category
                self._master_settings[s.name.lower()] = (
                    s  # Index by lowercase human-readable name for lookup
                )
                self._master_settings[s.id.lower()] = (
                    s  # Index by lowercase unique ID for lookup
                )

    def _verify_checksum(self, data: dict) -> bool:
        """Verifies the integrity of the profile data using the checksum."""
        # stored_checksum: the SHA-256 hex digest that was embedded in the profile file at export time
        stored_checksum = data.get("checksum")
        if not stored_checksum:
            return False  # No checksum present; cannot verify integrity, so reject the profile

        # Build a sanitized copy to avoid mutating caller-owned dictionaries.
        # data_without_checksum: the profile dict with the "checksum" key removed so we can recompute the hash
        data_without_checksum = {k: v for k, v in data.items() if k != "checksum"}
        # data_str: deterministic JSON string of the profile data (sort_keys ensures consistent ordering)
        data_str = json.dumps(data_without_checksum, sort_keys=True)
        # calculated_checksum: SHA-256 hex digest recomputed from the sanitized data
        calculated_checksum = hashlib.sha256(data_str.encode()).hexdigest()
        # Return True only if the stored checksum matches the recomputed one (profile has not been tampered with)
        return stored_checksum == calculated_checksum

    def _deserialize_registry_value(self, setting_value: Any, value_type: str) -> Any:
        """Restore encoded profile values back to native registry-compatible types."""
        # Handles the special case where binary (REG_BINARY) values were base64-encoded during profile export.
        # setting_value: the raw value from the JSON (may be a dict with __encoding__="base64" for binary data)
        # value_type: the Windows Registry value type string (e.g., "REG_BINARY", "REG_DWORD", "REG_SZ")
        if (
            value_type == "REG_BINARY"  # Only decode if the type is binary
            and isinstance(
                setting_value, dict
            )  # And the value was stored as a structured dict
            and setting_value.get("__encoding__")
            == "base64"  # With a base64 encoding marker
            and isinstance(
                setting_value.get("data"), str
            )  # And the actual encoded data is a string
        ):
            try:
                return base64.b64decode(
                    setting_value["data"]
                )  # Decode the base64 string back to raw bytes
            except (ValueError, TypeError):
                return setting_value  # If decoding fails, return the original value unchanged
        return setting_value  # Non-binary or non-encoded values pass through unchanged

    def load_profile(self, file_path: str) -> Tuple[bool, str, Profile | None]:
        """Loads and validates a profile from disk."""
        # file_path: absolute or relative path to a WinSet JSON profile file on disk
        if not os.path.exists(file_path):
            return (
                False,
                f"File not found: {file_path}",
                None,
            )  # Early exit if the file doesn't exist

        try:
            # Open the profile file with UTF-8 encoding (standard for WinSet JSON exports)
            with open(file_path, "r", encoding="utf-8") as f:
                # data: the parsed JSON dictionary containing profile metadata and settings
                data = json.load(f)

            # If the profile contains a checksum, verify it hasn't been tampered with or corrupted
            if "checksum" not in data:
                return (
                    False,
                    "Profile missing integrity checksum. Only verified profiles can be imported.",
                    None,
                )

            if not self._verify_checksum(data):
                return (
                    False,
                    "Profile integrity check failed. The file may have been tampered with or corrupted.",
                    None,
                )

            # Reconstruct the profile object from the top-level metadata (name, description, version, etc.)
            profile = Profile.import_from_dict(data)

            # settings_data: dict mapping setting IDs to their serialized setting data (rich or simplified format)
            settings_data = data.get("settings", {})
            for (
                sid,
                sdata,
            ) in (
                settings_data.items()
            ):  # sid: the setting's unique ID string; sdata: its serialized data
                setting = None  # Will hold the reconstructed Setting object if parsing succeeds

                # Check if this is a standard rich setting (has a "type" field with full metadata)
                # or a simplified preset setting (only stores ID/name and value)
                is_rich = isinstance(sdata, dict) and "type" in sdata

                if is_rich:
                    # stype: the SettingType enum value string (e.g., "REGISTRY") identifying the setting's category
                    stype = sdata.get("type")
                    if stype == SettingType.REGISTRY.value:
                        # Reconstruct a full RegistrySetting from the rich JSON data
                        setting = RegistrySetting(
                            id=sdata["id"],  # Unique identifier for the setting
                            name=sdata["name"],  # Human-readable display name
                            description=sdata[
                                "description"
                            ],  # Explanation of what the setting controls
                            category=SettingCategory(
                                sdata["category"]
                            ),  # Enum grouping (e.g., APPEARANCE, SECURITY)
                            setting_type=SettingType(
                                stype
                            ),  # The type of setting (REGISTRY, POWER, SERVICE, etc.)
                            value=self._deserialize_registry_value(
                                sdata["value"], sdata["value_type"]
                            ),  # The configured value, deserialized if binary
                            default_value=sdata.get(
                                "default_value"
                            ),  # The factory/default value for reset purposes
                            requires_admin=sdata.get(
                                "requires_admin", False
                            ),  # Whether applying needs admin elevation
                            requires_restart=sdata.get(
                                "requires_restart", False
                            ),  # Whether applying needs a system restart
                            hive=sdata[
                                "hive"
                            ],  # Registry hive (e.g., "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE")
                            key_path=sdata[
                                "key_path"
                            ],  # Full registry key path (e.g., "Software\\Microsoft\\Windows")
                            value_name=sdata[
                                "value_name"
                            ],  # The name of the value within the registry key
                            value_type=sdata[
                                "value_type"
                            ],  # Registry value type string (e.g., "REG_DWORD", "REG_SZ")
                            is_expanded=sdata.get(
                                "is_expanded", False
                            ),  # UI state: whether the setting's details are expanded
                        )
                    # ... other types (Power, Service) could be added here
                else:
                    # Hydrate from master settings if only ID/Name and Value provided (simplified preset format)
                    # lookup_key: the setting ID converted to lowercase for case-insensitive matching
                    lookup_key = sid.lower()
                    # master: the canonical Setting object from the loader that matches this ID or name
                    master = self._master_settings.get(
                        lookup_key
                    ) or self._master_settings.get(lookup_key.replace("_", " "))

                    if master:
                        # value: the user's configured value; extracted from dict form or used as-is if a bare value
                        value = sdata.get("value") if isinstance(sdata, dict) else sdata
                        if isinstance(master, RegistrySetting):
                            # Deserialize binary values before constructing the new setting
                            value = self._deserialize_registry_value(
                                value, master.value_type
                            )
                        if isinstance(master, RegistrySetting):
                            # Build a new RegistrySetting using master metadata but with the profile's value
                            setting = RegistrySetting(
                                id=master.id,  # Canonical ID from the master definition
                                name=master.name,  # Canonical display name
                                description=master.description,  # Canonical description
                                category=master.category,  # Canonical category enum
                                setting_type=master.setting_type,  # Canonical setting type enum
                                value=value,  # The user's configured value (possibly deserialized)
                                default_value=master.default_value,  # Default value from master
                                hive=master.hive,  # Registry hive from master
                                key_path=master.key_path,  # Registry key path from master
                                value_name=master.value_name,  # Registry value name from master
                                value_type=master.value_type,  # Registry value type from master
                            )

                if setting:
                    profile.add_setting(
                        setting
                    )  # Attach the reconstructed setting to the profile

            return (
                True,
                "Profile loaded successfully.",
                profile,
            )  # Success: return the fully populated profile

        except Exception as e:
            return (
                False,
                f"Error loading profile: {str(e)}",
                None,
            )  # Catch-all for JSON parse errors, missing keys, etc.

    def apply_profile(
        self, profile: Profile, safe_mode: bool = True
    ) -> Dict[str, bool]:
        """Applies a loaded profile to the system."""
        # profile: a fully loaded Profile object containing the settings to apply
        # safe_mode: when True, skips settings that require a system restart to avoid disruptive changes
        # results: dict mapping each setting ID to a boolean indicating whether it was applied successfully
        results = {}
        for (
            setting_id,
            setting,
        ) in (
            profile.settings.items()
        ):  # setting_id: the unique string ID; setting: the Setting object
            if safe_mode and setting.requires_restart:
                # In safe mode, mark restart-requiring settings as skipped (False) without applying them
                results[setting_id] = False
                continue

            if isinstance(setting, RegistrySetting):
                try:
                    self.registry._validate_key_path(setting.key_path)
                except ValueError:
                    results[setting_id] = False
                    continue
                success = self.registry.write_value(
                    hive=setting.hive,  # Target registry hive (e.g., HKEY_CURRENT_USER)
                    key_path=setting.key_path,  # Path to the registry key
                    value_name=setting.value_name,  # Name of the value to write
                    value_type=setting.value_type,  # Type of the value (REG_DWORD, REG_SZ, etc.)
                    value=setting.value,  # The actual data to write
                )
                results[setting_id] = success  # Record whether the write succeeded
                setting.is_applied = (
                    success  # Update the setting object's applied state for UI tracking
                )
            else:
                # Fallback for future non-registry settings (PowerSetting, ServiceSetting, etc.)
                try:
                    results[setting_id] = (
                        setting.apply()
                    )  # Call the setting's own apply() method
                except Exception:
                    results[setting_id] = False  # If apply() raises, record failure

        return results  # Return the full success/failure map for all settings in the profile
