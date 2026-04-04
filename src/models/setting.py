"""
Setting models for WinSet

This module defines the data models for all configurable settings in the
WinSet Windows Configuration Toolkit. It provides a type-safe, extensible
framework for representing, validating, applying, and exporting Windows
settings such as registry values, power plans, and service configurations.

Key design points:
- All settings inherit from the base `Setting` dataclass.
- `SettingType` classifies how a setting is applied (registry edit, file
  modification, system call, or power-plan change).
- `SettingCategory` groups settings by user-facing area (appearance,
  privacy, taskbar, etc.) for UI organisation.
- Concrete subclasses (`RegistrySetting`, `PowerSetting`, `ServiceSetting`)
  override `validate()`, `apply()`, and `export()` to implement
  type-specific behaviour.
"""

from dataclasses import dataclass, field, KW_ONLY
from typing import Any, Optional, Dict
from enum import Enum
from datetime import datetime

import winreg


# ---------------------------------------------------------------------------
# Enums
# ---------------------------------------------------------------------------


class SettingType(Enum):
    """Type of setting — determines *how* the setting is persisted and applied.

    Attributes:
        REGISTRY:  Setting is stored in the Windows Registry and applied by
                   writing a value to a specific registry key path.
        FILE:      Setting is controlled by modifying a configuration file on
                   disk (e.g. INI, JSON, or XML files).
        SYSTEM:    Setting is applied via a system-level call such as a
                   PowerShell command or Win32 API invocation (e.g. disabling
                   a Windows service).
        POWER:     Setting relates to a Windows Power Plan and is applied by
                   activating a specific power-plan GUID via `powercfg`.
    """

    REGISTRY = "registry"
    FILE = "file"
    SYSTEM = "system"
    POWER = "power"


class SettingCategory(Enum):
    """Category of setting — used for grouping in the WinSet UI and export.

    Each value corresponds to a logical area of Windows configuration that
    users can toggle or customise.

    Attributes:
        APPEARANCE:    Visual theming, colours, transparency, fonts, and
                       other cosmetic system-wide tweaks.
        FILE_EXPLORER: Options inside Windows/File Explorer such as showing
                       hidden files, file extensions, or the ribbon UI.
        TASKBAR:       Taskbar position, grouping, icons, notification area,
                       and related tweaks.
        POWER:         Active power plan selection and sleep/hibernate timers.
        PRIVACY:       Telemetry, advertising ID, diagnostic data, location,
                       and other privacy-related toggles.
        SYSTEM:        General OS behaviour — UAC level, automatic updates,
                       service startup types, etc.
        KEYBOARD:      Keyboard repeat rate, sticky keys, filter keys, and
                       input-method preferences.
        MOUSE:         Pointer speed, scroll direction, double-click interval,
                       and button-swap settings.
        NETWORK:       Network profile type (public/private), proxy settings,
                       firewall rules, and DNS preferences.
    """

    APPEARANCE = "appearance"
    FILE_EXPLORER = "file_explorer"
    TASKBAR = "taskbar"
    POWER = "power"
    PRIVACY = "privacy"
    SYSTEM = "system"
    KEYBOARD = "keyboard"
    MOUSE = "mouse"
    NETWORK = "network"


# ---------------------------------------------------------------------------
# Base Setting
# ---------------------------------------------------------------------------


@dataclass
class Setting:
    """Base class for all settings in the WinSet toolkit.

    This dataclass captures the common metadata that every setting shares,
    regardless of whether it is applied via the registry, a power plan, a
    service command, or a file edit.  Concrete subclasses override
    `validate`, `apply`, and `export` to implement type-specific logic.

    Attributes:
        id:               Unique identifier for this setting (e.g.
                           `"appearance.dark_mode"`).  Used as a dictionary
                           key and for cross-referencing in profiles.
        name:             Human-readable display name shown in the WinSet UI
                           (e.g. `"Enable Dark Mode"`).
        description:      Longer explanation of what the setting does,
                           shown in tooltips or detail panels.
        category:         Which `SettingCategory` this setting belongs to —
                           controls where it appears in the UI tree.
        setting_type:     Which `SettingType` governs how the setting is
                           persisted (registry, file, system, or power).
        value:            The *current* value of this setting.  Its concrete
                           type depends on the subclass (int, str, list, etc.).
        default_value:    The factory / out-of-box value; used when the user
                           clicks "Reset to default".
        is_applied:       `True` if the setting has been successfully written
                           to the system.  Defaults to `False` on creation.
        requires_admin:   `True` if applying this setting requires elevation
                           (administrator privileges).
        requires_restart: `True` if the system or explorer must be restarted
                           for the change to take effect.
        last_modified:    Timestamp of the most recent successful `apply()`.
                          `None` if the setting has never been applied.
    """

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
        """Validate if a value is acceptable for this setting.

        Each subclass implements its own validation rules (e.g. range checks
        for DWORD values, type checks for strings).  The base implementation
        always raises `NotImplementedError` to force subclasses to override.

        Args:
            value: The candidate value to validate.

        Returns:
            `True` if the value is acceptable, `False` otherwise.

        Raises:
            NotImplementedError: If the subclass has not overridden this method.
        """
        raise NotImplementedError(
            f"Validate not implemented for {self.__class__.__name__}"
        )

    def apply(self) -> bool:
        """Apply this setting to the system.

        Subclasses implement the actual write logic (registry write, PowerShell
        command, file modification, etc.).  On success the implementation
        should set `self.is_applied = True` and update `self.last_modified`.

        Returns:
            `True` if the setting was applied successfully, `False` otherwise.

        Raises:
            NotImplementedError: If the subclass has not overridden this method.
        """
        raise NotImplementedError(
            f"Apply not implemented for {self.__class__.__name__}"
        )

    def export(self) -> dict:
        """Convert this setting to a JSON-serializable dictionary.

        The returned dict is suitable for saving to a profile file (JSON) or
        sending over an API.  Subclasses call `super().export()` and merge in
        their own additional fields.

        Returns:
            A plain dictionary containing id, name, description, value,
            category, type, and privilege/restart requirements.
        """
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "value": self.value,
            "category": self.category.value,
            "type": self.setting_type.value,
            "requires_admin": self.requires_admin,
            "requires_restart": self.requires_restart,
        }


# ---------------------------------------------------------------------------
# Registry Setting
# ---------------------------------------------------------------------------


@dataclass
class RegistrySetting(Setting):
    r"""Setting stored in the Windows Registry.

    Represents a single registry value that WinSet can read, validate, and
    write.  The `hive`, `key_path`, `value_name`, and `value_type` fields
    fully qualify the target location in the registry.

    Examples:
        A DWORD toggle under HKCU:
            RegistrySetting(
                id="appearance.dark_mode",
                name="Dark Mode",
                ...,
                hive="HKEY_CURRENT_USER",
                key_path=r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                value_name="AppsUseLightTheme",
                value_type="REG_DWORD",
                value=0,
                default_value=1,
            )
    """

    # KW_ONLY sentinel: all fields declared *below* this line are keyword-only.
    # This is required because the base Setting class has optional fields
    # (e.g. last_modified) that would otherwise conflict with these
    # required (no-default) fields in the subclass.
    _: KW_ONLY

    hive: str  # Registry hive root, e.g. "HKEY_CURRENT_USER" (HKCU) or
    # "HKEY_LOCAL_MACHINE" (HKLM).  Passed directly to
    # `RegistryHandler.write_value()`.

    key_path: str  # Full path beneath the hive to the containing key,
    # e.g. r"Software\Microsoft\Windows\CurrentVersion\Themes".
    # Backslashes separate sub-keys; a raw string is preferred.

    value_name: str  # Name of the value *within* the key, e.g. "AppsUseLightTheme".
    # Use an empty string "" for the key's (Default) value.

    value_type: str  # Registry data type as a string constant matching the
    # winreg constants: "REG_SZ", "REG_DWORD", "REG_QWORD",
    # "REG_BINARY", "REG_MULTI_SZ", or "REG_EXPAND_SZ".
    # Determines how `validate()` checks the value and how
    # `RegistryHandler` serialises it.

    # --- Optional fields with defaults ---

    is_expanded: bool = False  # When `True`, the value is treated as
    # REG_EXPAND_SZ so that environment-variable
    # placeholders like %SYSTEMROOT% are expanded
    # at read time by the OS.

    options: Optional[Dict[str, Any]] = None  # Optional mapping of display-label
    # → underlying value for settings exposed as a
    # dropdown in the UI (e.g. {"On": 1, "Off": 0}).
    # `None` when the setting uses free-form input.

    is_range: bool = False  # When `True`, the UI should render a slider /
    # range control instead of a free-text field.
    # The `options` dict (if present) may then carry
    # "min" and "max" keys.

    def __post_init__(self):
        """Validate and normalise fields after dataclass initialisation.

        Ensures that `setting_type` is always `SettingType.REGISTRY`, even if
        the caller accidentally passed a different type.  Uses
        `object.__setattr__` because dataclass fields are frozen after init.
        """
        # Set the setting_type automatically if not already set
        if self.setting_type == SettingType.REGISTRY:
            pass  # Already correct
        else:
            # Force to registry type — this avoids subtle bugs where a
            # RegistrySetting is accidentally constructed with a different
            # SettingType.
            object.__setattr__(self, "setting_type", SettingType.REGISTRY)

    def validate(self, value: Any) -> bool:
        """Validate a value against the expected registry data type.

        Each registry type has different constraints:
        - REG_DWORD:   unsigned 32-bit integer  (0 – 4 294 967 295)
        - REG_QWORD:   unsigned 64-bit integer  (0 – 18 446 744 073 709 551 615)
        - REG_SZ / REG_EXPAND_SZ: non-binary string
        - REG_BINARY:  bytes or bytearray
        - REG_MULTI_SZ: list of strings

        Args:
            value: The candidate value to type-check.

        Returns:
            `True` if the value matches the expected type and range,
            `False` otherwise.
        """
        if self.value_type == "REG_DWORD":
            # DWORD is a 32-bit unsigned integer
            return isinstance(value, int) and 0 <= value <= 4294967295
        elif self.value_type == "REG_SZ" or self.value_type == "REG_EXPAND_SZ":
            # String types — value must be a Python str
            return isinstance(value, str)
        elif self.value_type == "REG_BINARY":
            # Raw binary blob
            return isinstance(value, (bytes, bytearray))
        elif self.value_type == "REG_MULTI_SZ":
            # Multi-string is stored as a list of individual strings
            return isinstance(value, list) and all(isinstance(s, str) for s in value)
        elif self.value_type == "REG_QWORD":
            # QWORD is a 64-bit unsigned integer
            return isinstance(value, int) and 0 <= value <= 18446744073709551615
        return True  # Unknown type — accept any value as a fallback

    def apply(self) -> bool:
        """Write this setting to the Windows Registry.

        Delegates the actual registry write to `RegistryHandler`, which
        handles privilege elevation, error handling, and logging.

        Returns:
            `True` if the value was written successfully, `False` otherwise.
        """
        # Local import to avoid circular dependency issues at module level.
        # RegistryHandler lives in src.core and may itself import from
        # src.models, so a top-level import could cause an import cycle.
        from src.core.registry_handler import RegistryHandler

        handler = RegistryHandler()
        success = handler.write_value(
            hive=self.hive,  # e.g. "HKEY_CURRENT_USER"
            key_path=self.key_path,  # e.g. r"Software\...\Themes"
            value_name=self.value_name,  # e.g. "AppsUseLightTheme"
            value_type=self.value_type,  # e.g. "REG_DWORD"
            value=self.value,  # the actual data to write
        )
        return success

    def export(self) -> dict:
        """Export this registry setting to a plain dictionary.

        Extends the base `Setting.export()` output with registry-specific
        fields (hive, key_path, value_name, value_type, is_expanded) so
        the profile JSON fully captures how to re-apply the setting.

        Returns:
            A JSON-serializable dictionary.
        """
        base_dict = super().export()
        base_dict.update(
            {
                "hive": self.hive,
                "key_path": self.key_path,
                "value_name": self.value_name,
                "value_type": self.value_type,
                "is_expanded": self.is_expanded,
            }
        )
        return base_dict


# ---------------------------------------------------------------------------
# Power Setting
# ---------------------------------------------------------------------------


@dataclass
class PowerSetting(Setting):
    """Setting for Windows Power Plans.

    Represents the *active* power plan on the system.  Applying this setting
    switches the active plan to the one identified by `plan_guid` using the
    `powercfg /setactive` command via PowerShell.

    Attributes:
        plan_guid: The GUID of the target power plan, e.g.
                   `"8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"` for the
                   High Performance plan.  Obtained via `powercfg /list`.
    """

    _: KW_ONLY

    plan_guid: str  # GUID identifying the power plan to activate.
    # Windows ships with well-known GUIDs for Balanced,
    # Power Saver, and High Performance; custom plans have
    # user-generated GUIDs.

    def __post_init__(self):
        """Ensure `setting_type` is `SettingType.POWER` after init.

        Guarantees that even if a caller provides the wrong type, the
        instance is corrected before first use.
        """
        if self.setting_type != SettingType.POWER:
            object.__setattr__(self, "setting_type", SettingType.POWER)

    def apply(self) -> bool:
        """Activate the power plan identified by `plan_guid`.

        Uses `PowerShellHandler.set_power_plan()` which internally runs
        `powercfg /setactive <guid>`.

        Returns:
            `True` if the plan was activated successfully, `False` otherwise.
            Also updates `self.is_applied` to reflect the result.
        """
        # Local import to avoid potential circular imports at module load.
        from src.core.powershell_handler import PowerShellHandler

        ps = PowerShellHandler()
        success, _ = ps.set_power_plan(self.plan_guid)  # _ discards stdout/stderr text
        self.is_applied = success  # Track whether the change took effect
        return success


# ---------------------------------------------------------------------------
# Service Setting
# ---------------------------------------------------------------------------


@dataclass
class ServiceSetting(Setting):
    """Setting for Windows Services (start-up type configuration).

    Controls the startup behaviour of a Windows service identified by
    `service_name`.  Commonly used to disable unnecessary background
    services (e.g. telemetry, search indexer) for performance or privacy.

    Attributes:
        service_name:  The canonical Windows service name (not the display
                       name), e.g. `"DiagTrack"` for Connected User
                       Experiences and Telemetry.
        startup_type:  Desired start-up mode.  One of `"Disabled"`,
                       `"Manual"`, or `"Automatic"`.  Defaults to `"Disabled"`.
    """

    _: KW_ONLY

    service_name: str  # Internal service name as registered with the
    # Service Control Manager (SCM).  Must match
    # exactly (case-insensitive on Windows).

    startup_type: str = "Disabled"  # Desired startup type.  Accepted values:
    # - "Disabled"   — service cannot start
    # - "Manual"     — service starts on demand
    # - "Automatic"  — service starts at boot
    # Default is "Disabled" which is the most
    # common choice for WinSet's privacy/perf
    # profiles.

    def __post_init__(self):
        """Ensure `setting_type` is `SettingType.SYSTEM` after init.

        Service settings are classified as SYSTEM because they are applied
        via PowerShell commands rather than direct registry writes.
        """
        if self.setting_type != SettingType.SYSTEM:
            object.__setattr__(self, "setting_type", SettingType.SYSTEM)

    def apply(self) -> bool:
        """Apply the desired startup type to the Windows service.

        For `"Disabled"` startup type, delegates to
        `PowerShellHandler.disable_service()` which runs
        `Stop-Service` + `Set-Service -StartupType Disabled`.

        For other startup types (`"Manual"`, `"Automatic"`), constructs
        and runs a raw `Set-Service` PowerShell command.

        Returns:
            `True` if the service was configured successfully, `False` otherwise.
            Also updates `self.is_applied`.
        """
        # Local import to avoid potential circular imports at module load.
        from src.core.powershell_handler import PowerShellHandler

        ps = PowerShellHandler()
        if self.startup_type == "Disabled":
            # Use the dedicated helper which also stops a running service
            success, _ = ps.disable_service(self.service_name)
        else:
            # For Manual / Automatic, run a generic Set-Service command.
            # Add more service methods to PowerShellHandler if needed.
            command = f"Set-Service -Name '{self.service_name}' -StartupType {self.startup_type}"
            success, _ = ps.run_command(command)

        self.is_applied = success  # Track whether the change took effect
        return success
