import json
import os
import platform
from typing import List

from src.models.profile import Profile
from src.models.setting import RegistrySetting
from src.core.registry_handler import RegistryHandler

class ProfileExporter:
    """Exports current system settings to a JSON profile file."""

    def __init__(self):
        """Initialize the exporter with a RegistryHandler."""
        self.registry = RegistryHandler()

    def _get_windows_version(self) -> str:
        """Retrieves the Windows OS version."""
        return f"{platform.system()} {platform.release()} ({platform.version()})"

    def export_profile(self, settings_to_export: List[RegistrySetting], output_path: str, profile_name: str = "Exported Profile", description: str = "") -> bool:
        """Reads current system values for provided settings and saves them to a file."""
        try:
            # Create a new Profile object with metadata
            profile = Profile(
                name=profile_name,
                description=description,
                windows_version=self._get_windows_version()
            )

            # Loop through the requested settings to capture
            for setting in settings_to_export:
                # Read the current value directly from the Windows Registry
                current_value = self.registry.read_value(
                    hive=setting.hive,
                    key_path=setting.key_path,
                    value_name=setting.value_name
                )
                
                # If the registry value exists, update the setting and add it
                if current_value is not None:
                    # Convert bytes to string if needed for JSON serialization
                    if isinstance(current_value, bytes):
                        current_value = current_value.decode('utf-8', errors='replace')
                    setting.value = current_value
                    profile.add_setting(setting)

            # Ensure the output directory exists before saving
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

            # Write the exported profile dictionary (which includes the checksum logic) to JSON
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(profile.export(), f, indent=4)

            return True

        except Exception as e:
            print(f"Error exporting profile: {e}")
            return False
