import os
import glob
import json
import logging
from typing import List, Dict, Tuple, Optional
from src.storage.importer import ProfileImporter
from src.models.profile import Profile
from src.models.setting import RegistrySetting

# Configure logging
logger = logging.getLogger(__name__)


class PresetManager:
    """Discovers, loads, and manages default configuration presets."""

    def __init__(self, presets_dir: str = None):
        """
        Initialize the preset manager.
        Args:
            presets_dir: Path to the directory containing .json preset files.
                         Defaults to the 'presets' folder in the project root.
        """
        if presets_dir is None:
            # Assume presets are loaded from a folder adjacent to main executable/script
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            self.presets_dir = os.path.join(base_dir, "presets")
        else:
            self.presets_dir = presets_dir

        self.importer = ProfileImporter()
        self.available_presets: Dict[str, str] = {} # Map 'name' -> 'filepath'
        self.custom_presets_dir = os.path.join(self.presets_dir, "custom")
        os.makedirs(self.custom_presets_dir, exist_ok=True)
        self.refresh_presets()
        logger.info(f"Preset manager initialized with presets from {self.presets_dir}")

    def refresh_presets(self):
        """Scans the presets directory for available JSON files."""
        self.available_presets.clear()
        
        if not os.path.exists(self.presets_dir):
            logger.warning(f"Presets directory not found: {self.presets_dir}")
            return

        # Load built-in presets
        json_files = glob.glob(os.path.join(self.presets_dir, "*.json"))
        
        for file_path in json_files:
            # Skip custom presets directory
            if "custom" in file_path:
                continue
                
            # Load the file just to extract its metadata name securely without applying it
            try:
                success, msg, profile = self.importer.load_profile(file_path)
                if success and profile:
                    # We store it by lowercase internal name for easier lookup
                    preset_id = os.path.basename(file_path).replace('.json', '').lower()
                    self.available_presets[preset_id] = file_path
                    logger.debug(f"Loaded built-in preset: {preset_id}")
            except Exception as e:
                logger.error(f"Failed to parse preset file {file_path}: {e}")
        
        # Load custom presets
        custom_files = glob.glob(os.path.join(self.custom_presets_dir, "*.json"))
        
        for file_path in custom_files:
            try:
                success, msg, profile = self.importer.load_profile(file_path)
                if success and profile:
                    preset_id = os.path.basename(file_path).replace('.json', '').lower()
                    # Mark as custom preset
                    custom_preset_id = f"custom_{preset_id}"
                    self.available_presets[custom_preset_id] = file_path
                    logger.debug(f"Loaded custom preset: {custom_preset_id}")
            except Exception as e:
                logger.error(f"Failed to parse custom preset file {file_path}: {e}")

    def get_preset_list(self) -> List[str]:
        """Returns a list of available preset identifiers."""
        return list(self.available_presets.keys())

    def get_preset_info(self, preset_id: str) -> Optional[Dict[str, Any]]:
        """Get information about a preset without loading the full profile.
        
        Args:
            preset_id: The identifier of the preset.
            
        Returns:
            Dictionary with preset info or None if not found.
        """
        preset_id = preset_id.lower()
        if preset_id not in self.available_presets:
            return None
            
        file_path = self.available_presets[preset_id]
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            is_custom = preset_id.startswith("custom_")
            
            return {
                'id': preset_id,
                'name': data.get('name', preset_id),
                'description': data.get('description', ''),
                'version': data.get('version', '1.0'),
                'is_custom': is_custom,
                'file_path': file_path
            }
        except Exception as e:
            logger.error(f"Failed to read preset info for {preset_id}: {e}")
            return None

    def load_preset(self, preset_id: str) -> Tuple[bool, str, Profile | None]:
        """Loads a specific preset profile into memory."""
        preset_id = preset_id.lower()
        if preset_id not in self.available_presets:
            return False, f"Preset '{preset_id}' not found.", None
            
        file_path = self.available_presets[preset_id]
        return self.importer.load_profile(file_path)

    def apply_preset(self, preset_id: str, safe_mode: bool = True) -> Tuple[bool, str, Dict[str, bool]]:
        """Directly loads and applies a preset to the system."""
        success, msg, profile = self.load_preset(preset_id)
        if not success:
            return False, msg, {}
            
        results = self.importer.apply_profile(profile, safe_mode=safe_mode)
        
        # Check if any settings applied successfully
        if any(results.values()):
            return True, f"Preset '{preset_id}' applied successfully.", results
        else:
            return False, f"Failed to apply settings from preset '{preset_id}'.", results

    def create_custom_preset(
        self, 
        name: str, 
        description: str = "", 
        settings: Optional[List[RegistrySetting]] = None
    ) -> Tuple[bool, str, Optional[str]]:
        """Create a new custom preset from current settings.
        
        Args:
            name: Name of the new preset.
            description: Description of the preset.
            settings: List of settings to include (if None, will use current system values).
            
        Returns:
            Tuple of (success, message, preset_id).
        """
        try:
            # Create a sanitized preset ID
            preset_id = name.lower().replace(" ", "_")
            file_name = f"{preset_id}.json"
            file_path = os.path.join(self.custom_presets_dir, file_name)
            
            # Check if preset already exists
            if os.path.exists(file_path):
                return False, f"Custom preset '{name}' already exists.", None
            
            # Create profile
            profile = Profile(
                name=name,
                description=description,
                windows_version="Custom Preset"
            )
            
            # If settings are provided, use them
            if settings:
                for setting in settings:
                    profile.add_setting(setting)
            
            # Export to file
            profile_data = profile.export()
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(profile_data, f, indent=4)
            
            # Refresh presets to include the new one
            self.refresh_presets()
            
            custom_preset_id = f"custom_{preset_id}"
            logger.info(f"Created custom preset: {custom_preset_id}")
            
            return True, f"Custom preset '{name}' created successfully.", custom_preset_id
            
        except Exception as e:
            logger.error(f"Failed to create custom preset: {e}")
            return False, f"Error creating custom preset: {str(e)}", None

    def delete_custom_preset(self, preset_id: str) -> Tuple[bool, str]:
        """Delete a custom preset.
        
        Args:
            preset_id: The identifier of the preset to delete.
            
        Returns:
            Tuple of (success, message).
        """
        preset_id = preset_id.lower()
        
        if not preset_id.startswith("custom_"):
            return False, "Cannot delete built-in presets."
        
        if preset_id not in self.available_presets:
            return False, f"Preset '{preset_id}' not found."
        
        try:
            file_path = self.available_presets[preset_id]
            os.remove(file_path)
            
            # Remove from available presets
            del self.available_presets[preset_id]
            
            logger.info(f"Deleted custom preset: {preset_id}")
            return True, f"Custom preset '{preset_id}' deleted successfully."
            
        except Exception as e:
            logger.error(f"Failed to delete custom preset {preset_id}: {e}")
            return False, f"Error deleting custom preset: {str(e)}"

    def update_custom_preset(
        self, 
        preset_id: str, 
        name: Optional[str] = None, 
        description: Optional[str] = None, 
        settings: Optional[List[RegistrySetting]] = None
    ) -> Tuple[bool, str]:
        """Update an existing custom preset.
        
        Args:
            preset_id: The identifier of the preset to update.
            name: New name for the preset (optional).
            description: New description (optional).
            settings: Updated list of settings (optional).
            
        Returns:
            Tuple of (success, message).
        """
        preset_id = preset_id.lower()
        
        if not preset_id.startswith("custom_"):
            return False, "Cannot update built-in presets."
        
        if preset_id not in self.available_presets:
            return False, f"Preset '{preset_id}' not found."
        
        try:
            file_path = self.available_presets[preset_id]
            
            # Load existing preset
            success, msg, profile = self.load_preset(preset_id)
            if not success:
                return False, f"Failed to load existing preset: {msg}"
            
            # Update fields if provided
            if name:
                profile.name = name
            if description is not None:
                profile.description = description
            if settings is not None:
                profile.settings.clear()
                for setting in settings:
                    profile.add_setting(setting)
            
            # Save updated preset
            profile_data = profile.export()
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(profile_data, f, indent=4)
            
            logger.info(f"Updated custom preset: {preset_id}")
            return True, f"Custom preset '{preset_id}' updated successfully."
            
        except Exception as e:
            logger.error(f"Failed to update custom preset {preset_id}: {e}")
            return False, f"Error updating custom preset: {str(e)}"

    def export_preset(self, preset_id: str, output_path: str) -> Tuple[bool, str]:
        """Export a preset to a custom location.
        
        Args:
            preset_id: The identifier of the preset to export.
            output_path: Destination path for the exported preset.
            
        Returns:
            Tuple of (success, message).
        """
        preset_id = preset_id.lower()
        
        if preset_id not in self.available_presets:
            return False, f"Preset '{preset_id}' not found."
        
        try:
            source_path = self.available_presets[preset_id]
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            
            # Copy the file
            import shutil
            shutil.copy2(source_path, output_path)
            
            logger.info(f"Exported preset {preset_id} to {output_path}")
            return True, f"Preset '{preset_id}' exported successfully to {output_path}."
            
        except Exception as e:
            logger.error(f"Failed to export preset {preset_id}: {e}")
            return False, f"Error exporting preset: {str(e)}"

    def import_preset(self, file_path: str, overwrite: bool = False) -> Tuple[bool, str, Optional[str]]:
        """Import a preset from a custom location.
        
        Args:
            file_path: Path to the preset file to import.
            overwrite: If True, overwrite existing preset with same name.
            
        Returns:
            Tuple of (success, message, preset_id).
        """
        try:
            if not os.path.exists(file_path):
                return False, f"File not found: {file_path}", None
            
            # Load the preset to get its name
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            preset_name = data.get('name', 'imported_preset')
            preset_id = preset_name.lower().replace(" ", "_")
            dest_file_name = f"{preset_id}.json"
            dest_path = os.path.join(self.custom_presets_dir, dest_file_name)
            
            # Check if preset already exists
            if os.path.exists(dest_path) and not overwrite:
                return False, f"Preset '{preset_name}' already exists. Use overwrite=True to replace.", None
            
            # Copy the file
            import shutil
            shutil.copy2(file_path, dest_path)
            
            # Refresh presets
            self.refresh_presets()
            
            custom_preset_id = f"custom_{preset_id}"
            logger.info(f"Imported preset: {custom_preset_id}")
            
            return True, f"Preset '{preset_name}' imported successfully.", custom_preset_id
            
        except Exception as e:
            logger.error(f"Failed to import preset from {file_path}: {e}")
            return False, f"Error importing preset: {str(e)}", None
