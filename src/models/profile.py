from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional
import json
import hashlib

from .setting import Setting

@dataclass
class Profile:
    """Complete settings profile"""
    name: str
    created: datetime = field(default_factory=datetime.now)
    modified: datetime = field(default_factory=datetime.now)
    settings: Dict[str, Setting] = field(default_factory=dict)
    version: str = "1.0"
    windows_version: str = ""
    description: str = ""
    tags: List[str] = field(default_factory=list)
    
    @property
    def setting_count(self) -> int:
        return len(self.settings)
    
    @property
    def checksum(self) -> str:
        """Generate checksum for integrity verification"""
        data = json.dumps(self.export(), sort_keys=True)
        return hashlib.sha256(data.encode()).hexdigest()
    
    def add_setting(self, setting: Setting):
        """Add or update a setting"""
        self.settings[setting.id] = setting
        self.modified = datetime.now()
    
    def remove_setting(self, setting_id: str):
        """Remove a setting"""
        if setting_id in self.settings:
            del self.settings[setting_id]
            self.modified = datetime.now()
    
    def apply_all(self, safe_mode: bool = True) -> Dict[str, bool]:
        """Apply all settings to the system"""
        results = {}
        for setting_id, setting in self.settings.items():
            if safe_mode and setting.requires_restart:
                # Skip restart-required settings in safe mode
                results[setting_id] = False
                continue
            
            try:
                results[setting_id] = setting.apply()
            except Exception as e:
                results[setting_id] = False
                # Log error
        return results
    
    def export(self) -> dict:
        """Export profile to dictionary"""
        return {
            "name": self.name,
            "created": self.created.isoformat(),
            "modified": self.modified.isoformat(),
            "version": self.version,
            "windows_version": self.windows_version,
            "description": self.description,
            "tags": self.tags,
            "settings": {
                sid: s.export() for sid, s in self.settings.items()
            },
            "checksum": self.checksum
        }
    
    @classmethod
    def import_from_dict(cls, data: dict) -> 'Profile':
        """Create profile from dictionary"""
        profile = cls(
            name=data["name"],
            created=datetime.fromisoformat(data["created"]),
            modified=datetime.fromisoformat(data["modified"]),
            version=data["version"],
            windows_version=data["windows_version"],
            description=data.get("description", ""),
            tags=data.get("tags", [])
        )
        
        # Reconstruct settings (simplified - real version would recreate objects)
        for sid, sdata in data.get("settings", {}).items():
            # This would need proper reconstruction based on type
            pass
            
        return profile