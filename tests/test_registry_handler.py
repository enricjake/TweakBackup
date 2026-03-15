"""
Tests for registry handler
"""

import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from core.registry_handler import RegistryHandler
from models.setting import RegistrySetting

class TestRegistryHandler:
    """Test RegistryHandler class"""
    
    def test_handler_initialization(self):
        """Test creating registry handler"""
        handler = RegistryHandler()
        assert handler is not None
        
    def test_hive_conversion(self):
        """Test hive string to constant conversion"""
        handler = RegistryHandler()
        
        # Test valid hives
        assert handler._get_hive_constant("HKEY_CURRENT_USER") is not None
        assert handler._get_hive_constant("HKEY_LOCAL_MACHINE") is not None
        assert handler._get_hive_constant("HKEY_CLASSES_ROOT") is not None
        
        # Test invalid hive
        with pytest.raises(ValueError):
            handler._get_hive_constant("INVALID_HIVE")
            
    def test_type_conversion(self):
        """Test registry type string to constant conversion"""
        handler = RegistryHandler()
        
        assert handler._get_type_constant("REG_SZ") == 1  # winreg.REG_SZ
        assert handler._get_type_constant("REG_DWORD") == 4  # winreg.REG_DWORD
        assert handler._get_type_constant("REG_BINARY") == 3  # winreg.REG_BINARY
        
        with pytest.raises(ValueError):
            handler._get_type_constant("INVALID_TYPE")
            
    def test_read_write_mock(self, mock_registry, monkeypatch):
        """Test registry operations with mock"""
        handler = RegistryHandler()
        
        # Mock the actual registry operations
        def mock_open_key(hive, key, reserved, access):
            return "mock_key"
            
        def mock_set_value(key, value_name, reserved, type_const, value):
            mock_registry.values[value_name] = value
            return True
            
        monkeypatch.setattr('winreg.OpenKey', mock_open_key)
        monkeypatch.setattr('winreg.SetValueEx', mock_set_value)
        
        # Test writing
        result = handler.write_value(
            "HKEY_CURRENT_USER",
            "Software\\Test",
            "TestValue",
            "REG_DWORD",
            42
        )
        assert result == True
        assert mock_registry.values["TestValue"] == 42