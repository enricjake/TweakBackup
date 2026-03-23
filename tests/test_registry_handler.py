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

    def test_bulk_operations_mock(self, mock_registry, monkeypatch):
        """Test bulk registry operations with mock"""
        handler = RegistryHandler()

        # Mock the actual registry operations
        def mock_open_key(hive, key, reserved, access):
            return "mock_key"

        def mock_set_value(key, value_name, reserved, type_const, value):
            mock_registry.values[value_name] = value
            return True

        def mock_query_value(key, value_name):
            return mock_registry.values.get(value_name, None), 1  # 1 = REG_SZ

        monkeypatch.setattr('winreg.OpenKey', mock_open_key)
        monkeypatch.setattr('winreg.SetValueEx', mock_set_value)
        monkeypatch.setattr('winreg.QueryValueEx', mock_query_value)

        # Test bulk write
        operations = [
            ("HKEY_CURRENT_USER", "Software\\Test", "Value1", "REG_DWORD", 1),
            ("HKEY_CURRENT_USER", "Software\\Test", "Value2", "REG_SZ", "test"),
        ]

        results = handler.write_multiple_values(operations)
        assert len(results) == 2
        assert all(results.values())

                # Test bulk read
        read_results = handler.read_multiple_values(
            "HKEY_CURRENT_USER",
            "Software\\Test",
            ["Value1", "Value2", "Value3"]
        )
        assert len(read_results) == 3
        assert read_results["Value1"] == 1
        assert read_results["Value2"] == "test"
        assert read_results["Value3"] is None

    def test_key_operations_mock(self, monkeypatch):
        """Test key existence checking with mock"""
        handler = RegistryHandler()
        
                                        # Create a mock PyHKEY object that behaves like a real PyHKEY
        class MockPyHKEY:
            def __init__(self):
                self.closed = False
                
            def Close(self):
                self.closed = True
                return None
                
            def __enter__(self):
                return self
                
            def __exit__(self, exc_type, exc_val, exc_tb):
                self.Close()
                
        def mock_open_key_success(hive, key, reserved, access):
            return MockPyHKEY()
            
        def mock_open_key_fail(hive, key, reserved, access):
            raise FileNotFoundError("Key not found")
            
        # Mock CloseKey to accept our mock object
        def mock_close_key(key):
            if hasattr(key, 'Close'):
                key.Close()
            return None
            
        # Test key exists
        monkeypatch.setattr('winreg.OpenKey', mock_open_key_success)
        monkeypatch.setattr('winreg.CloseKey', mock_close_key)
        assert handler.key_exists("HKEY_CURRENT_USER", "Software\\Test") == True
        
        # Test key doesn't exist
        monkeypatch.setattr('winreg.OpenKey', mock_open_key_fail)
        assert handler.key_exists("HKEY_CURRENT_USER", "Software\\Test") == False
