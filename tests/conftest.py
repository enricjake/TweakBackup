"""
Pytest configuration and shared fixtures
"""

import pytest  # Pytest framework for fixtures and test configuration
import sys  # Used to modify the Python module search path
import os  # Provides OS path operations for locating source directories
import tempfile  # Creates temporary directories/files for isolated test environments
from pathlib import Path  # Object-oriented filesystem path handling

# Set up Tcl/Tk library paths before any Tkinter imports
if sys.platform == 'win32':
    python_dir = os.path.dirname(sys.executable)
    tcl_dir = os.path.join(python_dir, 'tcl')
    if os.path.exists(tcl_dir):
        os.environ['TCL_LIBRARY'] = os.path.join(tcl_dir, 'tcl8.6')
        os.environ['TK_LIBRARY'] = os.path.join(tcl_dir, 'tk8.6')

# Add src to path so test modules can import application code from the src package
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))


@pytest.fixture
def temp_dir():
    """Create a temporary directory for test files"""
    # tempfile.TemporaryDirectory creates a short-lived directory that is
    # automatically deleted when the context manager exits.
    # 'yield' hands the directory path to the test; cleanup runs after the test finishes.
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir  # tmpdir: str path to the temporary directory


@pytest.fixture
def sample_profile_data():
    """Sample profile data for testing"""
    # Returns a dict that mimics the structure of a WinSet profile,
    # including a name, description, version, and a collection of
    # named settings each with a value and description.
    return {
        "name": "Test Profile",  # Display name of the profile
        "description": "A test profile",  # Human-readable description
        "version": "1.0",  # Schema version of the profile format
        "settings": {
            "test_setting_1": {
                "value": 1,  # Numeric value for the first test setting
                "description": "Test setting 1",  # Description of the first setting
            },
            "test_setting_2": {
                "value": "test",  # String value for the second test setting
                "description": "Test setting 2",  # Description of the second setting
            },
        },
    }


@pytest.fixture
def sample_registry_setting():
    """Sample registry setting for testing"""
    # Returns a dict representing a single Windows registry setting,
    # including the hive, key path, value name, type, and metadata.
    return {
        "hive": "HKEY_CURRENT_USER",  # Registry hive root (HKCU)
        "key_path": "Software\\WinSet\\Test",  # Full subkey path within the hive
        "value_name": "TestValue",  # Name of the registry value entry
        "value_type": "REG_DWORD",  # Registry data type (32-bit integer)
        "value": 1,  # The stored value
        "name": "Test Setting",  # Display name for the setting
        "description": "A test registry setting",  # Human-readable description
    }


@pytest.fixture
def mock_registry(monkeypatch):
    """Mock registry access for testing without touching the real Windows registry"""

    # MockRegistry is an in-memory stand-in for the winreg module.
    # It stores key/value pairs in dicts so tests can exercise registry
    # read/write logic without requiring actual Windows registry access.
    class MockRegistry:
        def __init__(self):
            self.values = {}  # Flat dict: value_name -> value (tracks all written values)
            self.keys = {}  # Dict: key_id -> {value_name: (value, type)} (per-key store)

        def OpenKey(self, hive, key_path, reserved, access):
            # Simulates winreg.OpenKey by returning a composite key identifier string.
            # hive: str registry hive name (e.g. "HKEY_CURRENT_USER")
            # key_path: str subkey path (e.g. "Software\\WinSet\\Test")
            # reserved: int reserved parameter (unused in mock)
            # access: int access rights flags (unused in mock)
            # Returns: str key identifier in the format "HIVE\\key_path"
            key_id = f"{hive}\\{key_path}"
            if key_id not in self.keys:
                self.keys[key_id] = {}  # Auto-create key entry if it doesn't exist
            return key_id  # Return the composite key identifier

        def CreateKey(self, hive, key_path):
            # Simulates winreg.CreateKey by creating (or overwriting) a key entry.
            # hive: str registry hive name
            # key_path: str subkey path to create
            # Returns: str key identifier for the newly created key
            key_id = f"{hive}\\{key_path}"
            self.keys[key_id] = {}  # Initialize empty value store for the key
            return key_id

        def SetValueEx(self, key, value_name, reserved, value_type, value):
            # Simulates winreg.SetValueEx by storing a value under a key.
            # key: str key identifier (returned by OpenKey/CreateKey)
            # value_name: str name of the value entry to set
            # reserved: int reserved parameter (unused)
            # value_type: str/int registry type constant (e.g. "REG_DWORD")
            # value: the actual data to store
            # Returns: bool True on success
            if key not in self.keys:
                self.keys[key] = {}  # Ensure the key dict exists
            self.keys[key][value_name] = (value, value_type)  # Store value with type
            self.values[value_name] = value  # Also store in the flat lookup dict
            return True

        def QueryValueEx(self, key, value_name):
            # Simulates winreg.QueryValueEx by looking up a stored value.
            # key: str key identifier
            # value_name: str name of the value entry to read
            # Returns: tuple (value, type) if found, or (None, None) if missing
            if key in self.keys and value_name in self.keys[key]:
                return self.keys[key][value_name]  # Return (value, type) tuple
            return (None, None)  # Value not found — return null tuple

        def CloseKey(self, key):
            # Simulates winreg.CloseKey — no-op since nothing needs closing.
            # key: str key identifier (ignored)
            pass

    mock_reg = MockRegistry()  # Instantiate the mock registry object

    # Replace real winreg module functions with mock implementations.
    # monkeypatch.setattr patches the target attribute for the duration of the test.
    monkeypatch.setattr("winreg.OpenKey", mock_reg.OpenKey)  # Patch key opening
    monkeypatch.setattr("winreg.CreateKey", mock_reg.CreateKey)  # Patch key creation
    monkeypatch.setattr("winreg.SetValueEx", mock_reg.SetValueEx)  # Patch value writing
    monkeypatch.setattr(
        "winreg.QueryValueEx", mock_reg.QueryValueEx
    )  # Patch value reading
    monkeypatch.setattr("winreg.CloseKey", mock_reg.CloseKey)  # Patch key closing

    # Replace real winreg hive constants with simple string values for consistency.
    monkeypatch.setattr("winreg.HKEY_CURRENT_USER", "HKEY_CURRENT_USER")
    monkeypatch.setattr("winreg.HKEY_LOCAL_MACHINE", "HKEY_LOCAL_MACHINE")
    monkeypatch.setattr("winreg.HKEY_CLASSES_ROOT", "HKEY_CLASSES_ROOT")
    monkeypatch.setattr("winreg.HKEY_CURRENT_CONFIG", "HKEY_CURRENT_CONFIG")
    monkeypatch.setattr("winreg.HKEY_USERS", "HKEY_USERS")

    # Replace real winreg type constants with simple string values.
    monkeypatch.setattr("winreg.REG_SZ", "REG_SZ")  # String type
    monkeypatch.setattr("winreg.REG_DWORD", "REG_DWORD")  # 32-bit integer type
    monkeypatch.setattr("winreg.REG_BINARY", "REG_BINARY")  # Binary data type
    monkeypatch.setattr("winreg.REG_MULTI_SZ", "REG_MULTI_SZ")  # Multi-string type
    monkeypatch.setattr(
        "winreg.REG_EXPAND_SZ", "REG_EXPAND_SZ"
    )  # Expandable string type
    monkeypatch.setattr("winreg.REG_QWORD", "REG_QWORD")  # 64-bit integer type

    return mock_reg  # Return the mock instance so tests can inspect its state
