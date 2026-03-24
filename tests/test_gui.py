"""
Tests for GUI components.
"""

import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import tkinter as tk

# Safe check for Tkinter availability
def _check_tk():
    try:
        # Try to initialize a real toolkit instance
        root = tk.Tk()
        root.withdraw()
        # Verify basic title access works
        _ = root.title()
        root.destroy()
        return True
    except: # Catch EVERYTHING (including TclError and SystemExit)
        return False

TK_AVAILABLE = _check_tk()


@pytest.mark.skipif(not TK_AVAILABLE, reason="Tkinter/Display not available in this environment")
class TestMainWindow:
    """Test MainWindow class."""
    
    @pytest.fixture(autouse=True)
    def setup_gui(self):
        """Setup for GUI tests."""
        from src.gui.main_window import MainWindow
        self.MainWindowClass = MainWindow
    
    def test_window_creation(self):
        """Test that main window can be created."""
        root = tk.Tk()
        try:
            from src.version import VERSION
            app = self.MainWindowClass(root)
            assert app is not None
            assert root.title() == f"WinSet - v{VERSION} - Windows Configuration Toolkit"
        finally:
            root.destroy()
    
    def test_category_selection(self):
        """Test category selection functionality."""
        root = tk.Tk()
        try:
            app = self.MainWindowClass(root)
            
            # Test select all
            app.select_all_categories()
            for var in app.category_vars.values():
                assert var.get() is True
            
            # Test clear all
            app.clear_all_categories()
            for var in app.category_vars.values():
                assert var.get() is False
                
        finally:
            root.destroy()