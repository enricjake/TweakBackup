"""
Tests for GUI components.
"""

import pytest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from src.gui.main_window import MainWindow
import tkinter as tk


try:
    _root = tk.Tk()
    _root.destroy()
    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False


@pytest.mark.skipif(not TK_AVAILABLE, reason="Tkinter/Display not available in this environment")
class TestMainWindow:
    """Test MainWindow class."""
    
    def test_window_creation(self):
        """Test that main window can be created."""
        root = tk.Tk()
        try:
            app = MainWindow(root)
            assert app is not None
            assert root.title() == "WinSet - Windows Configuration Toolkit"
        finally:
            root.destroy()
    
    def test_category_selection(self):
        """Test category selection functionality."""
        root = tk.Tk()
        try:
            app = MainWindow(root)
            
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