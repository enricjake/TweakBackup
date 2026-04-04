"""
Tests for GUI components.
"""

import pytest  # Pytest framework for test markers and fixtures
import sys  # Used to modify the Python module search path
import os  # Provides environment variable checks and path operations

# Add project root to path so 'src' package can be imported
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import tkinter as tk  # Python's standard GUI toolkit — used to create Tk root windows


# Safe check for Tkinter availability
def _check_tk():
    # Returns True if a Tkinter display is available, False otherwise.
    # This prevents test failures in headless CI environments (e.g. GitHub Actions)
    # where no display server is present.

    # Always skip in GitHub Actions/Headless CI by default
    if os.getenv("GITHUB_ACTIONS") == "true":
        return False  # GITHUB_ACTIONS env var is set to 'true' in CI runners

    try:
        # Try to initialize a real Tk toolkit instance
        root = tk.Tk()  # Create a temporary Tk root window
        root.withdraw()  # Hide the window from the screen immediately
        # Verify basic title access works (confirms Tk is fully functional)
        _ = root.title()  # Read the window title — will raise if Tk is broken
        root.destroy()  # Tear down the temporary window and release resources
        return True  # Tkinter is fully operational
    except Exception as e:  # Catch EVERYTHING (including TclError and SystemExit)
        print(f"Tkinter check failed: {e}")
        return False  # Tkinter is not available or not functional


# Module-level flag: True only when a real Tk display is present.
# Used by @pytest.mark.skipif to conditionally skip all GUI tests.
TK_AVAILABLE = _check_tk()


@pytest.mark.skipif(
    not TK_AVAILABLE, reason="Tkinter/Display not available in this environment"
)
class TestMainWindow:
    """Test MainWindow class."""

    @pytest.fixture(autouse=True)
    def setup_gui(self):
        """Setup for GUI tests."""
        # Import the MainWindow class lazily (only when tests actually run)
        # to avoid import errors in environments without tkinter.
        from src.gui.main_window import MainWindow  # The main application window class

        self.MainWindowClass = (
            MainWindow  # Store the class reference for use in test methods
        )

    def test_window_creation(self):
        """Test that main window can be created."""
        try:
            root = tk.Tk()  # Create a real Tk root window for this test
        except tk.TclError as e:
            pytest.skip(f"Tkinter initialization failed: {e}")
        try:
            from src.version import VERSION  # Import the app version string

            app = self.MainWindowClass(
                root
            )  # Instantiate MainWindow with the root Tk window
            assert app is not None  # Verify the app object was created successfully
            # Verify the window title includes the version number and app tagline
            assert (
                root.title() == f"WinSet - v{VERSION} - Windows Configuration Toolkit"
            )
        finally:
            root.destroy()  # Always destroy the root window, even if an assertion fails

    def test_category_selection(self):
        """Test category selection functionality."""
        try:
            root = tk.Tk()  # Create a real Tk root window for this test
        except tk.TclError as e:
            pytest.skip(f"Tkinter initialization failed: {e}")
        try:
            app = self.MainWindowClass(root)  # Instantiate MainWindow

            # Test select all — enables every category checkbox
            app.select_all_categories()  # Method that sets all category BooleanVars to True
            for var in (
                app.category_vars.values()
            ):  # Iterate over all category BooleanVar objects
                assert var.get() is True  # Each checkbox variable should now be True

            # Test clear all — disables every category checkbox
            app.clear_all_categories()  # Method that sets all category BooleanVars to False
            for var in (
                app.category_vars.values()
            ):  # Re-iterate over all category BooleanVar objects
                assert var.get() is False  # Each checkbox variable should now be False

        finally:
            root.destroy()  # Always destroy the root window to free resources
