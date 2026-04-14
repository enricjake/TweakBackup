#!/usr/bin/env python3
"""
WinSet - Windows Configuration Toolkit
Main entry point for the application
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox
import ctypes
import logging
from logging.handlers import RotatingFileHandler

# Root directory for application data in the user's LOCALAPPDATA folder.
# Used to store logs and other persistent data for the WinSet application.
_app_data_root = os.getenv("LOCALAPPDATA", os.path.expanduser("~"))

# Path to the WinSet log directory inside LOCALAPPDATA.
# All application logs are stored here.
_log_dir = os.path.join(_app_data_root, "WinSet")

# Create the log directory if it does not already exist.
os.makedirs(_log_dir, exist_ok=True)

# Configure the Python logging system with two handlers:
#   1. RotatingFileHandler: writes logs to a file, rotating after 1 MB with up to 3 backups.
#   2. StreamHandler: also prints log messages to the console (stdout).
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        RotatingFileHandler(
            filename=os.path.join(_log_dir, "winset.log"),
            maxBytes=1024 * 1024,  # 1 MB max file size before rotation
            backupCount=3,  # Keep up to 3 rotated backup log files
        ),
        logging.StreamHandler(),  # Also log to console for developer visibility
    ],
)

# Application logger for this module, used throughout winset.py for diagnostic messages.
logger = logging.getLogger(__name__)

# Add the 'src' directory to Python's module search path so that
# application modules (e.g. src.gui.main_window) can be imported directly.
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

def check_admin():
    """Check if the current process is running with administrator privileges on Windows.

    Returns:
        bool: True if the process has admin rights, False otherwise.
    """
    try:
        # Windows-specific API call via ctypes to determine admin status.
        return ctypes.windll.shell32.IsUserAnAdmin()
    except (OSError, AttributeError):
        # Fallback: if the API is unavailable (e.g. non-Windows), assume not admin.
        return False


def is_frozen():
    """Check if the application is running as a frozen PyInstaller executable.

    Returns:
        bool: True if running as .exe, False if running as script.
    """
    return getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")


def get_executable_path():
    """Get the path to the current executable.

    For frozen PyInstaller apps, returns the bundled .exe path.
    For script execution, returns sys.executable.

    Returns:
        str: Path to the executable.
    """
    if is_frozen():
        # PyInstaller temp extraction folder - use sys.executable for the actual .exe
        return sys.executable
    return sys.executable

def main():
    """Main entry point for the WinSet application.

    Performs the following steps:
    1. Creates the root Tkinter window (hidden initially).
    2. Checks for administrator privileges; prompts user to elevate if needed.
    3. Imports and initializes the MainWindow GUI.
    4. Starts the Tkinter event loop.
    """
    logger.info("Starting WinSet application")

    # Create the root Tkinter window but keep it hidden until the GUI is fully initialized.
    root = tk.Tk()
    root.withdraw()

    # Verify that the application has administrator privileges, which are
    # required for modifying Windows Registry and system settings.
    if not check_admin():
        logger.warning("Application not running as administrator")
        result = messagebox.askyesno(
            "Administrator Rights Required",
            "WinSet needs administrator privileges to modify system settings.\n\n"
            "Do you want to restart as administrator?",
        )

        if result:
            # Relaunch the application with elevated (administrator) privileges
            # using the Windows ShellExecute 'runas' verb.
            logger.info("Relaunching as administrator")
            exe_path = get_executable_path()
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", exe_path, " ".join(sys.argv), None, 1
            )

        # Close the hidden root window and exit since we're either relaunching or the user declined.
        root.destroy()
        sys.exit(0)

    logger.info("Running with administrator privileges")

    # Import MainWindow here (lazy import) to avoid circular import issues,
    # since src.gui.main_window imports from other src modules.
    from src.gui.main_window import MainWindow

    # Initialize the main application window, passing the existing root Tk instance.
    app = MainWindow(root)

    # Reveal the window now that the GUI is fully set up.
    root.deiconify()

    # Start the Tkinter event loop — blocks until the window is closed.
    root.mainloop()

    logger.info("Application closed")


if __name__ == "__main__":
    # Entry point when running this script directly (not imported as a module).
    main()

