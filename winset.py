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

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        RotatingFileHandler(
            filename=os.path.join(os.getenv('LOCALAPPDATA', '~'), 'WinSet', 'winset.log'),
            maxBytes=1024*1024,  # 1MB
            backupCount=3
        ),
        logging.StreamHandler()  # Also log to console
    ]
)
        
logger = logging.getLogger(__name__)

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def check_admin():
    """Check if running as administrator"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except (OSError, AttributeError):
        return False

def main():
    """Main entry point"""
    logger.info("Starting WinSet application")

    # Create the root window but hide it initially
    root = tk.Tk()
    root.withdraw()

    # Check if running as admin
    if not check_admin():
        logger.warning("Application not running as administrator")
        result = messagebox.askyesno(
            "Administrator Rights Required",
            "WinSet needs administrator privileges to modify system settings.\n\n"
            "Do you want to restart as administrator?"
        )
        
        if result:
            # Relaunch as admin
            logger.info("Relaunching as administrator")
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )

        root.destroy()
        sys.exit(0)
    
    logger.info("Running with administrator privileges")

    # Import main window here to avoid circular imports
    from src.gui.main_window import MainWindow
    
    # Root is already created, so just initialize MainWindow
    app = MainWindow(root)
    root.deiconify() # Show the window now that it's ready
    root.mainloop()

    logger.info("Application closed")

if __name__ == "__main__":
    main()