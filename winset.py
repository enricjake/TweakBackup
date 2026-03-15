#!/usr/bin/env python3
"""
WinSet - Windows Configuration Toolkit
Main entry point for the application
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

def check_admin():
    """Check if running as administrator"""
    import ctypes
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except (OSError, AttributeError):
        return False

def main():
    """Main entry point"""
    # Check if running as admin
    if not check_admin():
        root = tk.Tk()
        root.withdraw()
        result = messagebox.askyesno(
            "Administrator Rights Required",
            "WinSet needs administrator privileges to modify system settings.\n\n"
            "Do you want to restart as administrator?"
        )
        root.destroy()
        
        if result:
            # Relaunch as admin
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
        sys.exit(0)
    
    # Import main window here to avoid circular imports
    from src.gui.main_window import MainWindow
    
    # Create and run main window
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()