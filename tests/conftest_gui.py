"""
Configuration for GUI tests - set up environment for Tkinter
"""

import os
import sys

# Set up Tcl/Tk library paths before any Tkinter imports
if sys.platform == 'win32':
    python_dir = os.path.dirname(sys.executable)
    tcl_dir = os.path.join(python_dir, 'tcl')
    if os.path.exists(tcl_dir):
        os.environ['TCL_LIBRARY'] = os.path.join(tcl_dir, 'tcl8.6')
        os.environ['TK_LIBRARY'] = os.path.join(tcl_dir, 'tk8.6')
        print(f"Set TCL_LIBRARY to: {os.environ['TCL_LIBRARY']}")
        print(f"Set TK_LIBRARY to: {os.environ['TK_LIBRARY']}")