"""
Build script for WinSet — packages the app into a standalone Windows .exe
using PyInstaller.

Usage:
    python scripts/build_exe.py
"""

import subprocess
import sys
import shutil
from pathlib import Path


def clean_dist():
    """Remove previous build artefacts."""
    for folder in ("dist", "build"):
        path = Path(folder)
        if path.exists():
            shutil.rmtree(path)
            print(f"Cleaned {folder}/")


def build():
    """Run PyInstaller to produce a single-file executable."""
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",          # Single .exe output
        "--windowed",         # No console window (GUI app)
        "--name", "WinSet",
        "--icon", "docs/icon.ico" if Path("docs/icon.ico").exists() else "NONE",
        "winset.py",
    ]

    print("Running PyInstaller…")
    result = subprocess.run(cmd, check=False)

    if result.returncode != 0:
        print("Build failed.")
        sys.exit(result.returncode)

    print("Build complete — output is in dist/WinSet.exe")


if __name__ == "__main__":
    clean_dist()
    build()
