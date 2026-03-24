"""
Build script for WinSet — packages the app into a standalone Windows .exe
using PyInstaller.

Usage:
    python scripts/build_exe.py
"""

import subprocess
import sys
import shutil
import os
from pathlib import Path
import importlib.util


# The root of the repository, which is the parent of the 'scripts' directory.
PROJECT_ROOT = Path(__file__).resolve().parent.parent


def clean_dist():
    """Remove previous build artefacts."""
    for folder in ("dist", "build"):
        # Paths are resolved relative to the project root
        path = PROJECT_ROOT / folder
        if path.exists():
            try:
                shutil.rmtree(path)
                print(f"Cleaned {folder}/")
            except PermissionError:
                print(f"Error: Access denied while cleaning '{folder}'.")
                print("Please ensure 'WinSet.exe' is closed and try again.")
                sys.exit(1)


def build():
    """Run PyInstaller to produce a single-file executable."""
    if importlib.util.find_spec("PyInstaller") is None:
        print("Error: PyInstaller module not found.")
        print("Please run: pip install pyinstaller")
        sys.exit(1)

    main_script = PROJECT_ROOT / "winset.py"
    icon_path = PROJECT_ROOT / "docs" / "icon.ico"
    presets_data_path = f"{PROJECT_ROOT / 'presets'}{os.pathsep}presets"
    resources_data_path = f"{PROJECT_ROOT / 'src' / 'resources'}{os.pathsep}src/resources"

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",          # Single .exe output
        "--windowed",         # No console window (GUI app)
        "--name", "WinSet",
        "--icon", str(icon_path) if icon_path.exists() else "NONE",
        "--add-data", presets_data_path,  # Bundle the presets folder
        "--add-data", resources_data_path, # Bundle the settings resources
        str(main_script),
    ]

    print("Running PyInstaller…")
    # Run PyInstaller from the project root. This ensures the 'dist' and 'build'
    # folders are created in a predictable location.
    # Use Popen to have better control over the subprocess
    try:
        process = subprocess.Popen(cmd, cwd=PROJECT_ROOT)
        # Wait for the process with no timeout (PyInstaller can take a while)
        stdout, stderr = process.communicate()
    except KeyboardInterrupt:
        print("\nBuild interrupted. Terminating PyInstaller…")
        process.terminate()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()
        print("Build cancelled.")
        sys.exit(1)

    if process.returncode != 0:
        print("Build failed.")
        sys.exit(process.returncode)

    print(f"Build complete — output is in {PROJECT_ROOT / 'dist' / 'WinSet.exe'}")


if __name__ == "__main__":
    clean_dist()
    build()
