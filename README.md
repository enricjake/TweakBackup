# 🔧 TweakBackup - Windows Settings Exporter

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![Downloads](https://img.shields.io/github/downloads/yourusername/tweakbackup/total)](https://github.com/enricjake/tweakbackup/releases)

**Never manually reconfigure Windows after reinstall again!**

TweakBackup exports ALL your Windows settings to a single file. After reinstalling Windows, import that file and all your preferences are back in seconds.

## ✨ Features

- 🎨 **System Appearance** - Theme, wallpaper, accent colors, icons
- 📁 **File Explorer** - Hidden files, extensions, folder views
- 🖱️ **Taskbar & Start Menu** - Pinned apps, layout, search settings
- ⚡ **Power Settings** - Plans, sleep timers, lid actions
- 🔒 **Privacy Options** - Permissions, telemetry, location
- 💾 **One-Click Export/Import** - Single JSON file
- 🔄 **Profile Manager** - Save multiple configurations
- 🛡️ **Safe Mode** - Registry backup before changes

## 🚀 Quick Start

1. Download the latest `TweakBackup.exe` from [Releases](https://github.com/enricjake/tweakbackup/releases)
2. Run as Administrator (required for registry access)
3. Click "Export" to save your settings
4. On new PC, run and click "Import"

## 📸 Screenshots

*[Add screenshots here]*

## 🛠️ Building from Source

```bash
git clone https://github.com/yourusername/tweakbackup.git
cd tweakbackup
pip install -r requirements.txt
python tweakbackup.py
