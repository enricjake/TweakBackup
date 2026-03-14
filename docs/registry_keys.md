# Windows Registry Keys for WinSet

## 📋 Overview
This document contains all Windows Registry keys that WinSet can export and import. Each setting includes its path, value name, data type, and possible values.

**⚠️ WARNING**: Modifying the Windows Registry can cause system instability. Always create a system restore point before making changes.

---

## 🎨 Category 1: System Appearance

### Desktop Wallpaper
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `Wallpaper` |
| **Type** | `REG_SZ` |
| **Description** | Path to current wallpaper image |
| **Example** | `C:\Users\Username\Pictures\wallpaper.jpg` |

### Wallpaper Style (Stretch/Fill/Fit)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `WallpaperStyle` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Centered, `2` = Stretched, `6` = Fit, `10` = Fill, `22` = Span |
| **Description** | How wallpaper fits the screen |

### Tile Wallpaper
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `TileWallpaper` |
| **Type** | `REG_SZ` |
| **Values** | `0` = No tile, `1` = Tile |
| **Description** | Whether wallpaper is tiled |

### App Theme (Light/Dark)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Themes\Personalize` |
| **Value** | `AppsUseLightTheme` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Light mode, `0` = Dark mode |
| **Description** | Theme for applications |

### System Theme (Light/Dark)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Themes\Personalize` |
| **Value** | `SystemUsesLightTheme` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Light mode, `0` = Dark mode |
| **Description** | Theme for system elements (taskbar, start menu) |

### Accent Color
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\DWM` |
| **Value** | `AccentColor` |
| **Type** | `REG_DWORD` |
| **Format** | `0x00BBGGRR` (hex: 00 Blue Green Red) |
| **Example** | `0x00FF8000` = Blue accent |
| **Description** | Windows accent color in BGR format |

### Accent Color (Auto-Selected)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\DWM` |
| **Value** | `ColorizationColor` |
| **Type** | `REG_DWORD` |
| **Description** | Auto-selected accent color from wallpaper |

### Accent Color on Title Bars
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\DWM` |
| **Value** | `ColorPrevalence` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show accent on title bars, `0` = Don't show |
| **Description** | Show accent color on window title bars |

### Show Accent on Start/Taskbar
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Themes\Personalize` |
| **Value** | `ColorPrevalence` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Don't show |
| **Description** | Show accent color on Start menu and taskbar |

### Transparency Effects
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Themes\Personalize` |
| **Value** | `EnableTransparency` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Enable transparency effects in Windows |

### Taskbar Transparency
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `TaskbarTransparency` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Make taskbar transparent |

### Desktop Icon Visibility - This PC
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel` |
| **Value** | `{20D04FE0-3AEA-1069-A2D8-08002B30309D}` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Show, `1` = Hide |
| **Description** | Show/hide "This PC" icon on desktop |

### Desktop Icon Visibility - Network
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel` |
| **Value** | `{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Show, `1` = Hide |
| **Description** | Show/hide "Network" icon on desktop |

### Desktop Icon Visibility - Recycle Bin
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel` |
| **Value** | `{645FF040-5081-101B-9F08-00AA002F954E}` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Show, `1` = Hide |
| **Description** | Show/hide Recycle Bin on desktop |

### Desktop Icon Visibility - User's Files
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel` |
| **Value** | `{59031a47-3f72-44a7-89c5-5595fe6b30ee}` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Show, `1` = Hide |
| **Description** | Show/hide user's files folder on desktop |

### Desktop Icon Size
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\Shell\Bags\1\Desktop` |
| **Value** | `IconSize` |
| **Type** | `REG_DWORD` |
| **Values** | `32` = Small, `48` = Medium, `96` = Large |
| **Description** | Size of desktop icons in pixels |

### Desktop Icon Spacing (Horizontal)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop\WindowMetrics` |
| **Value** | `IconSpacing` |
| **Type** | `REG_SZ` |
| **Format** | `-1125` (negative number = pixels) |
| **Description** | Horizontal spacing between desktop icons |

### Desktop Icon Spacing (Vertical)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop\WindowMetrics` |
| **Value** | `IconVerticalSpacing` |
| **Type** | `REG_SZ` |
| **Format** | `-1125` (negative number = pixels) |
| **Description** | Vertical spacing between desktop icons |

### Font Smoothing (ClearType)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `FontSmoothing` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Disabled, `2` = Enabled |
| **Description** | Enable/disable font smoothing |

### Font Smoothing Type
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `FontSmoothingType` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Standard, `2` = ClearType |
| **Description** | Type of font smoothing |

### Visual Effects - Animations
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `UserPreferencesMask` |
| **Type** | `REG_BINARY` |
| **Description** | Complex binary value controlling visual effects |

### Show Shadows Under Windows
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `ListviewShadow` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Show shadows under windows |

### Show Shadows Under Mouse
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `ShowInfoTip` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Show pop-up descriptions for folder items |

---

## 📁 Category 2: File Explorer Settings

### Hidden Files
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `Hidden` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show hidden files, `2` = Don't show |
| **Description** | Show/hide hidden files and folders |

### File Extensions
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `HideFileExt` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Show extensions, `1` = Hide extensions |
| **Description** | Show/hide file name extensions |

### Protected Operating System Files
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `ShowSuperHidden` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show/hide protected OS files |

### Empty Drives
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `HideDrivesWithNoMedia` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Show, `1` = Hide |
| **Description** | Hide empty drives in Computer folder |

### Navigation Pane - Show All Folders
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `NavPaneShowAllFolders` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show all folders in navigation pane |

### Navigation Pane - Expand to Open Folder
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `NavPaneExpandToCurrentFolder` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Automatically expand to current folder |

### Quick Access - Show Recent Files
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer` |
| **Value** | `ShowRecent` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show recently used files in Quick Access |

### Quick Access - Show Frequent Folders
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer` |
| **Value** | `ShowFrequent` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show frequently used folders in Quick Access |

### File Explorer - Open to Quick Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `LaunchTo` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Quick Access, `2` = This PC |
| **Description** | Default location when opening File Explorer |

### Check Boxes to Select Items
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `AutoCheckSelect` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Show check boxes for item selection |

### Item Check Boxes
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `UseCheckBoxes` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Enable check boxes for item selection |

### File Explorer - View Type
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Streams` |
| **Value** | `Settings` |
| **Type** | `REG_BINARY` |
| **Description** | Binary value containing view settings (Details, Icons, List, etc.) |

### Folder Merge Conflicts
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `MergeFolderConflicts` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Don't merge, `1` = Merge |
| **Description** | How to handle folder merge conflicts |

### Sharing Wizard
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `SharingWizardOn` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Use, `0` = Don't use |
| **Description** | Use Sharing Wizard for network sharing |

### Recycle Bin - Delete Confirmation
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `ConfirmFileDelete` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = No confirmation, `1` = Confirm |
| **Description** | Show delete confirmation dialog |

### Recycle Bin - Maximum Size
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\BitBucket\Volume\{drive_guid}` |
| **Value** | `MaxCapacity` |
| **Type** | `REG_DWORD` |
| **Description** | Maximum size of Recycle Bin for each drive (percentage) |

### Recycle Bin - Files Removed Immediately
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\BitBucket\Volume\{drive_guid}` |
| **Value** | `NukeOnDelete` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Send to Recycle Bin, `1` = Remove immediately |
| **Description** | Don't move files to Recycle Bin, delete immediately |

### Folder Tips
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `ShowInfoTip` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show pop-up descriptions for folder items |

### Sync Provider Notifications
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `ShowSyncProviderNotifications` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show sync provider notifications (OneDrive, etc.) |

### File Explorer - Display Size Info
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `FolderContentsInfoTip` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Display file size information in folder tips |

### Encrypted/Compressed Files in Color
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `ShowEncryptCompressedColor` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show encrypted/compressed files in alternate color |

### Always Show Icons, Never Thumbnails
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `IconsOnly` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Always show icons, `0` = Show thumbnails |
| **Description** | Show icons instead of file thumbnails |

### Display File Icon on Thumbnails
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `ShowTypeOverlay` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Display file type icon on thumbnails |

### Separate Process for Folder Windows
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `SeparateProcess` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Launch folder windows in separate processes |

---

## 🖱️ Category 3: Taskbar & Start Menu

### Taskbar Alignment
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `TaskbarAl` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Left, `1` = Center |
| **Description** | Taskbar icon alignment |

### Taskbar Size (Small/Large Icons)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `TaskbarSmallIcons` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Small icons, `0` = Large icons |
| **Description** | Use small taskbar icons |

### Taskbar Location on Screen
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3` |
| **Value** | `Settings` |
| **Type** | `REG_BINARY` |
| **Description** | Binary value containing taskbar position (Bottom, Top, Left, Right) |

### Auto-Hide Taskbar
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3` |
| **Value** | `Settings` |
| **Type** | `REG_BINARY` |
| **Description** | Binary value containing auto-hide setting |

### Lock Taskbar
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `TaskbarSizeMove` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Locked, `1` = Unlocked |
| **Description** | Lock/unlock the taskbar |

### Taskbar Badges
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `TaskbarBadges` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Show taskbar badges for apps |

### Taskbar Corner Overflow
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer` |
| **Value** | `EnableAutoTray` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Show all, `1` = Auto-hide |
| **Description** | Control system tray icon behavior |

### Taskbar Widgets
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `TaskbarDa` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show widgets button on taskbar |

### Taskbar Chat Icon
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `TaskbarMn` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show chat icon (Microsoft Teams) on taskbar |

### Taskbar Task View Button
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `ShowTaskViewButton` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show Task View button on taskbar |

### Taskbar People Button
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `PeopleBand` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show People button on taskbar |

### Taskbar Search Box
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Search` |
| **Value** | `SearchboxTaskbarMode` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Hidden, `1` = Search icon, `2` = Search box |
| **Description** | Control search box appearance on taskbar |

### Taskbar Corner Icons - Pen Menu
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\PenWorkspace` |
| **Value** | `PenWorkspaceButtonDesiredVisibility` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show pen menu icon in system tray |

### Taskbar Corner Icons - Touch Keyboard
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CPSS\TouchKeyboard` |
| **Value** | `Cornervisibility` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show touch keyboard icon in system tray |

### Taskbar Corner Icons - Virtual Touchpad
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CPSS\TouchPad` |
| **Value** | `Cornervisibility` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show virtual touchpad icon in system tray |

### Combine Taskbar Buttons
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `TaskbarGlomLevel` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Always combine, `1` = Combine when full, `2` = Never combine |
| **Description** | Control how taskbar buttons are grouped |

### Multiple Displays - Show Taskbar on All Displays
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `MMTaskbarEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Main display only, `1` = All displays |
| **Description** | Show taskbar on all displays |

### Multiple Displays - Where to Show Buttons
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `MMTaskbarMode` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = All taskbars, `1` = Main taskbar, `2` = Taskbar where window is open |
| **Description** | Control where taskbar buttons appear on multiple displays |

### Multiple Displays - Combine Buttons on Other Taskbars
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `MMTaskbarGlomLevel` |
| **Type** | `REG_DWORD` |
| **Values** | Same as TaskbarGlomLevel |
| **Description** | Button combining behavior on secondary taskbars |

### Start Menu - Show Recently Added Apps
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `Start_NotifyNewApps` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Highlight newly installed apps in Start menu |

### Start Menu - Show Most Used Apps
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `Start_TrackProgs` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show most used apps in Start menu |

### Start Menu - Show Suggestions
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager` |
| **Value** | `SystemPaneSuggestionsEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Show app suggestions in Start menu |

### Start Menu - Show Recently Opened Items
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `Start_TrackDocs` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Show, `0` = Hide |
| **Description** | Show recently opened items in Start menu |

### Start Menu - Size
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `Start_LargeMetroView` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Small, `1` = Large |
| **Description** | Control Start menu size |

### Start Menu Layout XML
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount\$$start.layout\Current` |
| **Value** | `Data` |
| **Type** | `REG_BINARY` |
| **Description** | Binary data containing Start menu layout configuration |

### Pinned Taskbar Apps
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband` |
| **Value** | `Favorites` |
| **Type** | `REG_BINARY` |
| **Description** | Binary list of pinned taskbar applications |

### Pinned Start Menu Apps
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage2` |
| **Value** | `Favorites` |
| **Type** | `REG_BINARY` |
| **Description** | Binary list of pinned Start menu items |

---

## ⚡ Category 4: Power Settings

### Active Power Plan
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes` |
| **Value** | `ActivePowerScheme` |
| **Type** | `REG_SZ` |
| **Format** | GUID (e.g., `381b4222-f694-41f0-9685-ff5bb260df2e`) |
| **Description** | GUID of currently active power plan |

### Power Plan GUIDs Reference
| Plan | GUID |
|------|------|
| **Balanced** | `381b4222-f694-41f0-9685-ff5bb260df2e` |
| **Power Saver** | `a1841308-3541-4fab-bc81-f71556f20b4a` |
| **High Performance** | `8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c` |
| **Ultimate Performance** | `e9a42b02-d5df-448d-aa00-03f14749eb61` |

### Monitor Timeout (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\PowerCfg\PowerPolicies` |
| **Value** | `VideoTimeoutAC` |
| **Type** | `REG_DWORD` |
| **Unit** | Minutes |
| **Description** | Minutes before monitor turns off when plugged in |

### Monitor Timeout (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\PowerCfg\PowerPolicies` |
| **Value** | `VideoTimeoutDC` |
| **Type** | `REG_DWORD` |
| **Unit** | Minutes |
| **Description** | Minutes before monitor turns off on battery |

### Sleep Timeout (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\PowerCfg\PowerPolicies` |
| **Value** | `SleepTimeoutAC` |
| **Type** | `REG_DWORD` |
| **Unit** | Minutes |
| **Description** | Minutes before PC sleeps when plugged in |

### Sleep Timeout (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\PowerCfg\PowerPolicies` |
| **Value** | `SleepTimeoutDC` |
| **Type** | `REG_DWORD` |
| **Unit** | Minutes |
| **Description** | Minutes before PC sleeps on battery |

### Hibernate Timeout (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\PowerSettings\{238C9FA8-0AAD-41ED-83F4-97BE242C8F20}\{9d7815a6-7ee4-497e-8888-515a05f02364}` |
| **Value** | `Attributes` (for settings), `DefaultPowerSchemeValues` for values |
| **Description** | Complex - requires reading subkeys per power scheme |

### Hibernate Timeout (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Description** | Similar to above, under DC subkey |

### Lid Close Action (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{4f971e89-eebd-4455-a8de-9e59040e7347}\{5ca83367-6e45-459f-a27b-476b1d01c936}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Do nothing, `1` = Sleep, `2` = Hibernate, `3` = Shut down |
| **Description** | Action when laptop lid is closed (plugged in) |

### Lid Close Action (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{4f971e89-eebd-4455-a8de-9e59040e7347}\{5ca83367-6e45-459f-a27b-476b1d01c936}` |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | Same as above |
| **Description** | Action when laptop lid is closed (on battery) |

### Power Button Action (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{4f971e89-eebd-4455-a8de-9e59040e7347}\{7648efa3-dd9c-4e3e-b566-50f929386280}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | Same as lid close action |
| **Description** | Action when power button pressed (plugged in) |

### Power Button Action (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Action when power button pressed (on battery) |

### Sleep Button Action (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{4f971e89-eebd-4455-a8de-9e59040e7347}\{96996bc0-ad50-47ec-923b-6f41874dd9eb}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Action when sleep button pressed (plugged in) |

### Sleep Button Action (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Action when sleep button pressed (on battery) |

### Hard Disk Turn Off After (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{0012ee47-9041-4b5d-9b77-535fba8b1442}\{6738e2c4-e8a5-4a42-b16a-e040e769756e}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Unit** | Minutes (0 = Never) |
| **Description** | Turn off hard disk after (plugged in) |

### Hard Disk Turn Off After (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Turn off hard disk after (on battery) |

### USB Selective Suspend (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{2a737441-1930-4402-8d77-b2bebba308a3}\{48e6b7a6-50f5-4782-a5d4-53bb8f07e226}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Enable USB selective suspend (plugged in) |

### USB Selective Suspend (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Enable USB selective suspend (on battery) |

### PCI Express Link State Power Management (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{501a4d13-42af-4429-9fd1-a8218c268e20}\{ee12f906-d277-404b-b6da-e5fa1a576df5}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Off, `1` = Moderate, `2` = Maximum |
| **Description** | PCI Express power management (plugged in) |

### PCI Express Link State Power Management (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | PCI Express power management (on battery) |

### Processor Power Management - Minimum State (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{54533251-82be-4824-96c1-47b60b740d00}\{893dee8e-2bef-41e0-89c6-b55d0929964c}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Range** | `0` - `100` (percentage) |
| **Description** | Minimum processor state (plugged in) |

### Processor Power Management - Minimum State (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Minimum processor state (on battery) |

### Processor Power Management - Maximum State (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{54533251-82be-4824-96c1-47b60b740d00}\{bc5038f7-23e0-4960-96da-33abaf5935ec}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Range** | `0` - `100` (percentage) |
| **Description** | Maximum processor state (plugged in) |

### Processor Power Management - Maximum State (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Maximum processor state (on battery) |

### Processor Power Management - System Cooling Policy (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{54533251-82be-4824-96c1-47b60b740d00}\{94d3a615-a899-4ac5-ae2b-e4d8f634367f}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Passive, `1` = Active |
| **Description** | Cooling policy (plugged in) |

### Processor Power Management - System Cooling Policy (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Cooling policy (on battery) |

### Display Brightness (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{7516b95f-f776-4464-8c53-06167f40cc99}\{aded5e82-b909-4619-9949-f5d71dac0bcb}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Range** | `0` - `100` (percentage) |
| **Description** | Display brightness (plugged in) |

### Display Brightness (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Display brightness (on battery) |

### Adaptive Brightness (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{7516b95f-f776-4464-8c53-06167f40cc99}\{fbd9aa66-9553-4097-ba44-ed6e9d65eab8}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Enable adaptive brightness (plugged in) |

### Adaptive Brightness (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Enable adaptive brightness (on battery) |

### Wireless Adapter Power Saving Mode (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{19cbb8fa-5279-450e-9fac-8a3d5fedd0c1}\{12bbebe6-58d6-4636-95bb-3217ef867c1a}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Maximum Performance, `1` = Medium, `2` = Maximum Power Saving |
| **Description** | Wireless adapter power saving (plugged in) |

### Wireless Adapter Power Saving Mode (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Wireless adapter power saving (on battery) |

### Battery Level to Enter Critical Action
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{e73a048d-bf27-4f12-9731-8b2076e8891f}\{9a66d8d7-4ff7-4ef9-b5a2-5a326ca2a469}` |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Range** | `0` - `100` (percentage) |
| **Description** | Battery level that triggers critical action |

### Critical Battery Action
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{e73a048d-bf27-4f12-9731-8b2076e8891f}\{637ea02f-bbcb-4015-8e2c-a1c7b9c0b546}` |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | Same as lid close action |
| **Description** | Action when battery reaches critical level |

### Low Battery Level
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{e73a048d-bf27-4f12-9731-8b2076e8891f}\{8183ba9a-e910-48da-8769-14ae6dc1170a}` |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Range** | `0` - `100` (percentage) |
| **Description** | Battery level considered "low" |

### Low Battery Notification
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{e73a048d-bf27-4f12-9731-8b2076e8891f}\{bcded951-187b-4d05-bccc-f7e51960c258}` |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Off, `1` = On |
| **Description** | Show notification when battery is low |

### Low Battery Action
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{e73a048d-bf27-4f12-9731-8b2076e8891f}\{d8742dcb-3e6a-4b3c-b3fe-374623cdcf06}` |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | Same as lid close action |
| **Description** | Action when battery reaches low level |

### Reserve Battery Level
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{e73a048d-bf27-4f12-9731-8b2076e8891f}\{f3f50225-6c77-4c6a-8356-ebc6a5f9d4f5}` |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Range** | `0` - `100` (percentage) |
| **Description** | Reserve battery level (hibernation file reserve) |

### Hibernate After (Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{238C9FA8-0AAD-41ED-83F4-97BE242C8F20}\{9d7815a6-7ee4-497e-8888-515a05f02364}` |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Unit** | Minutes |
| **Description** | Hibernate after (on battery) |

### Hibernate After (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Hibernate after (plugged in) |

### Allow Hybrid Sleep (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{238C9FA8-0AAD-41ED-83F4-97BE242C8F20}\{94ac6d29-73ce-41a6-809f-6363ba21b47e}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Allow hybrid sleep (plugged in) |

### Allow Hybrid Sleep (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Allow hybrid sleep (on battery) |

### Allow Wake Timers (Plugged In)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\{ActivePowerPlan}\{238C9FA8-0AAD-41ED-83F4-97BE242C8F20}\{bd3b718a-0680-4d9d-8ab2-e1d2b4ac806d}` |
| **Value** | `ACSettingIndex` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable, `2` = Important timers only |
| **Description** | Allow wake timers (plugged in) |

### Allow Wake Timers (On Battery)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | Same as above |
| **Value** | `DCSettingIndex` |
| **Type** | `REG_DWORD` |
| **Description** | Allow wake timers (on battery) |

---

## 🔒 Category 5: Privacy Settings

### Camera Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access camera |

### Microphone Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access microphone |

### Location Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access location |

### Notifications Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\userNotificationListener` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to read notifications |

### Account Info Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\userAccountInformation` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access account info |

### Contacts Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\contacts` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access contacts |

### Calendar Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\appointments` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access calendar |

### Call History Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\phoneCall` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access call history |

### Email Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\email` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access email |

### Messaging Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\chat` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to read/send messages |

### Radio Access (Bluetooth/Wi-Fi)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\radios` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to control radios |

### App Diagnostics Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\appDiagnostics` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access diagnostic info |

### Documents Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\documentsLibrary` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access Documents folder |

### Pictures Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\picturesLibrary` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access Pictures folder |

### Videos Access
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\videosLibrary` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access Videos folder |

### File System Access (Broad)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\broadFileSystemAccess` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Allow apps to access entire file system |

### Telemetry Level
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SOFTWARE\Policies\Microsoft\Windows\DataCollection` |
| **Value** | `AllowTelemetry` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Security, `1` = Basic, `3` = Full |
| **Description** | Diagnostic data collection level |

### Telemetry (Alternative Location)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection` |
| **Value** | `AllowTelemetry` |
| **Type** | `REG_DWORD` |
| **Values** | Same as above |
| **Description** | Alternative location for telemetry setting |

### Advertising ID
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo` |
| **Value** | `Enabled` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Allow apps to use advertising ID |

### Advertising ID (Policy)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo` |
| **Value** | `ID` |
| **Type** | `REG_SZ` |
| **Description** | The actual advertising ID (GUID) |

### Activity History
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Privacy` |
| **Value** | `ActivityHistoryEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Store activity history on this device |

### Publish Activity to Timeline
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Privacy` |
| **Value` | `PublishUserActivities` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Let Windows collect my activities from this PC |

### Collect Activity on Cloud
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Privacy` |
| **Value** | `UploadUserActivities` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Upload activity history to Microsoft |

### Suggested Content in Settings
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager` |
| **Value** | `SubscribedContent-338387Enabled` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Show suggested content in Settings app |

### Tailored Experiences
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Privacy` |
| **Value** | `TailoredExperiencesWithDiagnosticDataEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Let Microsoft use diagnostic data for tailored experiences |

### Inking & Typing Personalization
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\InputPersonalization` |
| **Value** | `RestrictImplicitTextCollection` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Enable, `1` = Disable |
| **Description** | Inking & typing data collection |

### Inking & Typing Dictionary
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\InputPersonalization` |
| **Value** | `RestrictImplicitInkCollection` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Enable, `1` = Disable |
| **Description** | Inking data collection |

### Find My Device
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SOFTWARE\Microsoft\Settings\FindMyDevice` |
| **Value** | `LocationSyncEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Find My Device feature |

### Camera Access for Desktop Apps
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam\NonPackaged` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Camera access for traditional desktop apps |

### Microphone Access for Desktop Apps
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Microphone access for traditional desktop apps |

### Location Access for Desktop Apps
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location\NonPackaged` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Values** | `Allow` = Allowed, `Deny` = Denied |
| **Description** | Location access for traditional desktop apps |

### Let Apps Use Camera Hardware
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Description** | Global camera access setting |

### Let Apps Use Microphone Hardware
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone` |
| **Value** | `Value` |
| **Type** | `REG_SZ` |
| **Description** | Global microphone access setting |

### Feedback Frequency
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Siuf\Rules` |
| **Value** | `NumberOfSIUFInPeriod` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Never, `1` = Automatically, etc. |
| **Description** | How often Windows asks for feedback |

---

## 🖱️ Category 6: Mouse & Keyboard Settings

### Mouse - Primary Button
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Mouse` |
| **Value** | `SwapMouseButtons` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Left-handed (primary right), `1` = Right-handed (primary left) |
| **Description** | Swap primary mouse button |

### Mouse - Double-Click Speed
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Mouse` |
| **Value** | `DoubleClickSpeed` |
| **Type** | `REG_SZ` |
| **Range** | `400` (fast) - `900` (slow) milliseconds |
| **Description** | Double-click speed threshold |

### Mouse - Double-Click Height
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Mouse` |
| **Value** | `DoubleClickHeight` |
| **Type** | `REG_SZ` |
| **Description** | Vertical tolerance for double-click (pixels) |

### Mouse - Double-Click Width
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Mouse` |
| **Value** | `DoubleClickWidth` |
| **Type** | `REG_SZ` |
| **Description** | Horizontal tolerance for double-click (pixels) |

### Mouse - Scroll Lines
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `WheelScrollLines` |
| **Type** | `REG_SZ` |
| **Values** | `-1` = One screen, `0` = No scrolling, `1`-`100` = Number of lines |
| **Description** | Lines to scroll per mouse wheel notch |

### Mouse - Scroll Wheel Delta
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value` | `WheelScrollChars` |
| **Type** | `REG_SZ` |
| **Description** | Characters to scroll horizontally per wheel notch |

### Mouse - Snap to Default Button
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Mouse` |
| **Value** | `SnapToDefaultButton` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Off, `1` = On |
| **Description** | Automatically move mouse to default dialog button |

### Mouse - Mouse Speed
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Mouse` |
| **Value** | `MouseSpeed` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Slow, `1` = Medium, `2` = Fast |
| **Description** | Mouse cursor speed (pointer speed) |

### Mouse - Mouse Threshold 1
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Mouse` |
| **Value** | `MouseThreshold1` |
| **Type** | `REG_SZ` |
| **Description** | Mouse acceleration threshold 1 |

### Mouse - Mouse Threshold 2
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Mouse` |
| **Value** | `MouseThreshold2` |
| **Type** | `REG_SZ` |
| **Description** | Mouse acceleration threshold 2 |

### Mouse - Enhance Pointer Precision
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Mouse` |
| **Value** | `MouseSensitivity` |
| **Type** | `REG_SZ` |
| **Values** | `1`-`20` (higher = more precise) |
| **Description** | Mouse sensitivity/pointer precision |

### Mouse - Cursor Scheme
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Cursors` |
| **Value** | `(Default)` |
| **Type** | `REG_SZ` |
| **Values** | `Windows Default`, `Windows Black`, `Windows Inverted`, etc. |
| **Description** | Mouse cursor scheme |

### Mouse - Cursor Paths
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Cursors` |
| **Values** | `Arrow`, `Hand`, `Wait`, `AppStarting`, etc. |
| **Type** | `REG_SZ` |
| **Description** | Paths to individual cursor files |

### Keyboard - Repeat Delay
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Keyboard` |
| **Value** | `KeyboardDelay` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Short (fast repeat), `3` = Long (slow repeat) |
| **Description** | Delay before keyboard repeat starts |

### Keyboard - Repeat Rate
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Keyboard` |
| **Value** | `KeyboardSpeed` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Slow, `31` = Fast |
| **Description** | Keyboard repeat rate (characters per second) |

### Touchpad - Tap to Click
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\PrecisionTouchPad` |
| **Value** | `TapEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Tap touchpad to click |

### Touchpad - Double-Tap to Drag
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\PrecisionTouchPad` |
| **Value** | `TapAndHoldEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Double-tap and drag on touchpad |

### Touchpad - Right-Click Zone
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\PrecisionTouchPad` |
| **Value** | `RightClickZoneEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Press bottom-right corner for right-click |

### Touchpad - Scrolling Direction
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\PrecisionTouchPad` |
| **Value** | `ScrollDirection` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Down motion scrolls down, `1` = Down motion scrolls up |
| **Description** | Touchpad scrolling direction |

### Touchpad - Sensitivity
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\PrecisionTouchPad` |
| **Value** | `CursorSpeed` |
| **Type** | `REG_DWORD` |
| **Range** | `0` - `4` (higher = more sensitive) |
| **Description** | Touchpad cursor speed |

### Touchpad - Three-Finger Gestures
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\PrecisionTouchPad` |
| **Value** | `ThreeFingerTapEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable (Cortana), `2` = Enable (Action Center), `3` = Enable (Play/Pause), `4` = Enable (Middle Click) |
| **Description** | Three-finger tap gesture |

### Touchpad - Four-Finger Gestures
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\PrecisionTouchPad` |
| **Value** | `FourFingerTapEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable (Action Center), `2` = Enable (Play/Pause), `3` = Enable (Cortana) |
| **Description** | Four-finger tap gesture |

### Touchpad - Swipe Gestures
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\PrecisionTouchPad` |
| **Value** | `SwipeGestureEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Enable three-finger swipes for multitasking |

---

## 🌐 Category 7: Regional & Language Settings

### Date Format - Short Date
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `sShortDate` |
| **Type** | `REG_SZ` |
| **Example** | `MM/dd/yyyy`, `dd/MM/yyyy`, `yyyy-MM-dd` |
| **Description** | Short date format |

### Date Format - Long Date
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `sLongDate` |
| **Type** | `REG_SZ` |
| **Example** | `dddd, MMMM dd, yyyy` |
| **Description** | Long date format |

### Time Format
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `sTimeFormat` |
| **Type** | `REG_SZ` |
| **Values** | `HH:mm:ss` = 24-hour, `h:mm:ss tt` = 12-hour |
| **Description** | Time format (12/24 hour) |

### Time Format with AM/PM
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `sTime` |
| **Type** | `REG_SZ` |
| **Description** | Time separator and format |

### AM Designator
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `sAM` |
| **Type** | `REG_SZ` |
| **Example** | `AM`, `a.m.` |
| **Description** | AM designator string |

### PM Designator
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `sPM` |
| **Type** | `REG_SZ` |
| **Example** | `PM`, `p.m.` |
| **Description** | PM designator string |

### First Day of Week
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `iFirstDayOfWeek` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Monday, `1` = Tuesday, `2` = Wednesday, `3` = Thursday, `4` = Friday, `5` = Saturday, `6` = Sunday |
| **Description** | First day of the week |

### Decimal Symbol
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `sDecimal` |
| **Type** | `REG_SZ` |
| **Values** | `.` or `,` |
| **Description** | Decimal separator symbol |

### Thousand Separator
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `sThousand` |
| **Type** | `REG_SZ` |
| **Values** | `,`, `.`, ` ` (space) |
| **Description** | Digit grouping/thousands separator |

### List Separator
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `sList` |
| **Type** | `REG_SZ` |
| **Values** | `,`, `;` |
| **Description** | List item separator |

### Currency Symbol
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `sCurrency` |
| **Type** | `REG_SZ` |
| **Example** | `$`, `€`, `£`, `¥` |
| **Description** | Local currency symbol |

### Currency Positive Format
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `iCurrency` |
| **Type** | `REG_SZ` |
| **Values** | `0` = $1, `1` = 1$, `2` = $ 1, `3` = 1 $ |
| **Description** | Positive currency format |

### Currency Negative Format
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `iNegCurr` |
| **Type** | `REG_SZ` |
| **Description** | Negative currency format (0-15) |

### Number of Decimal Digits
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `iDigits` |
| **Type** | `REG_SZ` |
| **Range** | `0` - `3` |
| **Description** | Number of decimal digits |

### Leading Zeros
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `iLZero` |
| **Type** | `REG_SZ` |
| **Values** | `0` = No leading zeros, `1` = Show leading zeros |
| **Description** | Show leading zeros in decimal numbers |

### Measurement System
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International` |
| **Value** | `iMeasure` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Metric, `1` = US |
| **Description** | Measurement system (Metric/US) |

### Time Zone
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\TimeZoneInformation` |
| **Value** | `TimeZoneKeyName` |
| **Type** | `REG_SZ` |
| **Example** | `Eastern Standard Time`, `Pacific Standard Time` |
| **Description** | Windows time zone name |

### Daylight Saving Time Enabled
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\TimeZoneInformation` |
| **Value** | `DynamicDaylightTimeDisabled` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Enable, `1` = Disable |
| **Description** | Automatically adjust for daylight saving |

### Country/Region
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\International\Geo` |
| **Value** | `Nation` |
| **Type** | `REG_SZ` |
| **Example** | `US`, `CA`, `GB` |
| **Description** | Country/region code |

---

## 🔧 Category 8: Accessibility Settings

### Narrator - Auto Start
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Narrator\NoRoam` |
| **Value** | `NarratorAutoStart` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Start Narrator automatically |

### Narrator - Voice
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Narrator\NoRoam` |
| **Value** | `CurrentVoice` |
| **Type** | `REG_SZ` |
| **Description** | Narrator voice selection |

### Narrator - Speed
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Narrator\NoRoam` |
| **Value** | `SpeechRate` |
| **Type** | `REG_DWORD` |
| **Range** | `0` - `15` (higher = faster) |
| **Description** | Narrator speech rate |

### Narrator - Volume
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Narrator\NoRoam` |
| **Value** | `SpeechVolume` |
| **Type** | `REG_DWORD` |
| **Range** | `0` - `100` |
| **Description** | Narrator volume |

### Magnifier - Zoom Level
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\ScreenMagnifier` |
| **Value** | `Magnification` |
| **Type** | `REG_DWORD` |
| **Range** | `100` - `1600` (percentage) |
| **Description** | Magnifier zoom level |

### Magnifier - Follow Focus
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\ScreenMagnifier` |
| **Value** | `FollowFocus` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Magnifier follows keyboard focus |

### Magnifier - Follow Mouse
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\ScreenMagnifier` |
| **Value** | `FollowMouse` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Magnifier follows mouse cursor |

### Magnifier - Follow Caret
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\ScreenMagnifier` |
| **Value** | `FollowCaret` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Magnifier follows text insertion point |

### Magnifier - Docked Mode
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\ScreenMagnifier` |
| **Value** | `WindowMode` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Full screen, `2` = Lens, `3` = Docked |
| **Description** | Magnifier view mode |

### High Contrast - Enable
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Accessibility\HighContrast` |
| **Value** | `Flags` |
| **Type** | `REG_SZ` |
| **Values** | `126` = On, `122` = Off |
| **Description** | Enable high contrast mode |

### High Contrast - Theme
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Accessibility\HighContrast` |
| **Value** | `High Contrast Scheme` |
| **Type** | `REG_SZ` |
| **Example** | `High Contrast #1`, `High Contrast #2` |
| **Description** | High contrast color scheme |

### Sticky Keys
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Accessibility\StickyKeys` |
| **Value** | `Flags` |
| **Type** | `REG_SZ` |
| **Values** | `506` = On, `58` = Off |
| **Description** | Enable Sticky Keys (press one key at a time for shortcuts) |

### Filter Keys
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Accessibility\Keyboard Response` |
| **Value** | `Flags` |
| **Type** | `REG_SZ` |
| **Values** | `122` = On, `58` = Off |
| **Description** | Enable Filter Keys (ignore brief/repeated keystrokes) |

### Toggle Keys
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Accessibility\ToggleKeys` |
| **Value** | `Flags` |
| **Type** | `REG_SZ` |
| **Values** | `62` = On, `58` = Off |
| **Description** | Hear tones when pressing Caps Lock, Num Lock, Scroll Lock |

### Mouse Keys
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Accessibility\MouseKeys` |
| **Value** | `Flags` |
| **Type** | `REG_SZ` |
| **Values** | `159` = On, `126` = Off |
| **Description** | Control mouse with keyboard numpad |

### Mouse Keys - Speed
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Accessibility\MouseKeys` |
| **Value** | `MouseSpeed` |
| **Type** | `REG_SZ` |
| **Description** | Mouse Keys cursor speed |

### Mouse Keys - Acceleration
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Accessibility\MouseKeys` |
| **Value** | `MouseTimeToMaxSpeed` |
| **Type** | `REG_SZ` |
| **Description** | Time to reach maximum Mouse Keys speed |

### Closed Captions
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Accessibility\Configuration` |
| **Value** | `ClosedCaptioning` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Enable closed captions |

### Caption Color
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows NT\CurrentVersion\Accessibility\Caption` |
| **Value** | `CaptionTextColor` |
| **Type** | `REG_DWORD` |
| **Description** | Closed caption text color |

### Caption Background Color
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows NT\CurrentVersion\Accessibility\Caption` |
| **Value** | `CaptionBackgroundColor` |
| **Type** | `REG_DWORD` |
| **Description** | Closed caption background color |

### Caption Font Size
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows NT\CurrentVersion\Accessibility\Caption` |
| **Value** | `CaptionFontSize` |
| **Type** | `REG_DWORD` |
| **Description** | Closed caption font size |

### Visual Notifications for Sound
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Accessibility\SoundSentry` |
| **Value** | `Flags` |
| **Type** | `REG_SZ` |
| **Values** | `2` = Flash active window, `6` = Flash desktop, etc. |
| **Description** | Show visual alerts for sounds (Sound Sentry) |

### Text Cursor - Thickness
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `CaretWidth` |
| **Type** | `REG_DWORD` |
| **Range** | `1` - `5` (pixels) |
| **Description** | Text cursor (caret) width |

### Text Cursor - Blink Rate
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `CursorBlinkRate` |
| **Type** | `REG_SZ` |
| **Range** | `200` - `1200` (milliseconds) |
| **Description** | Cursor blink rate (lower = faster) |

---

## 🎮 Category 9: Gaming Settings

### Game Mode
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\GameBar` |
| **Value** | `AllowAutoGameMode` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Enable Windows Game Mode |

### Game Bar - Enable
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\GameBar` |
| **Value** | `ShowStartupPanel` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Enable Game Bar |

### Game Bar - Shortcut
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\GameBar` |
| **Value** | `UseNexusForGameBar` |
| **Type** | `REG_DWORD` |
| **Description** | Use Nexus for Game Bar |

### Game Bar - Controller Shortcut
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\GameBar` |
| **Value** | `GamingMsguiEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Open Game Bar with Xbox button on controller |

### Captures - Record in Background
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\GameDVR` |
| **Value** | `HistoricalCaptureEnabled` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Record in background while playing |

### Captures - Max Recording Length
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\GameDVR` |
| **Value` | `HistoricalCaptureMaxDurationMS` |
| **Type** | `REG_DWORD` |
| **Description** | Maximum background recording length |

### Captures - Audio Quality
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\GameDVR` |
| **Value** | `AudioEncodingBitrate` |
| **Type** | `REG_DWORD` |
| **Description** | Audio bitrate for game captures |

### Captures - Video Quality
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\GameDVR` |
| **Value** | `VideoEncodingBitrate` |
| **Type** | `REG_DWORD` |
| **Description** | Video bitrate for game captures |

### Captures - Framerate
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\GameDVR` |
| **Value** | `VideoEncodingFrameRateMode` |
| **Type** | `REG_DWORD` |
| **Description** | Capture framerate (30/60 fps) |

### Captures - Save Location
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\GameDVR` |
| **Value** | `CaptureRootDirectory` |
| **Type** | `REG_SZ` |
| **Description** | Folder path for game captures |

### TruePlay (Anti-Cheat)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\GameBar` |
| **Value** | `UseTruePlay` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Enable TruePlay anti-cheat for games |

---

## 🖥️ Category 10: System & Performance

### Visual Effects - Adjust for Best Performance
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects` |
| **Value** | `VisualFXSetting` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Let Windows choose, `1` = Best appearance, `2` = Best performance, `3` = Custom |
| **Description** | Visual effects performance setting |

### Show Windows Contents While Dragging
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `DragFullWindows` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Show window contents while dragging |

### Smooth Edges of Screen Fonts
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `FontSmoothing` |
| **Type** | `REG_SZ` |
| **Values** | `0` = Disable, `2` = Enable |
| **Description** | Smooth edges of screen fonts (ClearType) |

### Animate Controls and Elements
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Control Panel\Desktop` |
| **Value** | `UserPreferencesMask` |
| **Type** | `REG_BINARY` |
| **Description** | Complex binary - controls various animation settings |

### Show Shadows Under Mouse
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `ShowInfoTip` |
| **Type** | `REG_DWORD` |
| **Values** | `1` = Enable, `0` = Disable |
| **Description** | Show shadows under mouse pointer |

### Show Thumbnails Instead of Icons
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `IconsOnly` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Show thumbnails, `1` = Show icons only |
| **Description** | Always show icons, never thumbnails |

### Save Taskbar Thumbnail Previews
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_CURRENT_USER` |
| **Key** | `Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced` |
| **Value** | `TaskbarThumbnails` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Show taskbar thumbnail previews |

### Virtual Memory Size
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management` |
| **Value** | `PagingFiles` |
| **Type** | `REG_MULTI_SZ` |
| **Example** | `C:\pagefile.sys 1024 4096` (initial max) |
| **Description** | Page file configuration |

### Disable Pagefile
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management` |
| **Value** | `DisablePagingExecutive` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Enable, `1` = Disable |
| **Description** | Disable pagefile for system executables |

### Large System Cache
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management` |
| **Value** | `LargeSystemCache` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = Disable, `1` = Enable |
| **Description** | Use large system cache (server optimization) |

### Processor Scheduling - Programs/Background
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\PriorityControl` |
| **Value** | `Win32PrioritySeparation` |
| **Type** | `REG_DWORD` |
| **Description** | Control foreground/background process priority |

### DEP (Data Execution Prevention)
| Property | Value |
|----------|-------|
| **Hive** | `HKEY_LOCAL_MACHINE` |
| **Key** | `SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management` |
| **Value** | `ExecuteOptions` |
| **Type** | `REG_DWORD` |
| **Values** | `0` = OptIn, `1` = OptOut, `2` = AlwaysOn, `3` = AlwaysOff |
| **Description** | DEP configuration |

---

## 📝 Important Notes

### ⚠️ Safety Warnings
1. **Always backup the registry** before making changes
2. **Create a system restore point** before importing settings
3. **Test on a VM** before applying to production systems
4. **Some settings require admin rights** - run TweakBackup as Administrator
5. **Some settings require reboot** to take effect

### 🔄 Windows Version Compatibility
| Setting Category | Windows 10 | Windows 11 |
|-----------------|------------|------------|
| System Appearance | ✅ Full | ✅ Full |
| File Explorer | ✅ Full | ✅ Full |
| Taskbar & Start | ⚠️ Some changes | ✅ Full |
| Power Settings | ✅ Full | ✅ Full |
| Privacy | ✅ Full | ✅ Full |
| Mouse/Keyboard | ✅ Full | ✅ Full |
| Regional | ✅ Full | ✅ Full |
| Accessibility | ✅ Full | ✅ Full |
| Gaming | ⚠️ Partial | ✅ Full |
| Performance | ✅ Full | ✅ Full |

### 📂 Registry Hive Reference
| Hive | Abbreviation | Scope |
|------|--------------|-------|
| `HKEY_CURRENT_USER` | HKCU | Current user only |
| `HKEY_LOCAL_MACHINE` | HKLM | All users, requires admin |
| `HKEY_USERS` | HKU | All user profiles |
| `HKEY_CLASSES_ROOT` | HKCR | File associations, COM objects |

### 💾 Export Format Example (JSON)
```json
{
  "profile_name": "My Gaming Settings",
  "windows_version": "Windows 11 Pro 23H2",
  "export_date": "2024-01-15T14:30:00",
  "settings": {
    "appearance_dark_mode": {
      "hive": "HKEY_CURRENT_USER",
      "key": "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
      "value": "AppsUseLightTheme",
      "type": "REG_DWORD",
      "data": 0
    },
    "taskbar_alignment": {
      "hive": "HKEY_CURRENT_USER",
      "key": "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced",
      "value": "TaskbarAl",
      "type": "REG_DWORD",
      "data": 0
    }
  }
}
