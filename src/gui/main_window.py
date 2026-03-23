"""
Main window for WinSet application - Professional Configuration Tool Design
"""
import subprocess

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
import os
import copy
import threading
import time
from typing import Optional, Dict, Any, List

# Import managers
from src.storage.exporter import ProfileExporter
from src.storage.importer import ProfileImporter
from src.presets.preset_manager import PresetManager
from src.utils.backup_manager import BackupManager
from src.core.registry_handler import RegistryHandler
from src.core.powershell_handler import PowerShellHandler
from src.core.setting_loader import SettingLoader
from src.core.history_manager import HistoryManager
from src.models.setting import RegistrySetting, SettingCategory, SettingType, Setting


class MainWindow:
    """Main application window with professional configuration tool design"""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("WinSet - Windows Configuration Toolkit")
        
        # Missing attribute declarations
        self.bg_color: str = ""
        self.fg_color: str = ""
        self.accent_color: str = ""
        self.border_color: str = ""
        self.main_frame: Optional[ttk.Frame] = None
        self.notebook: Optional[ttk.Notebook] = None
        self.home_frame: Optional[ttk.Frame] = None
        self.presets_frame: Optional[ttk.Frame] = None
        self.manual_frame: Optional[ttk.Frame] = None
        self.status_frame: Optional[ttk.Frame] = None
        self.status_label: Optional[ttk.Label] = None
        self.progress_bar: Optional[ttk.Progressbar] = None
        self.search_var: Optional[tk.StringVar] = None
        self.manual_canvas: Optional[tk.Canvas] = None
        self.manual_scrollable: Optional[ttk.Frame] = None
        self.manual_window_id: Optional[int] = None
        self.manual_vars: Dict[str, Any] = {}
        self.home_canvas: Optional[tk.Canvas] = None
        self.home_scrollable: Optional[ttk.Frame] = None
        self.home_window_id: Optional[int] = None
        self.root.geometry("1100x750")
        self.root.minsize(1000, 700)
        
        # Initialize Managers
        self.preset_manager = PresetManager()
        self.importer = ProfileImporter()
        self.exporter = ProfileExporter()
        self.backup_manager = BackupManager()
        self.registry_handler = RegistryHandler()
        self.powershell_handler = PowerShellHandler()
        self.setting_loader = SettingLoader()
        
        self.history_manager = HistoryManager()
        
        # UI State
        self.is_busy: bool = False
        self.expanded_settings: Dict[str, Any] = {}  # Track which settings are expanded
        self.expanded_setting_id: Optional[str] = None
        self.manual_row_widgets: Dict[str, Any] = {}
        self.restart_pending: bool = False
        
        self._setup_ui()
        self.center_window()
        
    def center_window(self):
        """Center window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def _setup_ui(self):
        """Create main UI layout with professional styling"""
        # Use clam theme for better compatibility
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except:
            try:
                style.theme_use('default')
            except:
                pass
        
        # Professional styling with Windows 11 Fluent Design
        bg_color = "#fafafa"  # Light gray background (Windows 11)
        fg_color = "#323130"  # Dark text (Windows 11)
        accent_color = "#0078d4"  # Windows 11 blue
        border_color = "#e1dfdd"  # Light border (Windows 11)
        tab_bg = "#f3f2f1"  # Tab background (Windows 11)
        tab_selected = "#ffffff"  # Selected tab background

        # Store colors for use in other methods
        self.bg_color = bg_color
        self.fg_color = fg_color
        self.accent_color = accent_color
        self.border_color = border_color
        
        # Configure styles - all text black for maximum readability
        # The 'clam' theme gives labels a gray default background; force it to the app surface.
        style.configure("TFrame", background=bg_color)
        style.configure("TLabel", background=bg_color, foreground="black", font=("Segoe UI", 9))
        style.configure("Header.TLabel", background=bg_color, font=("Segoe UI", 18, "bold"), foreground="black")
        style.configure("Bold.TLabel", background=bg_color, font=("Segoe UI", 10, "bold"), foreground="black")
        style.configure("Setting.TLabel", background=bg_color, font=("Segoe UI", 11, "bold"), foreground="black")
        style.configure("Description.TLabel", background=bg_color, font=("Segoe UI", 9), foreground="black")
        
        style.configure("TButton", 
                       background="#ffffff", 
                       foreground="black", 
                       font=("Segoe UI", 9),
                       padding=8,
                       borderwidth=1,
                       relief="solid")
        style.map("TButton",
                 background=[("active", "#e9ecef"), ("pressed", accent_color)],
                 foreground=[("pressed", "white")])
        
        style.configure("Accent.TButton", 
                       background=accent_color, 
                       foreground="white", 
                       font=("Segoe UI", 9, "bold"),
                       padding=8,
                       borderwidth=0,
                       relief="flat")
        style.map("Accent.TButton",
                 background=[("active", "#0056b3"), ("pressed", "#004085")])
        
        style.configure("TCheckbutton", 
                       background=bg_color, 
                       foreground="black", 
                       font=("Segoe UI", 9))
        
        style.configure("TEntry", 
                       fieldbackground="white", 
                       borderwidth=1,
                       relief="solid",
                       font=("Segoe UI", 9))

        style.configure(
            "TLabelframe",
            background=bg_color,
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "TLabelframe.Label",
            background=bg_color,
            foreground="black",
            font=("Segoe UI", 9, "bold"),
        )
        
        style.configure("TNotebook", background=bg_color, borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            background=tab_bg,
            foreground=fg_color,
            padding=[24, 12],
            font=("Segoe UI", 10, "normal"),
            borderwidth=0,
            relief="flat",
            focuscolor="none"
        )
        # Windows 11 style mapping with consistent padding to prevent shrinking
        style.map(
            "TNotebook.Tab",
            background=[("selected", tab_selected), ("active", "#edebe9")],
            foreground=[("selected", accent_color), ("active", fg_color)],
            padding=[("selected", [24, 12]), ("active", [24, 12])]
        )
        
        style.configure(
            "TProgressbar",
            background=accent_color,
            troughcolor="#ffffff",
            borderwidth=0,
        )
        
        style.configure("TCombobox", 
                       fieldbackground="white",
                       borderwidth=1,
                       relief="solid",
                       font=("Segoe UI", 9))
        
        style.configure("TScale", 
                       background=bg_color,
                       troughcolor="#ffffff",
                       borderwidth=1,
                       relief="solid")
        style.configure("Horizontal.TScale", background=bg_color)
        
        # Apply theme to root window
        self.root.configure(bg=bg_color)

        # Main container
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # Notebook for tabs
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create tabs
        self.home_frame = ttk.Frame(self.notebook)
        self.presets_frame = ttk.Frame(self.notebook)
        self.manual_frame = ttk.Frame(self.notebook)
        self.history_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.home_frame, text="Home")
        self.notebook.add(self.presets_frame, text="Presets")
        self.notebook.add(self.manual_frame, text="Manual Configuration")
        self.notebook.add(self.history_frame, text="History")

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self._create_home_tab()
        self._create_presets_tab()
        self._create_manual_tab()
        self._create_history_tab()

        # Status bar removed

        def _create_home_tab(self):
            """Create home tab content with scrolling support"""
            container = ttk.Frame(self.home_frame)
            container.pack(fill=tk.BOTH, expand=True)

            # Scrollable content area
            scroll_frame = ttk.Frame(container)
            scroll_frame.pack(fill=tk.BOTH, expand=True)

            self.home_canvas = tk.Canvas(scroll_frame, highlightthickness=0, bg=self.bg_color)
            home_scroll = ttk.Scrollbar(scroll_frame, orient="vertical", command=self.home_canvas.yview)
            self.home_scrollable = ttk.Frame(self.home_canvas)

            self.home_scrollable.bind("<Configure>", lambda e: self.home_canvas.configure(scrollregion=self.home_canvas.bbox("all")))
            self.home_window_id = self.home_canvas.create_window((0, 0), window=self.home_scrollable, anchor="nw")
            self.home_canvas.configure(yscrollcommand=home_scroll.set)

            self.home_canvas.bind(
                "<Configure>",
                lambda e: self.home_canvas.itemconfigure(self.home_window_id, width=e.width),
            )
            
            self.home_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            home_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Bind scroll events FIRST
            self._bind_home_scroll_events()
            
            # Content
            welcome_frame = ttk.LabelFrame(self.home_scrollable, text="Welcome", padding=20)
            welcome_frame.pack(fill=tk.X, padx=25, pady=(10, 15))
            
            welcome_label = ttk.Label(welcome_frame, text="WinSet", style="Header.TLabel")
            welcome_label.pack(anchor="w", pady=(0, 10))
            
            desc_label = ttk.Label(welcome_frame, 
                                text="Windows Configuration Toolkit\nEasily backup, restore, and optimize your Windows experience.", 
                                font=("Segoe UI", 12))
            desc_label.pack(anchor="w", pady=(0, 15))
            
            actions_frame = ttk.LabelFrame(self.home_scrollable, text="Quick Actions", padding=20)
            actions_frame.pack(fill=tk.X, padx=25, pady=(0, 15))
            
            ttk.Button(actions_frame, text="📤 Export Settings", command=self.export_settings, width=20).pack(fill=tk.X, pady=(0, 10))
            ttk.Button(actions_frame, text="📥 Import Settings", command=self.import_settings, width=20).pack(fill=tk.X, pady=(0, 10))
            ttk.Button(actions_frame, text="🔄 Create Restore Point", command=self.create_restore_point, width=20).pack(fill=tk.X, pady=(0, 10))
            
            tools_frame = ttk.LabelFrame(self.home_scrollable, text="System Tools", padding=20)
            tools_frame.pack(fill=tk.X, padx=25, pady=(0, 25))
            
            ttk.Button(tools_frame, text="⚙️ Control Panel", command=lambda: self.open_system_tool("control"), width=20).pack(fill=tk.X, pady=(0, 10))
            ttk.Button(tools_frame, text="🔧 Services", command=lambda: self.open_system_tool("services.msc"), width=20).pack(fill=tk.X, pady=(0, 10))
            ttk.Button(tools_frame, text="🖥️ MSConfig", command=lambda: self.open_system_tool("msconfig.exe")).pack(fill=tk.X, pady=(0, 10))
            ttk.Button(tools_frame, text="📁 Task Manager", command=lambda: self.open_system_tool("taskmgr"), width=20).pack(fill=tk.X, pady=(0, 10))
            ttk.Button(tools_frame, text="🔍 Programs & Features", command=lambda: self.open_system_tool("appwiz.cpl"), width=20).pack(fill=tk.X, pady=(0, 10))
            ttk.Button(tools_frame, text="🌐 Network Settings", command=lambda: self.open_system_tool("ncpa.cpl"), width=20).pack(fill=tk.X, pady=(0, 10))

        # src/gui/main_window.py

    def _create_presets_tab(self):
        """Create presets tab content"""
        container = ttk.Frame(self.presets_frame)
        container.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)
        
        ttk.Label(container, text="Configuration Presets", style="Header.TLabel").pack(anchor="w", pady=(0, 20))
        
        # Action Bar
        action_frame = ttk.Frame(container)
        action_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Button(action_frame, text="\u2795 Create Custom Preset", command=self.create_custom_preset, style="Accent.TButton").pack(side=tk.LEFT)
        
        # Presets grid
        presets_frame = ttk.Frame(container)
        presets_frame.pack(fill=tk.BOTH, expand=True)
        
        # Default presets
        presets = [
            {"id": "gaming", "title": "\U0001f3ae Gaming Mode", "desc": "Optimize system for gaming performance"},
            {"id": "developer", "title": "\U0001f4bb Developer Mode", "desc": "Configure for development workflows"},
            {"id": "privacy", "title": "\U0001f512 Privacy Max", "desc": "Maximum privacy and security settings"},
            {"id": "performance", "title": "\u26a1 Peak Performance", "desc": "Unlock full system performance"},
            {"id": "battery", "title": "\U0001f50b Battery Saver", "desc": "Optimize for extended battery life"}
        ]
        
        # Dynamically load additional presets
        existing_ids = {p["id"] for p in presets}
        # Use get_preset_list() to get all preset IDs
        try:
            preset_list = self.preset_manager.get_preset_list()
            for preset_id in preset_list:
                if preset_id not in existing_ids:
                    info = self.preset_manager.get_preset_info(preset_id)
                    if info:
                        presets.append({
                            "id": preset_id,
                            "title": info.get("name", preset_id),
                            "desc": info.get("description", "")
                        })
        except AttributeError:
            # Fallback for older versions or if method doesn't exist
            pass
        
        # Create preset buttons
        for i, preset in enumerate(presets):
            row = i // 3
            col = i % 3
            frame = ttk.Frame(presets_frame, relief=tk.RAISED, borderwidth=1)
            frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
            
            ttk.Label(frame, text=preset["title"], font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=(10, 5))
            ttk.Label(frame, text=preset["desc"], wraplength=200).pack(anchor="w", padx=10, pady=(0, 10))
            
            ttk.Button(frame, text="Apply Preset", 
                    command=lambda p=preset["id"]: self.apply_preset(p)).pack(pady=(0, 10))
        
        # Configure grid weights
        for i in range(3):
            presets_frame.columnconfigure(i, weight=1)
        
        # Dynamically load additional presets
        existing_ids = {p["id"] for p in presets}
        # Use get_preset_list() to get all preset IDs
        try:
            preset_list = self.preset_manager.get_preset_list()
            for preset_id in preset_list:
                if preset_id not in existing_ids:
                    info = self.preset_manager.get_preset_info(preset_id)
                    if info:
                        presets.append({
                            "id": preset_id,
                            "title": info.get("name", preset_id),
                            "desc": info.get("description", "")
                        })
        except AttributeError:
            # Fallback for older versions or if method doesn't exist
            pass

    def _create_manual_tab(self):
        """Create manual configuration tab"""
        container = ttk.Frame(self.manual_frame)
        container.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)
        
        # Header with search
        header_frame = ttk.Frame(container)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(header_frame, text="Manual Configuration", style="Header.TLabel").pack(side=tk.LEFT)
        
        self.search_var = tk.StringVar()
        search_frame = ttk.Frame(header_frame)
        search_frame.pack(side=tk.RIGHT)
        
        ttk.Label(search_frame, text="🔍 Search:").pack(side=tk.LEFT, padx=(0, 5))
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side=tk.LEFT)
        
        # Scrollable settings area
        scroll_frame = ttk.Frame(container)
        scroll_frame.pack(fill=tk.BOTH, expand=True)

        self.manual_canvas = tk.Canvas(scroll_frame, highlightthickness=0, bg=self.bg_color)
        manual_scroll = ttk.Scrollbar(scroll_frame, orient="vertical", command=self.manual_canvas.yview)
        self.manual_scrollable = ttk.Frame(self.manual_canvas)

        self.manual_scrollable.bind("<Configure>", lambda e: self.manual_canvas.configure(scrollregion=self.manual_canvas.bbox("all")))
        self.manual_window_id = self.manual_canvas.create_window((0, 0), window=self.manual_scrollable, anchor="nw")
        self.manual_canvas.configure(yscrollcommand=manual_scroll.set)

        # Make the inner frame always match the canvas width (no wasted space)
        self.manual_canvas.bind(
            "<Configure>",
            lambda e: self.manual_canvas.itemconfigure(self.manual_window_id, width=e.width),
        )
        
        self.manual_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        manual_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind events
        self._bind_scroll_events()
        search_entry.bind("<KeyRelease>", self._on_search)

    def _create_history_tab(self):
        """Create the UI for the Change History tab."""
        container = ttk.Frame(self.history_frame)
        container.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

        header_frame = ttk.Frame(container)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(header_frame, text="Change History", style="Header.TLabel").pack(side=tk.LEFT)
        ttk.Button(header_frame, text="🔄 Refresh", command=self._refresh_history_tab).pack(side=tk.RIGHT)

        tree_frame = ttk.Frame(container)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("Timestamp", "Setting", "Old Value", "New Value")
        self.history_tree = ttk.Treeview(tree_frame, columns=cols, show='headings')
        for col in cols:
            self.history_tree.heading(col, text=col)
        self.history_tree.column("Timestamp", width=160, anchor='w')
        self.history_tree.column("Setting", width=250, anchor='w')
        self.history_tree.column("Old Value", width=200, anchor='w')
        self.history_tree.column("New Value", width=200, anchor='w')

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)

        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        footer_frame = ttk.Frame(container)
        footer_frame.pack(fill=tk.X, pady=(15, 0))
        ttk.Button(footer_frame, text="↩️ Revert Selected Change", command=self._revert_selected_change, style="Accent.TButton").pack(side=tk.LEFT)
        ttk.Label(footer_frame, text="Note: History logging is currently active for the Manual Configuration tab only.", style="Description.TLabel").pack(side=tk.RIGHT)

    def _refresh_history_tab(self):
        """Clear and reload the history treeview with the latest data."""
        if not hasattr(self, 'history_tree'): return
        
        for i in self.history_tree.get_children():
            self.history_tree.delete(i)
        
        history_data = self.history_manager.get_history()
        for item in history_data:
            self.history_tree.insert("", "end", iid=item[0], values=item[1:])

    def _revert_selected_change(self):
        """Revert the currently selected change in the history treeview."""
        if not hasattr(self, 'history_tree'): return
        selected_items = self.history_tree.selection()
        if not selected_items:
            messagebox.showwarning("No Selection", "Please select a change from the list to revert.")
            return

        change_id = selected_items[0]
        details = self.history_manager.get_change_details(int(change_id))

        if not details:
            messagebox.showerror("Error", "Could not retrieve change details to perform revert.")
            return

        hive, key_path, value_name, value_type, old_value_str = details

        if messagebox.askyesno("Confirm Revert", f"Are you sure you want to revert the change to '{value_name}'?"):
            # Convert old_value back to its original type
            final_val = old_value_str
            if old_value_str != 'N/A':
                if value_type == "REG_DWORD":
                    try:
                        final_val = int(old_value_str)
                    except (ValueError, TypeError):
                        messagebox.showerror("Revert Error", f"Cannot convert old value '{old_value_str}' to an integer for REG_DWORD.")
                        return

            if self.registry_handler.write_value(hive, key_path, value_name, value_type, final_val):
                messagebox.showinfo("Success", f"Successfully reverted '{value_name}'.")
                self._refresh_history_tab()
            else:
                messagebox.showerror("Error", f"Failed to write to the registry to revert '{value_name}'.")

    def _bind_home_scroll_events(self):
        """Bind scroll events to home tab - simplified approach like manual tab"""
        self.home_canvas.bind("<MouseWheel>", self._on_home_mouse_wheel)
        self.home_frame.bind("<MouseWheel>", self._on_home_mouse_wheel)
        self.home_scrollable.bind("<MouseWheel>", self._on_home_mouse_wheel)
        
        # Also bind to the main window for global scroll handling
        self.root.bind("<MouseWheel>", self._on_global_mouse_wheel)

    def _bind_scroll_events(self):
        """Bind scroll events to work properly"""
        self.manual_canvas.bind("<MouseWheel>", self._on_mouse_wheel)
        self.manual_frame.bind("<MouseWheel>", self._on_mouse_wheel)
        self.manual_scrollable.bind("<MouseWheel>", self._on_mouse_wheel)
        
        # Home tab scroll events - use dedicated method
        self._bind_home_scroll_events()
        
        self.notebook.bind("<MouseWheel>", self._on_global_mouse_wheel)
        self.main_frame.bind("<MouseWheel>", self._on_global_mouse_wheel)
        self.manual_canvas.bind("<Map>", self._rebind_scroll_events)

    def _rebind_scroll_events(self, event=None):
        """Rebind scroll events to include new widgets"""
        def bind_recursive(widget):
            widget.bind("<MouseWheel>", self._on_mouse_wheel)
            for child in widget.winfo_children():
                bind_recursive(child)
        
        bind_recursive(self.manual_scrollable)
        
        # For home tab, just rebind the main containers
        if self.home_scrollable:
            self.home_scrollable.bind("<MouseWheel>", self._on_home_mouse_wheel)

    def _on_global_mouse_wheel(self, event):
        """Handle global mouse wheel events"""
        selected_tab = self.notebook.tab(self.notebook.select(), "text")
        if "Manual Configuration" in selected_tab:
            self._on_mouse_wheel(event)
        elif "Home" in selected_tab:
            self._on_home_mouse_wheel(event)

    def _on_tab_changed(self, event):
        """Handle tab change events."""
        selected_tab = self.notebook.tab(self.notebook.select(), "text")
        if "Manual Configuration" in selected_tab:
            self.refresh_manual_config()
            self.root.after(100, self._rebind_scroll_events)
        elif "Home" in selected_tab:
            # Rebind home tab scroll events when switching to home tab
            self.root.after(100, self._bind_home_scroll_events)
        elif "History" in selected_tab:
            self._refresh_history_tab()

    def _on_search(self, event):
        """Handle search box updates."""
        self.refresh_manual_config()

    def _on_mouse_wheel(self, event):
        """Handles mouse wheel scrolling for the manual config canvas."""
        if self.manual_canvas.winfo_exists():
            self.manual_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_home_mouse_wheel(self, event):
        """Handles mouse wheel scrolling for the home tab canvas."""
        if self.home_canvas and self.home_canvas.winfo_exists():
            # Fixed scroll direction - positive delta should scroll down
            self.home_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def refresh_manual_config(self):
        """Refresh the list of manual settings"""
        for widget in self.manual_scrollable.winfo_children():
            widget.destroy()
            
        search_query = self.search_var.get().lower()
        self.manual_vars = {}
        self.manual_row_widgets = {}

        for category in self.setting_loader.get_categories():
            settings = self.setting_loader.get_settings_for_category(category)
            filtered = [s for s in settings if search_query in s.name.lower() or search_query in (s.description or "").lower()]
            
            if not filtered: continue
                
            # Category header (flat, modern; avoids LabelFrame shading)
            category_container = ttk.Frame(self.manual_scrollable)
            category_container.pack(fill=tk.X, pady=(14, 6), padx=10)

            ttk.Label(
                category_container,
                text=category.value.replace("_", " ").title(),
                style="Bold.TLabel",
            ).pack(anchor="w")
            ttk.Separator(category_container, orient="horizontal").pack(fill=tk.X, pady=(6, 0))

            category_frame = ttk.Frame(self.manual_scrollable)
            category_frame.pack(fill=tk.X, pady=(6, 0), padx=10)

            for setting in filtered:
                self._create_setting_row(category_frame, setting)
        
        self.root.after(50, self._rebind_scroll_events)

    def _create_setting_row(self, parent, setting):
        """Create a row for a setting with professional controls"""
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, pady=4)
        
        # Check current state
        current_val = self.registry_handler.read_value(setting.hive, setting.key_path, setting.value_name)
        
        # Setting name (clickable)
        setting_id = f"setting_{setting.id}"
        is_expanded = (self.expanded_setting_id == setting_id)
        
        # Create expandable setting frame
        setting_frame = ttk.Frame(row_frame)
        setting_frame.pack(fill=tk.X)
        
        # Header with triangle and name
        header_frame = ttk.Frame(setting_frame)
        header_frame.pack(fill=tk.X)
        
        
        # Triangle indicator
        triangle = "▼" if is_expanded else "▶"
        triangle_label = ttk.Label(header_frame, text=triangle, font=("Segoe UI", 10), foreground=self.accent_color)
        triangle_label.pack(side=tk.LEFT, padx=(0, 8))
        
        # Setting name
        name_label = ttk.Label(header_frame, text=setting.name, style="Setting.TLabel", 
                            cursor="hand2")
        name_label.pack(side=tk.LEFT)
        
        # Bind click event to toggle expansion
        for widget in [triangle_label, name_label]:
            widget.bind("<Button-1>", lambda e, sid=setting_id: self._toggle_setting_expansion(sid))
        
        # Options frame (initially hidden)
        options_frame = ttk.Frame(setting_frame)
        if is_expanded:
            options_frame.pack(fill=tk.X, pady=(12, 0), padx=(25, 0))
        
        # Parse setting options
        options = self._parse_setting_options(setting, current_val)
        
        buttons = []
        slider_var = None
        if self._should_use_slider(setting, options):
            slider_var = self._create_slider_control(options_frame, setting, current_val, options, setting_id)
        elif self._should_use_text_entry(setting, options):
            buttons = self._create_text_control(options_frame, setting, current_val, options, setting_id)
        elif len(options) > 2:
            buttons = self._create_button_controls(options_frame, setting, current_val, options, setting_id)
        else:
            buttons = self._create_simple_controls(options_frame, setting, current_val, options, setting_id)

        if setting.description:
            ttk.Label(options_frame, text=setting.description, style="Description.TLabel", wraplength=700, justify=tk.LEFT).pack(anchor="w", pady=(8, 0))

        self.manual_row_widgets[setting_id] = {
            "setting": setting,
            "triangle": triangle_label,
            "options_frame": options_frame,
            "buttons": buttons,
            "slider_var": slider_var,
        }

    def _should_use_slider(self, setting, options):
        """Determine if setting should use slider control"""
        if getattr(setting, "is_range", False):
            return True

        # Only use sliders for real numeric ranges (e.g. "1-20") or known range settings.
        values_str = getattr(setting, "values", None)
        if isinstance(values_str, str):
            # Match things like "1-20" or "0 - 100" or "400 (fast) - 900 (slow)"
            import re
            if re.search(r"\b\d+\s*.*?-.*?\s*\d+\b", values_str):
                return True

        # Known registry-backed ranges
        if str(setting.value_name).lower() in {"mousesensitivity"}:
            return True

        return False

    def _should_use_text_entry(self, setting, options):
        """Determine if setting should use simple text entry control"""
        if self._should_use_slider(setting, options):
            return False
        if len(options) > 0:
            return False
        return True

    def _create_slider_control(self, parent, setting, current_val, options, setting_id, callback=None):
        """Create slider control for range-based settings"""
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(control_frame, text="Adjust:", style="Description.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        
        # Determine slider range from setting.values when possible
        import re
        values_str = str(getattr(setting, "values", "") or "")
        m = re.search(r"\b(\d+)\s*.*?-.*?\s*(\d+)\b", values_str)
        if m:
            from_val = int(m.group(1))
            to_val = int(m.group(2))
        else:
            from_val = 0
            to_val = 100

        # Create slider
        try:
            current_num = float(current_val) if current_val is not None else float(from_val)
        except (TypeError, ValueError):
            current_num = float(from_val)

        slider_var = tk.DoubleVar(value=current_num)
        slider = ttk.Scale(
            control_frame,
            from_=from_val,
            to=to_val,
            variable=slider_var,
            orient=tk.HORIZONTAL,
            length=300,
        )

        if from_val == 1 and to_val == 20 and str(setting.value_name).lower() in {"mousesensitivity"}:
            ttk.Label(control_frame, text="(Low ← → High)", style="Description.TLabel").pack(side=tk.LEFT, padx=(10, 0))
        
        slider.pack(side=tk.LEFT, padx=(0, 10))
        
        if callback:
            slider.bind("<ButtonRelease-1>", lambda e, s=setting, v=slider_var: callback(s, v.get()))
        else:
            slider.bind("<ButtonRelease-1>", lambda e, sid=setting_id, s=setting, v=slider_var: self._apply_slider_value(sid, s, v))

        return slider_var

    def _create_button_controls(self, parent, setting, current_val, options, setting_id, callback=None):
        """Create button controls for multi-option settings"""
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(control_frame, text="Select:", style="Description.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        
        buttons_frame = ttk.Frame(control_frame)
        buttons_frame.pack(side=tk.LEFT)
        
        buttons = []
        for option_name, option_value in options.items():
            is_selected = str(current_val) == str(option_value)
            
            option_btn = ttk.Button(buttons_frame, text=option_name,
                                style=("Accent.TButton" if is_selected else "TButton"),
                                width=12)
            option_btn.pack(side=tk.LEFT, padx=(0, 5))
            
            if callback:
                option_btn.bind("<Button-1>", lambda e, s=setting, v=option_value, btn=option_btn: callback(s, v, btn))
            else:
                option_btn.bind("<Button-1>", lambda e, sid=setting_id, s=setting, v=option_value: self._apply_setting_value(sid, s, v))
            buttons.append((option_btn, option_value))

        return buttons

    def _create_simple_controls(self, parent, setting, current_val, options, setting_id, callback=None):
        """Create simple controls for basic settings"""
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        buttons = []
        for option_name, option_value in options.items():
            is_selected = str(current_val) == str(option_value)
            
            option_btn = ttk.Button(control_frame, text=option_name,
                                style=("Accent.TButton" if is_selected else "TButton"),
                                width=15)
            option_btn.pack(side=tk.LEFT, padx=(0, 8))
            
            if callback:
                option_btn.bind("<Button-1>", lambda e, s=setting, v=option_value, btn=option_btn: callback(s, v, btn))
            else:
                option_btn.bind("<Button-1>", lambda e, sid=setting_id, s=setting, v=option_value: self._apply_setting_value(sid, s, v))
            buttons.append((option_btn, option_value))

        return buttons

    def _create_text_control(self, parent, setting, current_val, options, setting_id, callback=None):
        """Create simple text entry control for unbounded settings"""
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(control_frame, text="Value:", style="Description.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        
        entry_var = tk.StringVar(value=str(current_val) if current_val is not None else "")
        entry = ttk.Entry(control_frame, textvariable=entry_var, width=30)
        entry.pack(side=tk.LEFT, padx=(0, 10))
        
        def apply_action():
            if callback:
                callback(setting, entry_var.get())
            else:
                self._apply_setting_value(setting_id, setting, entry_var.get())

        apply_btn = ttk.Button(
            control_frame, 
            text="Set" if callback else "Apply", 
            command=apply_action, 
            width=10
        )
        apply_btn.pack(side=tk.LEFT)
        return []

    def _parse_setting_options(self, setting, current_val):
        """Parse setting options from values field or create default options"""
        options = {}
        
        # Check if setting has explicit options defined
        if hasattr(setting, 'options') and setting.options:
            return setting.options
        
        # Parse from values field in settings.json
        if hasattr(setting, 'values') and setting.values:
            values_str = str(setting.values)

            # If it's a numeric range (e.g. "1-20"), it's a slider setting, not discrete options.
            import re
            if re.search(r"\b\d+\s*-\s*\d+\b", values_str):
                return {}

            # Split on common separators.
            # Example formats:
            # - "0 = Left, 1 = Center"
            # - "0 = Off; 1 = On"
            # - "0=Slow,1=Medium,2=Fast"
            parts_list = re.split(r"[,;]", values_str)
            for option in parts_list:
                if '=' not in option:
                    continue
                left, right = option.split('=', 1)
                value = left.strip()
                name = right.strip()
                if value and name:
                    options[name] = value
        else:
            # Default boolean options
            if setting.value_type == "REG_DWORD":
                options = {
                    "Enable": "1",
                    "Disable": "0"
                }
            elif setting.value_type == "REG_SZ":
                options = {
                    "Enable": "1",
                    "Disable": "0"
                }
        
        return options

    def _toggle_setting_expansion(self, setting_id):
        """Toggle expansion state of a setting"""
        if self.expanded_setting_id == setting_id:
            self._collapse_setting_row(setting_id)
            self.expanded_setting_id = None
            return

        if self.expanded_setting_id is not None:
            self._collapse_setting_row(self.expanded_setting_id)

        self.expanded_setting_id = setting_id
        self._expand_setting_row(setting_id)

    def _collapse_setting_row(self, setting_id):
        row = self.manual_row_widgets.get(setting_id)
        if not row:
            return
        row["triangle"].config(text="▶")
        row["options_frame"].pack_forget()

    def _expand_setting_row(self, setting_id):
        row = self.manual_row_widgets.get(setting_id)
        if not row:
            return
        row["triangle"].config(text="▼")
        row["options_frame"].pack(fill=tk.X, pady=(12, 0), padx=(25, 0))

    def _apply_and_log_change(self, setting, new_value):
        """Shared logic to apply a setting, log it, and check for restart."""
        old_value = self.registry_handler.read_value(setting.hive, setting.key_path, setting.value_name)

        # Type conversion
        if setting.value_type == "REG_SZ":
            final_val = str(new_value)
        elif setting.value_type == "REG_DWORD":
            final_val = int(new_value)
        else:
            final_val = new_value

        success = self.registry_handler.write_value(setting.hive, setting.key_path, setting.value_name, setting.value_type, final_val)
        
        if success:
            self.history_manager.log_change(setting, old_value, final_val)
            if getattr(setting, 'requires_restart', False):
                self.restart_pending = True
            self.update_status(f"Updated: {setting.name}")
        else:
            messagebox.showerror("Error", f"Failed to update {setting.name}")
        
        return success

    def _apply_slider_value(self, setting_id, setting, var):
        """Apply slider value to setting"""
        new_val_num = int(round(var.get()))
        success = self._apply_and_log_change(setting, new_val_num)
        if success:
            self._update_row_selection_state(setting_id)

    def _apply_setting_value(self, setting_id, setting, value):
        """Apply a specific value to a setting"""
        success = self._apply_and_log_change(setting, value)
        if success:
            self._update_row_selection_state(setting_id)

    def _update_row_selection_state(self, setting_id):
        row = self.manual_row_widgets.get(setting_id)
        if not row:
            return

        setting = row.get("setting")
        if not setting:
            return

        current_val = self.registry_handler.read_value(setting.hive, setting.key_path, setting.value_name)
        for btn, opt_val in row.get("buttons", []) or []:
            is_selected = str(current_val) == str(opt_val)
            btn.configure(style=("Accent.TButton" if is_selected else "TButton"))

    def update_status(self, text, progress=None):
        if not self.status_label:
            return

        self.status_label.config(text=text)
        if progress is not None:
            if self.progress_bar:
                self.progress_bar.pack(side=tk.RIGHT, padx=15, pady=8)
                self.progress_bar.config(value=progress * 100)
        else:
            if self.progress_bar:
                self.progress_bar.pack_forget()
        self.root.update_idletasks()

    def run_async(self, func, *args, **kwargs):
        if self.is_busy: return
        self.is_busy = True
        
        def wrapper():
            try:
                func(*args, **kwargs)
            finally:
                self.is_busy = False
                self.update_status("Ready")
                if self.restart_pending:
                    self.root.after(100, self._prompt_for_restart)
                    self.restart_pending = False
                
        threading.Thread(target=wrapper, daemon=True).start()

    def _apply_settings_list(self, settings_list: List[Setting], profile_name: str):
        """Applies a list of setting objects to the system."""
        from src.models.profile import Profile
        temp_profile = Profile(name=profile_name)
        temp_profile.settings = {s.id: s for s in settings_list}
        
        # Check for restart requirement
        for setting in temp_profile.settings.values():
            if getattr(setting, 'requires_restart', False):
                self.restart_pending = True
                break

        def task():
            self.update_status("Creating Restore Point...", 0.2)
            self.backup_manager.create_restore_point(f"WinSet Before Apply: {profile_name}")
            self.update_status("Applying custom settings...", 0.5)
            
            results = temp_profile.apply_all(safe_mode=False)
            
            self.update_status("Apply Complete", 1.0)
            messagebox.showinfo("Success", f"Applied {sum(1 for v in results.values() if v)} settings.")
        
        self.run_async(task)

    def _save_settings_list_to_file(self, settings_list: List[Setting]):
        """Shows a save dialog and exports a list of settings to a JSON file."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON files", "*.json")],
            title="Save Custom Preset", initialdir=os.path.join(os.getcwd(), "presets")
        )
        if not file_path: return
        name = os.path.basename(file_path).replace(".json", "")
        if self.exporter.export_profile(settings_list, file_path, name):
            messagebox.showinfo("Success", "Preset saved successfully!")
        else:
            messagebox.showerror("Error", "Failed to save preset.")

    def _prompt_for_restart(self):
        """Show a dialog recommending a system restart."""
        msg = "Some of the applied settings require a system restart to take full effect. Would you like to restart now?"
        if messagebox.askyesno("Restart Recommended", msg):
            try:
                subprocess.run(["shutdown", "/r", "/t", "1"], check=True)
            except Exception as e:
                messagebox.showerror("Restart Failed", f"Could not restart the system automatically.\nPlease restart manually.\n\nError: {e}")

    def export_settings(self):
        # Logic moved to generic export_profile method on exporter, used here for full backup
        try:
            file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
            if not file_path: 
                return
            
            def task():
                self.update_status("Preparing export...", 0.1)
                settings = []
                all_cats = self.setting_loader.get_categories()
                
                for i, cat in enumerate(all_cats):
                    self.update_status(f"Reading {cat.value}...", 0.1 + (0.7 * (i/len(all_cats))))
                    category_settings = self.setting_loader.get_settings_for_category(cat)
                    
                    for s in category_settings:
                        val = self.registry_handler.read_value(s.hive, s.key_path, s.value_name)
                        if val is not None:
                            # Convert bytes to string if needed
                            if isinstance(val, bytes):
                                val = val.decode('utf-8', errors='replace')
                            s.value = val
                            settings.append(s)
            
                self.update_status("Saving...", 0.9)
                if self.exporter.export_profile(settings, file_path, "WinSet Export"):
                    self.update_status("Export Complete", 1.0)
                    messagebox.showinfo("Success", f"Profile saved to {os.path.basename(file_path)}")
                else:
                    messagebox.showerror("Export Failed", "Failed to export settings")
                    
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Export Error", f"Error exporting settings: {str(e)}")
        
        self.run_async(task)

    def create_custom_preset(self):
        """Launch wizard to create a custom preset"""
        self._show_preset_selection_step()

    def _show_preset_selection_step(self):
        """Step 1: Select settings to include"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Create Preset - Step 1: Select Settings")
        dialog.geometry("700x700")
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_dialog(dialog)

        # Header
        header = ttk.Frame(dialog, padding=20)
        header.pack(fill=tk.X)
        ttk.Label(header, text="Select Settings", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header, text="Choose which settings you want to include in your custom preset.", style="Description.TLabel").pack(anchor="w")

        # Content
        content = ttk.Frame(dialog)
        content.pack(fill=tk.BOTH, expand=True, padx=20)

        canvas = tk.Canvas(content, highlightthickness=0, bg=self.bg_color)
        scrollbar = ttk.Scrollbar(content, orient="vertical", command=canvas.yview)
        scrollable = ttk.Frame(canvas)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        window_id = canvas.create_window((0, 0), window=scrollable, anchor="nw")
        
        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))
        
        # Bind mousewheel for scrolling
        def _on_wheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_wheel)
        scrollable.bind("<MouseWheel>", _on_wheel)
        dialog.bind("<MouseWheel>", _on_wheel)

        # Populate settings
        vars_map = {} # setting_id -> (BooleanVar, SettingObject)
        
        for category in self.setting_loader.get_categories():
            settings = self.setting_loader.get_settings_for_category(category)
            if not settings: continue
            
            cat_frame = ttk.Frame(scrollable)
            cat_frame.pack(fill=tk.X, pady=(15, 5))
            ttk.Label(cat_frame, text=category.value.replace("_", " ").title(), style="Bold.TLabel").pack(anchor="w")
            
            for s in settings:
                row = ttk.Frame(scrollable)
                row.pack(fill=tk.X, padx=10, pady=2)
                var = tk.BooleanVar(value=False)
                vars_map[s.id] = (var, s)
                ttk.Checkbutton(row, text=s.name, variable=var).pack(anchor="w")

        # Footer
        footer = ttk.Frame(dialog, padding=20)
        footer.pack(fill=tk.X)
        
        def next_step():
            selected = [s for var, s in vars_map.values() if var.get()]
            if not selected:
                messagebox.showwarning("Selection Required", "Please select at least one setting.", parent=dialog)
                return
            dialog.destroy()
            self._show_preset_configuration_step(selected)

        ttk.Button(footer, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(footer, text="Next >", style="Accent.TButton", command=next_step).pack(side=tk.RIGHT, padx=5)

    def _show_preset_configuration_step(self, selected_settings):
        """Step 2: Configure values for selected settings"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Create Preset - Step 2: Configure Values")
        dialog.geometry("800x700")
        dialog.transient(self.root)
        dialog.grab_set()
        self.center_dialog(dialog)

        # Header
        header = ttk.Frame(dialog, padding=20)
        header.pack(fill=tk.X)
        ttk.Label(header, text="Configure Values", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header, text="Set the desired values for your preset.", style="Description.TLabel").pack(anchor="w")

        # Content
        content = ttk.Frame(dialog)
        content.pack(fill=tk.BOTH, expand=True, padx=20)

        canvas = tk.Canvas(content, highlightthickness=0, bg=self.bg_color)
        scrollbar = ttk.Scrollbar(content, orient="vertical", command=canvas.yview)
        scrollable = ttk.Frame(canvas)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)
        window_id = canvas.create_window((0, 0), window=scrollable, anchor="nw")
        
        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))

        # Bind mousewheel for scrolling
        def _on_wheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_wheel)
        scrollable.bind("<MouseWheel>", _on_wheel)
        dialog.bind("<MouseWheel>", _on_wheel)

        # State tracking
        preset_values = {} # setting_id -> value

        # Populate rows
        for setting in selected_settings:
            # Read current value as default
            current_val = self.registry_handler.read_value(setting.hive, setting.key_path, setting.value_name)
            if current_val is not None:
                if isinstance(current_val, bytes):
                    current_val = current_val.decode('utf-8', errors='replace')
                preset_values[setting.id] = current_val
            
            # Create Row
            row_frame = ttk.LabelFrame(scrollable, text=setting.name, padding=10)
            row_frame.pack(fill=tk.X, pady=5)
            
            options = self._parse_setting_options(setting, current_val)
            setting.is_preset_editor = True # Helper flag for text button label

            # Define update callback
            def update_value(s, v, btn=None):
                # Type conversion
                if s.value_type == "REG_SZ": final_v = str(v)
                elif s.value_type == "REG_DWORD": final_v = int(float(v))
                else: final_v = v
                
                preset_values[s.id] = final_v
                
                # Visual feedback for buttons
                if btn:
                    for child in btn.master.winfo_children():
                        if isinstance(child, ttk.Button):
                            child.configure(style="TButton")
                    btn.configure(style="Accent.TButton")
            
            # Reuse control creation logic with callback
            if self._should_use_slider(setting, options):
                self._create_slider_control(row_frame, setting, current_val, options, setting.id, callback=update_value)
            elif self._should_use_text_entry(setting, options):
                self._create_text_control(row_frame, setting, current_val, options, setting.id, callback=update_value)
            elif len(options) > 2:
                self._create_button_controls(row_frame, setting, current_val, options, setting.id, callback=update_value)
            else:
                self._create_simple_controls(row_frame, setting, current_val, options, setting.id, callback=update_value)

        # Footer
        footer = ttk.Frame(dialog, padding=20)
        footer.pack(fill=tk.X)

        def get_final_settings() -> List[Setting]:
            final_settings = []
            for s in selected_settings:
                if s.id in preset_values:
                    s_copy = copy.deepcopy(s)
                    s_copy.value = preset_values[s.id]
                    final_settings.append(s_copy)
            return final_settings

        def save_to_file():
            final_settings = get_final_settings()
            if not final_settings:
                messagebox.showwarning("No Values Configured", "Cannot save an empty preset.", parent=dialog)
                return
            dialog.destroy()
            self._save_settings_list_to_file(final_settings)

        def apply_settings():
            final_settings = get_final_settings()
            if not final_settings:
                messagebox.showwarning("No Values Configured", "Cannot apply an empty preset.", parent=dialog)
                return
            dialog.destroy()
            self._apply_settings_list(final_settings, "Custom Preset")
            if messagebox.askyesno("Save Preset?", "Settings applied. Would you also like to save these settings as a preset file for later use?"):
                self._save_settings_list_to_file(final_settings)

        ttk.Button(footer, text="Back", command=lambda: [dialog.destroy(), self._show_preset_selection_step()]).pack(side=tk.LEFT)
        ttk.Button(footer, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(footer, text="💾 Save to File...", command=save_to_file).pack(side=tk.RIGHT, padx=5)
        ttk.Button(footer, text="✅ Apply Settings", style="Accent.TButton", command=apply_settings).pack(side=tk.RIGHT, padx=5)

    def center_dialog(self, dialog):
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

    def import_settings(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not file_path: return
        
        success, msg, profile = self.importer.load_profile(file_path)
        if not success:
            messagebox.showerror("Import Error", msg)
            return
        
        self._show_import_selection_dialog(profile)

    def _show_import_selection_dialog(self, profile):
        """Show dialog to select specific settings from a profile"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Import: {profile.name}")
        dialog.geometry("600x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 300
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 300
        dialog.geometry(f"+{x}+{y}")
        dialog.configure(bg=self.bg_color)
        
        # Header
        header = ttk.Frame(dialog, padding=15)
        header.pack(fill=tk.X)
        ttk.Label(header, text="Select Settings to Import", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header, text=f"Profile: {profile.name}", style="Description.TLabel").pack(anchor="w")
        
        # Content
        content = ttk.Frame(dialog, padding=15)
        content.pack(fill=tk.BOTH, expand=True)
        
        # List/Tree
        tree_frame = ttk.Frame(content)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Use a canvas for the list to allow custom styling
        canvas = tk.Canvas(tree_frame, highlightthickness=0, bg=self.bg_color)
        scrollable = ttk.Frame(canvas)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.config(command=canvas.yview)
        canvas.config(yscrollcommand=tree_scroll.set)
        
        window_id = canvas.create_window((0, 0), window=scrollable, anchor="nw")
        
        def configure_scroll(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(window_id, width=event.width)
            
        scrollable.bind("<Configure>", configure_scroll)
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))
        
        # Bind mousewheel
        def _on_wheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind("<MouseWheel>", _on_wheel)
        scrollable.bind("<MouseWheel>", _on_wheel)
        dialog.bind("<MouseWheel>", _on_wheel)
        
        # Populate settings grouped by category
        vars_map = {}
        settings_by_cat = {}
        for s in profile.settings.values():
            # Handle category whether it's an object/enum or string
            cat = getattr(s, 'category', 'General')
            if hasattr(cat, 'value'): cat = cat.value
            cat_str = str(cat)
            
            if cat_str not in settings_by_cat:
                settings_by_cat[cat_str] = []
            settings_by_cat[cat_str].append(s)
            
        for cat, settings in settings_by_cat.items():
            cat_frame = ttk.Frame(scrollable)
            cat_frame.pack(fill=tk.X, pady=(10, 0))
            ttk.Label(cat_frame, text=cat.replace("_", " ").title(), style="Bold.TLabel").pack(anchor="w", padx=5)
            
            for s in settings:
                row = ttk.Frame(scrollable)
                row.pack(fill=tk.X, padx=15, pady=2)
                var = tk.BooleanVar(value=True)
                vars_map[s.id] = (var, s)
                ttk.Checkbutton(row, text=s.name, variable=var).pack(anchor="w")
                
        # Footer
        footer = ttk.Frame(dialog, padding=15)
        footer.pack(fill=tk.X)
        
        def apply():
            selected_settings = []
            for sid, (var, s) in vars_map.items():
                if var.get():
                    selected_settings.append(s)
            
            if not selected_settings:
                messagebox.showwarning("Warning", "No settings selected.")
                return
                
            dialog.destroy()
            
            # Apply only selected settings
            profile.settings = {s.id: s for s in selected_settings}
            
            if messagebox.askyesno("Confirm", f"Apply {len(selected_settings)} settings from '{profile.name}'?"):
                def task():
                    self.update_status("Creating Restore Point...", 0.2)
                    self.backup_manager.create_restore_point(f"WinSet Import: {profile.name}")
                    self.update_status("Applying...", 0.5)
                    results = self.importer.apply_profile(profile, safe_mode=False)
                    self.update_status("Import Complete", 1.0)
                    messagebox.showinfo("Success", f"Applied {sum(1 for v in results.values() if v)} settings.")
                self.run_async(task)
        
        ttk.Button(footer, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(footer, text="Confirm", style="Accent.TButton", command=apply).pack(side=tk.RIGHT, padx=5)
        
        def toggle_all(state):
            for var, _ in vars_map.values():
                var.set(state)
                
        ttk.Button(footer, text="Select All", command=lambda: toggle_all(True)).pack(side=tk.LEFT, padx=5)
        ttk.Button(footer, text="Deselect All", command=lambda: toggle_all(False)).pack(side=tk.LEFT, padx=5)

    def _refresh_history_tab(self):
        """Clear and reload the history treeview with the latest data."""
        for i in self.history_tree.get_children():
            self.history_tree.delete(i)
        
        history_data = self.history_manager.get_history()
        for item in history_data:
            self.history_tree.insert("", "end", iid=item[0], values=item[1:])

    def _revert_selected_change(self):
        """Revert the currently selected change in the history treeview."""
        selected_items = self.history_tree.selection()
        if not selected_items:
            messagebox.showwarning("No Selection", "Please select a change from the list to revert.")
            return

        change_id = selected_items[0]

        details = self.history_manager.get_change_details(int(change_id))
        if not details:
            messagebox.showerror("Error", "Could not retrieve change details to perform revert.")
            return

        hive, key_path, value_name, value_type, old_value_str = details

        if messagebox.askyesno("Confirm Revert", f"Are you sure you want to revert the change to '{value_name}'?"):
            # Convert old_value back to its original type
            final_val = old_value_str
            if old_value_str != "N/A":
                if value_type == "REG_DWORD":
                    try:
                        final_val = int(old_value_str)
                    except (ValueError, TypeError):
                        messagebox.showerror(
                            "Revert Error",
                            f"Cannot convert old value '{old_value_str}' to an integer for REG_DWORD.",
                        )
                        return

            if self.registry_handler.write_value(hive, key_path, value_name, value_type, final_val):
                messagebox.showinfo("Success", f"Successfully reverted '{value_name}'.")
                self._refresh_history_tab()
            else:
                messagebox.showerror("Error", f"Failed to write to the registry to revert '{value_name}'.")

    def apply_preset(self, preset_id):
        msg = (
            f"Are you sure you want to apply the '{preset_id.replace('_', ' ').title()}' preset?\n\n"
            "Applying a preset acts as an overlay: it only modifies the specific settings defined "
            "in the preset, leaving your other configurations unchanged.\n\n"
            "You can combine multiple presets by applying them one after another. "
            "If settings overlap, the most recently applied preset takes precedence."
        )
        
        if messagebox.askyesno("Confirm Preset Application", msg):
            # Check if any setting in the preset requires a restart
            _, _, profile = self.preset_manager.load_preset(preset_id)
            if profile:
                for setting in profile.settings:
                    if getattr(setting, 'requires_restart', False):
                        self.restart_pending = True
                        break

            def task():
                self.update_status("Creating Restore Point...", 0.1)
                self.backup_manager.create_restore_point(f"WinSet Before Preset: {preset_id}")
                self.update_status(f"Applying {preset_id}...", 0.4)
                success, msg, results = self.preset_manager.apply_preset(preset_id)
                if success:
                    self.update_status("Preset Applied", 1.0)
                    messagebox.showinfo("Success", f"Applied {sum(1 for v in results.values() if v)} settings.")
                else:
                    messagebox.showerror("Error", msg)
            self.run_async(task)

    def create_restore_point(self):
        def task():
            self.update_status("Creating Restore Point...", 0.5)
            success = self.backup_manager.create_restore_point("WinSet Manual Restore Point")
            if success:
                self.update_status("Restore Point Created", 1.0)
                messagebox.showinfo("Success", "System restore point created successfully.")
            else:
                messagebox.showerror("Error", "Failed to create restore point.")
        self.run_async(task)

    def open_system_tool(self, tool_name):
        """Open Windows system tools and utilities"""
        import subprocess
        import os
        
        try:
            # Use subprocess to run the system tool
            subprocess.Popen(tool_name, shell=True)
            self.update_status(f"Opened {tool_name}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open {tool_name}: {str(e)}")
