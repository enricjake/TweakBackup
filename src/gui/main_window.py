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
        self.expanded_setting_id: Optional[str] = None
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
        
        # Initialize category variables for category selection
        self.category_vars: Dict[str, tk.BooleanVar] = {}
        for category in self.setting_loader.get_categories():
            self.category_vars[category.name] = tk.BooleanVar(value=True)
        
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
    
    def select_all_categories(self):
        """Select all category checkboxes."""
        for var in self.category_vars.values():
            var.set(True)
    
    def clear_all_categories(self):
        """Clear all category checkboxes."""
        for var in self.category_vars.values():
            var.set(False)

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
        # Use tk.Frame instead of ttk.Frame for proper background color
        self.manual_scrollable = tk.Frame(self.manual_canvas, bg=self.bg_color)

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
        print(f"DEBUG: refresh_manual_config called, categories: {len(self.setting_loader.get_categories())}")
        for widget in self.manual_scrollable.winfo_children():
            widget.destroy()
            
        search_query = self.search_var.get().lower()
        self.manual_vars = {}
        self.manual_row_widgets = {}

        for category in self.setting_loader.get_categories():
            settings = self.setting_loader.get_settings_for_category(category)
            filtered = [s for s in settings if search_query in s.name.lower() or search_query in (s.description or "").lower()]
            
            if not filtered: continue
            print(f"DEBUG: Adding category {category.name} with {len(filtered)} settings")
                
            # Category header (flat, modern; avoids LabelFrame shading)
            category_container = tk.Frame(self.manual_scrollable, bg=self.bg_color)
            category_container.pack(fill=tk.X, pady=(14, 6), padx=10)

            ttk.Label(
                category_container,
                text=category.value.replace("_", " ").title(),
                style="Bold.TLabel",
            ).pack(anchor="w")
            ttk.Separator(category_container, orient="horizontal").pack(fill=tk.X, pady=(6, 0))

            category_frame = tk.Frame(self.manual_scrollable, bg=self.bg_color)
            category_frame.pack(fill=tk.X, pady=(6, 0), padx=10)

            for setting in filtered:
                self._create_setting_row(category_frame, setting)
        
        self.root.after(50, self._rebind_scroll_events)

    def _create_setting_row(self, parent, setting):
        """Create a row for a setting with professional controls"""
        row_frame = tk.Frame(parent, bg=self.bg_color)
        row_frame.pack(fill=tk.X, pady=4)
        
        # Check current state
        current_val = self.registry_handler.read_value(setting.hive, setting.key_path, setting.value_name)
        
        # Setting name (clickable)
        setting_id = f"setting_{setting.id}"
        is_expanded = (self.expanded_setting_id == setting_id)
        
        # Create expandable setting frame
        setting_frame = tk.Frame(row_frame, bg=self.bg_color)
        setting_frame.pack(fill=tk.X)
        
        # Header with triangle and name
        header_frame = tk.Frame(setting_frame, bg=self.bg_color)
        header_frame.pack(fill=tk.X)
        
        
        # Triangle indicator
        triangle = "▼" if is_expanded else "▶"
        triangle_label = ttk.Label(header_frame, text=triangle, font=("Segoe UI", 10), foreground=self.accent_color)
        triangle_label.pack(side=tk.LEFT, padx=(0, 8))
        
        # Setting name
        name_label = ttk.Label(header_frame, text=setting.name, style="Setting.TLabel", 
                            cursor="hand2")
        name_label.pack(side=tk.LEFT)
        
        # Current value display (user-friendly)
        friendly_label = self._get_friendly_label_for_value(setting, current_val)
        current_value_label = ttk.Label(header_frame, text=f"Current: {friendly_label}", 
                                      style="Description.TLabel", foreground=self.accent_color)
        current_value_label.pack(side=tk.LEFT, padx=(10, 0))
        
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

        # Add system settings link for complex settings
        if self._should_show_system_settings_link(setting):
            link_frame = ttk.Frame(options_frame)
            link_frame.pack(anchor="w", pady=(8, 0))
            
            link_btn = self._create_system_settings_link(link_frame, setting)
            link_btn.pack(side=tk.LEFT)
            
            # Add tooltip explaining the link
            self._create_tooltip(link_btn, f"Open Windows Settings to manage {setting.name}")

        self.manual_row_widgets[setting_id] = {
            "setting": setting,
            "triangle": triangle_label,
            "options_frame": options_frame,
            "buttons": buttons,
            "slider_var": slider_var,
            "current_value_label": current_value_label,  # Store reference for updates
        }

    def _should_use_slider(self, setting, options):
        """Determine if setting should use slider control"""
        if getattr(setting, "is_range", False):
            return True

        # Exclude specific settings from slider treatment - these should use buttons
        exclude_from_slider = [
            "wallpaperstyle",  # Wallpaper Style - discrete options, not a range
            "taskbarglomlevel",  # Taskbar button grouping
            "mmtaskbarmode",  # Multi-monitor taskbar
            "searchboxtaskbarmode",  # Search box appearance
            "launchto",  # File Explorer default location
            "autoscroll",  # Auto-scroll in IE/Edge
            "dragfullwindows",  # Drag full windows
            "listviewshadow",  # Show shadows under windows
            "showinfotip",  # Show pop-up descriptions
            "taskbartransparency",  # Taskbar transparency
            "enabletransparency",  # General transparency
            "colorprevalence",  # Accent color settings
            "settings",  # REG_BINARY settings like taskbar position
        ]
        
        setting_name_lower = str(setting.value_name).lower()
        setting_display_name_lower = str(setting.name).lower()
        
        if any(term in setting_name_lower for term in exclude_from_slider):
            return False

        # Only use sliders for real numeric ranges (e.g. "1-20") or known range settings.
        values = getattr(setting, "values", None)
        
        # If values is a dictionary, check if it's a range (numeric keys)
        if isinstance(values, dict):
            # Check if all keys are numeric (indicating a range)
            try:
                keys = [int(k) for k in values.keys()]
                # Don't use slider for small discrete option sets (<= 5 options)
                if len(keys) <= 5:
                    return False
                # More than 5 numeric options suggest a range/slider
                if len(keys) > 5:
                    return True
            except ValueError:
                pass
            return False
        
        # For string values, check for range patterns
        if isinstance(values, str):
            # Match things like "1-20" or "0 - 100" or "400 (fast) - 900 (slow)"
            import re
            if re.search(r"\b\d+\s*.*?-.*?\s*\d+\b", values):
                return True

        # Known registry-backed ranges and specific settings
        # Mouse sensitivity and speed settings
        if any(term in setting_name_lower for term in ["mousesensitivity", "mousespeed"]):
            return True
        if "mouse speed" in setting_display_name_lower or "mouse sensitivity" in setting_display_name_lower:
            return True
            
        # Double-click speed
        if "doubleclickspeed" in setting_name_lower or "double-click speed" in setting_display_name_lower:
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
            # Default ranges for specific settings
            setting_name_lower = str(setting.value_name).lower()
            setting_display_name_lower = str(setting.name).lower()
            
            if "mousesensitivity" in setting_name_lower or "mouse sensitivity" in setting_display_name_lower:
                from_val, to_val = 1, 20
            elif "mousespeed" in setting_name_lower or "mouse speed" in setting_display_name_lower:
                if "enhance pointer precision" in setting_display_name_lower:
                    from_val, to_val = 1, 20
                else:
                    from_val, to_val = 0, 2  # For basic mouse speed (Slow/Medium/Fast)
            elif "doubleclickspeed" in setting_name_lower or "double-click speed" in setting_display_name_lower:
                from_val, to_val = 200, 600  # Typical double-click speed range
            else:
                from_val, to_val = 0, 100

        # Create slider
        try:
            current_num = float(current_val) if current_val is not None else float(from_val)
        except (TypeError, ValueError):
            current_num = float(from_val)

        slider_var = tk.DoubleVar(value=current_num)
        
        # Current value label
        value_label = ttk.Label(control_frame, text=f"Current: {int(current_num)}", style="Description.TLabel")
        value_label.pack(side=tk.LEFT, padx=(10, 5))
        
        # Update value label when slider moves
        def update_label(*args):
            value_label.config(text=f"Current: {int(slider_var.get())}")
        
        slider_var.trace('w', update_label)
        
        slider = ttk.Scale(
            control_frame,
            from_=from_val,
            to=to_val,
            variable=slider_var,
            orient=tk.HORIZONTAL,
            length=250,
        )

        # Add descriptive labels for specific settings
        setting_display_name_lower = str(setting.name).lower()
        if from_val == 1 and to_val == 20 and ("mousesensitivity" in str(setting.value_name).lower() or "enhance pointer precision" in setting_display_name_lower):
            ttk.Label(control_frame, text="(Low ← → High)", style="Description.TLabel").pack(side=tk.LEFT, padx=(5, 0))
        elif from_val == 0 and to_val == 2 and "mouse speed" in setting_display_name_lower:
            ttk.Label(control_frame, text="(Slow ← → Fast)", style="Description.TLabel").pack(side=tk.LEFT, padx=(5, 0))
        elif "double-click speed" in setting_display_name_lower:
            ttk.Label(control_frame, text="(Slow ← → Fast)", style="Description.TLabel").pack(side=tk.LEFT, padx=(5, 0))
        elif "TaskbarAl" in str(setting.value_name):
            ttk.Label(control_frame, text="(Left ← → Center)", style="Description.TLabel").pack(side=tk.LEFT, padx=(5, 0))
        
        slider.pack(side=tk.LEFT, padx=(0, 10))
        
        if callback:
            slider.bind("<ButtonRelease-1>", lambda e, s=setting, v=slider_var: callback(s, v.get()))
        else:
            slider.bind("<ButtonRelease-1>", lambda e, sid=setting_id, s=setting, v=slider_var: self._apply_slider_value(sid, s, v))

        return slider_var

    def _parse_binary_setting_value(self, setting, current_val):
        """Parse REG_BINARY settings to extract meaningful values"""
        if not current_val:
            return None
        
        try:
            # Convert binary data to bytes if it's a string
            if isinstance(current_val, str):
                import binascii
                current_val = binascii.unhexlify(current_val.replace(' ', ''))
            
            # Taskbar position parsing (StuckRects3)
            if "StuckRects3" in setting.key_path:
                return self._parse_taskbar_binary(current_val)
            
            # Visual effects parsing (UserPreferencesMask)
            if "UserPreferencesMask" in str(setting.key_path):
                return self._parse_visual_effects_binary(current_val)
            
            # File Explorer view settings
            if "Explorer\\Streams" in setting.key_path:
                return self._parse_explorer_view_binary(current_val)
            
            # Start Menu layout
            if "CloudStore" in setting.key_path and "start.layout" in setting.key_path:
                return self._parse_start_menu_layout_binary(current_val)
            
            # Pinned taskbar apps
            if "Explorer\\Taskband" in setting.key_path:
                return self._parse_pinned_apps_binary(current_val, "taskbar")
            
            # Pinned Start Menu items
            if "Explorer\\StartPage2" in setting.key_path:
                return self._parse_pinned_apps_binary(current_val, "startmenu")
            
        except Exception:
            pass  # Return None if parsing fails
        
        return None

    def _parse_taskbar_binary(self, binary_data):
        """Parse taskbar binary data to extract position and auto-hide settings"""
        if len(binary_data) < 20:
            return None
        
        try:
            # Taskbar position is typically at offset 12 in the binary data
            # 0 = Bottom, 1 = Left, 2 = Top, 3 = Right
            position = binary_data[12] if len(binary_data) > 12 else 0
            position_map = {0: "Bottom", 1: "Left", 2: "Top", 3: "Right"}
            
            # Auto-hide is typically at offset 8
            # 0 = Auto-hide enabled, 2 = Auto-hide disabled
            auto_hide = binary_data[8] if len(binary_data) > 8 else 2
            auto_hide_map = {0: "Enable", 2: "Disable"}
            
            # Determine which setting this is based on the setting name
            setting_name = str(getattr(self, '_current_setting_name', '')).lower()
            
            if 'location' in setting_name or 'position' in setting_name:
                return position_map.get(position, "Bottom")
            elif 'auto-hide' in setting_name or 'autohide' in setting_name:
                return auto_hide_map.get(auto_hide, "Disable")
                
        except Exception:
            pass
        
        return None

    def _parse_visual_effects_binary(self, binary_data):
        """Parse visual effects binary data to extract animation settings"""
        if len(binary_data) < 8:
            return None
        
        try:
            # Visual effects are typically controlled by bits in the binary data
            # This is a simplified approach - actual bit positions may vary
            byte_3 = binary_data[3] if len(binary_data) > 3 else 0
            
            # Check specific bits for common visual effects
            animations_enabled = (byte_3 & 0x02) == 0  # Bit 1 controls animations
            controls_animated = (byte_3 & 0x04) == 0  # Bit 2 controls control animations
            
            setting_name = str(getattr(self, '_current_setting_name', '')).lower()
            
            if 'animation' in setting_name:
                return "Enable" if not animations_enabled else "Disable"
            elif 'control' in setting_name:
                return "Enable" if not controls_animated else "Disable"
                
        except Exception:
            pass
        
        return None

    def _parse_explorer_view_binary(self, binary_data):
        """Parse File Explorer view settings binary data"""
        if len(binary_data) < 10:
            return None
        
        try:
            # This is a simplified parser for Explorer view settings
            # The actual structure is complex, so we'll provide a descriptive result
            return "Custom view settings"
        except Exception:
            pass
        
        return None

    def _parse_start_menu_layout_binary(self, binary_data):
        """Parse Start Menu layout binary data"""
        if len(binary_data) < 10:
            return None
        
        try:
            # Start Menu layout is complex - provide descriptive result
            return "Custom Start Menu layout"
        except Exception:
            pass
        
        return None

    def _parse_pinned_apps_binary(self, binary_data, app_type):
        """Parse pinned apps binary data for taskbar or Start Menu"""
        if len(binary_data) < 10:
            return None
        
        try:
            # Pinned apps data is complex - count approximate number of pinned items
            # This is a simplified approach
            if app_type == "taskbar":
                return "Pinned taskbar apps configured"
            else:
                return "Pinned Start Menu items configured"
        except Exception:
            pass
        
        return None

    def _parse_multi_sz_setting_value(self, setting, current_val):
        """Parse REG_MULTI_SZ settings like Virtual Memory Size"""
        if not current_val:
            return "Not configured"
        
        try:
            # Handle Virtual Memory Size specifically
            if "Virtual Memory" in setting.name or "PagingFiles" in setting.value_name:
                return self._parse_virtual_memory_config(current_val)
            
            # For other multi-string values, join them
            if isinstance(current_val, (list, tuple)):
                return ", ".join(str(v) for v in current_val if v)
            elif isinstance(current_val, str):
                # Split on null characters if it's a single string with multiple values
                parts = current_val.split('\x00')
                return ", ".join(part for part in parts if part.strip())
                
        except Exception:
            pass
        
        return str(current_val)

    def _parse_virtual_memory_config(self, current_val):
        """Parse virtual memory configuration from registry"""
        try:
            if isinstance(current_val, str):
                # Format: "C:\pagefile.sys 1024 4096"
                parts = current_val.split()
                if len(parts) >= 3:
                    path = parts[0]
                    initial_size = parts[1]
                    max_size = parts[2]
                    
                    # Convert to MB if needed
                    try:
                        initial_mb = int(initial_size)
                        max_mb = int(max_size)
                        
                        if initial_mb >= 1024:
                            initial_display = f"{initial_mb // 1024} GB"
                        else:
                            initial_display = f"{initial_mb} MB"
                            
                        if max_mb >= 1024:
                            max_display = f"{max_mb // 1024} GB"
                        else:
                            max_display = f"{max_mb} MB"
                        
                        return f"{initial_display} - {max_display}"
                    except ValueError:
                        return f"{initial_size} - {max_size}"
                        
            elif isinstance(current_val, (list, tuple)):
                # Handle multiple pagefiles
                configs = []
                for config in current_val:
                    if config.strip():
                        parts = config.split()
                        if len(parts) >= 3:
                            path = parts[0]
                            initial_size = parts[1]
                            max_size = parts[2]
                            configs.append(f"{path}: {initial_size}-{max_size}")
                return "; ".join(configs) if configs else "Not configured"
                
        except Exception:
            pass
        
        return "Custom configuration"

    def _get_option_hint(self, setting, option_name):
        """Get hint text for a specific option"""
        if hasattr(setting, 'option_hints') and setting.option_hints:
            return setting.option_hints.get(option_name, "")
        return ""

    def _get_friendly_label_for_value(self, setting, current_val):
        """Get user-friendly label for a current registry value"""
        if current_val is None:
            return "Unknown"
        
        # Handle REG_BINARY settings specially
        if setting.value_type == "REG_BINARY":
            # Store setting name for binary parsing
            self._current_setting_name = setting.name
            parsed_value = self._parse_binary_setting_value(setting, current_val)
            if parsed_value:
                return parsed_value
        
        # Handle REG_MULTI_SZ settings specially (like Virtual Memory)
        if setting.value_type == "REG_MULTI_SZ":
            parsed_value = self._parse_multi_sz_setting_value(setting, current_val)
            if parsed_value:
                return parsed_value
        
        # Parse options to find the label for this value
        options = self._parse_setting_options(setting, current_val)
        for label, value in options.items():
            if str(current_val) == str(value):
                return label
        
        # If no friendly label found, return the raw value
        return str(current_val)

    def _launch_system_settings(self, setting):
        """Launch the appropriate Windows system settings for this setting"""
        import subprocess

        # ------------------------------------------------------------------ #
        # Full per-setting deep-link map.  All names match settings.json.      #
        # Fallback logic (category, then generic) follows below.              #
        # ------------------------------------------------------------------ #
        SETTINGS_MAP: dict[str, str] = {
            # ── System Appearance ──────────────────────────────────────────
            "Desktop Wallpaper":                        "ms-settings:personalization-background",
            "Wallpaper Style (Stretch/Fill/Fit)":       "ms-settings:personalization-background",
            "Tile Wallpaper":                           "ms-settings:personalization-background",
            "App Theme (Light/Dark)":                   "ms-settings:personalization-colors",
            "System Theme (Light/Dark)":                "ms-settings:personalization-colors",
            "Accent Color":                             "ms-settings:personalization-colors",
            "Accent Color (Auto-Selected)":             "ms-settings:personalization-colors",
            "Accent Color on Title Bars":               "ms-settings:personalization-colors",
            "Show Accent on Start/Taskbar":             "ms-settings:personalization-colors",
            "Transparency Effects":                     "ms-settings:personalization-colors",
            "Taskbar Transparency":                     "ms-settings:personalization-colors",
            "Desktop Icon Visibility - This PC":        "ms-settings:themes",
            "Desktop Icon Visibility - Network":        "ms-settings:themes",
            "Desktop Icon Visibility - Recycle Bin":    "ms-settings:themes",
            "Desktop Icon Visibility - User's Files":   "ms-settings:themes",
            "Desktop Icon Size":                        "ms-settings:display",
            "Desktop Icon Spacing (Horizontal)":        "ms-settings:display",
            "Desktop Icon Spacing (Vertical)":          "ms-settings:display",
            "Font Smoothing (ClearType)":               "ms-settings:display",
            "Font Smoothing Type":                      "ms-settings:display",
            "Visual Effects - Animations":              "ms-settings:display-advancedgraphics",
            "Show Shadows Under Windows":               "ms-settings:display-advancedgraphics",
            "Show Shadows Under Mouse":                 "ms-settings:display-advancedgraphics",

            # ── File Explorer Settings ─────────────────────────────────────
            # Explorer folder-options have no ms-settings URI; open via shell
            "Hidden Files":                             "__explorer_options__",
            "File Extensions":                         "__explorer_options__",
            "Protected Operating System Files":         "__explorer_options__",
            "Empty Drives":                             "__explorer_options__",
            "Navigation Pane - Show All Folders":       "__explorer_options__",
            "Navigation Pane - Expand to Open Folder":  "__explorer_options__",
            "Quick Access - Show Recent Files":         "ms-settings:privacy-general",
            "Quick Access - Show Frequent Folders":     "ms-settings:privacy-general",
            "File Explorer - Open to Quick Access":     "__explorer_options__",
            "Check Boxes to Select Items":              "__explorer_options__",
            "Item Check Boxes":                         "__explorer_options__",
            "File Explorer - View Type":                "__explorer_options__",
            "Folder Merge Conflicts":                   "__explorer_options__",
            "Sharing Wizard":                           "ms-settings:network-shareadvanced",
            "Recycle Bin - Delete Confirmation":        "__explorer_options__",
            "Recycle Bin - Maximum Size":               "__explorer_options__",
            "Recycle Bin - Files Removed Immediately":  "__explorer_options__",
            "Folder Tips":                              "__explorer_options__",
            "Sync Provider Notifications":              "ms-settings:privacy-general",
            "File Explorer - Display Size Info":        "__explorer_options__",
            "Encrypted/Compressed Files in Color":      "__explorer_options__",
            "Always Show Icons, Never Thumbnails":      "__explorer_options__",
            "Display File Icon on Thumbnails":          "__explorer_options__",
            "Separate Process for Folder Windows":      "__explorer_options__",

            # ── Taskbar & Start Menu ───────────────────────────────────────
            "Taskbar Alignment":                        "ms-settings:taskbar",
            "Taskbar Size (Small/Large Icons)":         "ms-settings:taskbar",
            "Taskbar Location on Screen":               "ms-settings:taskbar",
            "Auto-Hide Taskbar":                        "ms-settings:taskbar",
            "Lock Taskbar":                             "ms-settings:taskbar",
            "Taskbar Badges":                           "ms-settings:taskbar",
            "Taskbar Corner Overflow":                  "ms-settings:taskbar",
            "Taskbar Widgets":                          "ms-settings:taskbar",
            "Taskbar Chat Icon":                        "ms-settings:taskbar",
            "Taskbar Task View Button":                 "ms-settings:taskbar",
            "Taskbar People Button":                    "ms-settings:taskbar",
            "Taskbar Search Box":                       "ms-settings:taskbar",
            "Taskbar Corner Icons - Pen Menu":          "ms-settings:taskbar",
            "Taskbar Corner Icons - Touch Keyboard":    "ms-settings:taskbar",
            "Taskbar Corner Icons - Virtual Touchpad":  "ms-settings:taskbar",
            "Combine Taskbar Buttons":                  "ms-settings:taskbar",
            "Multiple Displays - Show Taskbar on All Displays":      "ms-settings:taskbar",
            "Multiple Displays - Where to Show Buttons":             "ms-settings:taskbar",
            "Multiple Displays - Combine Buttons on Other Taskbars": "ms-settings:taskbar",
            "Start Menu - Show Recently Added Apps":    "ms-settings:personalization-start",
            "Start Menu - Show Most Used Apps":         "ms-settings:personalization-start",
            "Start Menu - Show Suggestions":            "ms-settings:personalization-start",
            "Start Menu - Show Recently Opened Items":  "ms-settings:personalization-start",
            "Start Menu - Size":                        "ms-settings:personalization-start",
            "Start Menu Layout XML":                    "ms-settings:personalization-start",
            "Pinned Taskbar Apps":                      "ms-settings:taskbar",
            "Pinned Start Menu Apps":                   "ms-settings:personalization-start",

            # ── Power Settings ─────────────────────────────────────────────
            "Active Power Plan":                        "ms-settings:powersleep",
            "Monitor Timeout (Plugged In)":             "ms-settings:powersleep",
            "Monitor Timeout (On Battery)":             "ms-settings:powersleep",
            "Sleep Timeout (Plugged In)":               "ms-settings:powersleep",
            "Sleep Timeout (On Battery)":               "ms-settings:powersleep",
            "Hibernate After (Plugged In)":             "ms-settings:powersleep",
            "Hibernate After (On Battery)":             "ms-settings:powersleep",
            "Lid Close Action (Plugged In)":            "ms-settings:powersleep",
            "Lid Close Action (On Battery)":            "ms-settings:powersleep",
            "Power Button Action (Plugged In)":         "ms-settings:powersleep",
            "Power Button Action (On Battery)":         "ms-settings:powersleep",
            "Sleep Button Action (Plugged In)":         "ms-settings:powersleep",
            "Sleep Button Action (On Battery)":         "ms-settings:powersleep",
            "Fast Startup":                             "ms-settings:powersleep",
            "Hibernate":                                "ms-settings:powersleep",
            "USB Selective Suspend":                    "ms-settings:powersleep",
            "Processor Power (Minimum, Plugged In)":    "ms-settings:powersleep",
            "Processor Power (Minimum, On Battery)":    "ms-settings:powersleep",
            "Processor Power (Maximum, Plugged In)":    "ms-settings:powersleep",
            "Processor Power (Maximum, On Battery)":    "ms-settings:powersleep",
            "Processor Cooling Policy":                 "ms-settings:powersleep",
            "Screen Brightness (Plugged In)":           "ms-settings:display",
            "Screen Brightness (On Battery)":           "ms-settings:display",
            "Adaptive Brightness":                      "ms-settings:display",
            "Wireless Adapter Power Mode":              "ms-settings:powersleep",
            "PCI Express Link State":                   "ms-settings:powersleep",
            "Battery Low Level":                        "ms-settings:batterysaver",
            "Battery Critical Level":                   "ms-settings:batterysaver",
            "Battery Low Notification":                 "ms-settings:batterysaver",
            "Battery Critical Action":                  "ms-settings:batterysaver",
            "Battery Saver":                            "ms-settings:batterysaver",
            "Battery Saver Threshold":                  "ms-settings:batterysaver",
            "Allow Wake Timers (Plugged In)":           "ms-settings:powersleep",
            "Allow Wake Timers (On Battery)":           "ms-settings:powersleep",
            "Hard Disk Timeout (Plugged In)":           "ms-settings:powersleep",
            "Hard Disk Timeout (On Battery)":           "ms-settings:powersleep",
            "Sleep State":                              "ms-settings:powersleep",
            "Hybrid Sleep (Plugged In)":                "ms-settings:powersleep",
            "Hybrid Sleep (On Battery)":                "ms-settings:powersleep",
            "Video Playback (Plugged In)":              "ms-settings:powersleep",
            "Video Playback (On Battery)":              "ms-settings:powersleep",
            "Network Connectivity in Standby":          "ms-settings:powersleep",
            "Energy Saver Brightness Reduction":        "ms-settings:batterysaver",

            # ── Privacy Options ────────────────────────────────────────────
            "Advertising ID":                           "ms-settings:privacy-general",
            "Tailored Experiences":                     "ms-settings:privacy-general",
            "Diagnostics & Feedback Level":             "ms-settings:privacy-feedback",
            "Feedback Frequency":                       "ms-settings:privacy-feedback",
            "Inking & Typing Personalization":          "ms-settings:privacy-speechtyping",
            "Speech Recognition":                       "ms-settings:privacy-speech",
            "Timeline / Activity History":              "ms-settings:privacy-activityhistory",
            "Activity History - Enabled":               "ms-settings:privacy-activityhistory",
            "Publish User Activities":                  "ms-settings:privacy-activityhistory",
            "Upload User Activities":                   "ms-settings:privacy-activityhistory",
            "Location Services":                        "ms-settings:privacy-location",
            "Location Sync":                            "ms-settings:privacy-location",
            "Camera Access":                            "ms-settings:privacy-webcam",
            "Microphone Access":                        "ms-settings:privacy-microphone",
            "Notifications Access":                     "ms-settings:privacy-notifications",
            "Account Info Access":                      "ms-settings:privacy-accountinfo",
            "Contacts Access":                          "ms-settings:privacy-contacts",
            "Calendar Access":                          "ms-settings:privacy-calendar",
            "Call History Access":                      "ms-settings:privacy-callhistory",
            "Email Access":                             "ms-settings:privacy-email",
            "Tasks Access":                             "ms-settings:privacy-tasks",
            "Messaging Access":                         "ms-settings:privacy-messaging",
            "Radios Access":                            "ms-settings:privacy-radios",
            "Background Apps":                          "ms-settings:privacy-backgroundapps",
            "App Diagnostics":                          "ms-settings:privacy-appdiagnostics",
            "Automatic File Downloads":                 "ms-settings:privacy-automaticfiledownloads",
            "Documents Library Access":                 "ms-settings:privacy-documents",
            "Pictures Library Access":                  "ms-settings:privacy-pictures",
            "Videos Library Access":                    "ms-settings:privacy-videos",
            "Network Access":                           "ms-settings:privacy-customdevices",
            "Clipboard History":                        "ms-settings:clipboard",
            "Cloud Clipboard Sync":                     "ms-settings:clipboard",
            "Find My Device":                           "ms-settings:findmydevice",

            # ── Keyboard & Mouse ───────────────────────────────────────────
            "Mouse - Primary Button":                   "ms-settings:mousetouchpad",
            "Mouse - Double-Click Speed":               "ms-settings:mousetouchpad",
            "Mouse - Double-Click Height":              "ms-settings:mousetouchpad",
            "Mouse - Double-Click Width":               "ms-settings:mousetouchpad",
            "Mouse - Scroll Lines":                     "ms-settings:mousetouchpad",
            "Mouse - Scroll Wheel Delta":               "ms-settings:mousetouchpad",
            "Mouse - Snap to Default Button":           "ms-settings:mousetouchpad",
            "Mouse - Mouse Speed":                      "ms-settings:mousetouchpad",
            "Mouse - Mouse Threshold 1":                "ms-settings:mousetouchpad",
            "Mouse - Mouse Threshold 2":                "ms-settings:mousetouchpad",
            "Mouse - Enhance Pointer Precision":        "ms-settings:mousetouchpad",
            "Mouse - Cursor Scheme":                    "ms-settings:easeofaccess-cursorandpointersize",
            "Mouse - Cursor Size":                      "ms-settings:easeofaccess-cursorandpointersize",
            "Keyboard - Repeat Delay":                  "ms-settings:keyboard",
            "Keyboard - Repeat Rate":                   "ms-settings:keyboard",
            "Touchpad - Tap to Click":                  "ms-settings:devices-touchpad",
            "Touchpad - Double-Tap to Drag":            "ms-settings:devices-touchpad",
            "Touchpad - Right-Click Zone":              "ms-settings:devices-touchpad",
            "Touchpad - Scrolling Direction":           "ms-settings:devices-touchpad",
            "Touchpad - Sensitivity":                   "ms-settings:devices-touchpad",
            "Touchpad - Three-Finger Gestures":         "ms-settings:devices-touchpad",
            "Touchpad - Four-Finger Gestures":          "ms-settings:devices-touchpad",
            "Touchpad - Swipe Gestures":                "ms-settings:devices-touchpad",

            # ── System & Performance / Accessibility / Gaming ──────────────
            # Regional & Language
            "Date Format - Short Date":                 "ms-settings:dateandtime",
            "Date Format - Long Date":                  "ms-settings:dateandtime",
            "Time Format":                              "ms-settings:dateandtime",
            "Time Format with AM/PM":                   "ms-settings:dateandtime",
            "AM Designator":                            "ms-settings:dateandtime",
            "PM Designator":                            "ms-settings:dateandtime",
            "First Day of Week":                        "ms-settings:dateandtime",
            "Decimal Symbol":                           "ms-settings:regionlanguage",
            "Thousand Separator":                       "ms-settings:regionlanguage",
            "List Separator":                           "ms-settings:regionlanguage",
            "Currency Symbol":                          "ms-settings:regionlanguage",
            "Currency Positive Format":                 "ms-settings:regionlanguage",
            "Currency Negative Format":                 "ms-settings:regionlanguage",
            "Number of Decimal Digits":                 "ms-settings:regionlanguage",
            "Leading Zeros":                            "ms-settings:regionlanguage",
            "Measurement System":                       "ms-settings:regionlanguage",
            "Time Zone":                                "ms-settings:dateandtime",
            "Daylight Saving Time Enabled":             "ms-settings:dateandtime",
            "Country/Region":                           "ms-settings:regionlanguage",

            # Accessibility - Narrator
            "Narrator - Auto Start":                    "ms-settings:easeofaccess-narrator",
            "Narrator - Voice":                         "ms-settings:easeofaccess-narrator",
            "Narrator - Speed":                         "ms-settings:easeofaccess-narrator",
            "Narrator - Volume":                        "ms-settings:easeofaccess-narrator",

            # Accessibility - Magnifier
            "Magnifier - Zoom Level":                   "ms-settings:easeofaccess-magnifier",
            "Magnifier - Follow Focus":                 "ms-settings:easeofaccess-magnifier",
            "Magnifier - Follow Mouse":                 "ms-settings:easeofaccess-magnifier",
            "Magnifier - Follow Caret":                 "ms-settings:easeofaccess-magnifier",
            "Magnifier - Docked Mode":                  "ms-settings:easeofaccess-magnifier",

            # Accessibility - High Contrast / Keys
            "High Contrast - Enable":                   "ms-settings:easeofaccess-highcontrast",
            "High Contrast - Theme":                    "ms-settings:easeofaccess-highcontrast",
            "Sticky Keys":                              "ms-settings:easeofaccess-keyboard",
            "Filter Keys":                              "ms-settings:easeofaccess-keyboard",
            "Toggle Keys":                              "ms-settings:easeofaccess-keyboard",
            "Mouse Keys":                               "ms-settings:easeofaccess-mouse",
            "Mouse Keys - Speed":                       "ms-settings:easeofaccess-mouse",
            "Mouse Keys - Acceleration":                "ms-settings:easeofaccess-mouse",

            # Accessibility - Closed Captions
            "Closed Captions":                          "ms-settings:easeofaccess-closedcaptioning",
            "Caption Color":                            "ms-settings:easeofaccess-closedcaptioning",
            "Caption Background Color":                 "ms-settings:easeofaccess-closedcaptioning",
            "Caption Font Size":                        "ms-settings:easeofaccess-closedcaptioning",
            "Visual Notifications for Sound":           "ms-settings:easeofaccess-otheroptions",
            "Text Cursor - Thickness":                  "ms-settings:easeofaccess-cursor",
            "Text Cursor - Blink Rate":                 "ms-settings:easeofaccess-cursor",

            # Gaming
            "Game Mode":                                "ms-settings:gaming-gamemode",
            "Game Bar - Enable":                        "ms-settings:gaming-gamebar",
            "Game Bar - Shortcut":                      "ms-settings:gaming-gamebar",
            "Game Bar - Controller Shortcut":           "ms-settings:gaming-gamebar",
            "Captures - Record in Background":          "ms-settings:gaming-gamedvr",
            "Captures - Max Recording Length":          "ms-settings:gaming-gamedvr",
            "Captures - Audio Quality":                 "ms-settings:gaming-gamedvr",
            "Captures - Video Quality":                 "ms-settings:gaming-gamedvr",
            "Captures - Framerate":                     "ms-settings:gaming-gamedvr",
            "Captures - Save Location":                 "ms-settings:gaming-gamedvr",
            "TruePlay (Anti-Cheat)":                    "ms-settings:gaming-trueplay",

            # System & Performance
            "Visual Effects - Adjust for Best Performance": "ms-settings:display-advancedgraphics",
            "Show Windows Contents While Dragging":     "ms-settings:display-advancedgraphics",
            "Smooth Edges of Screen Fonts":             "ms-settings:display-advancedgraphics",
            "Animate Controls and Elements":            "ms-settings:display-advancedgraphics",
            "Show Thumbnails Instead of Icons":         "ms-settings:display-advancedgraphics",
            "Save Taskbar Thumbnail Previews":          "ms-settings:taskbar",
            "Virtual Memory Size":                      "ms-settings:about",
            "Disable Pagefile":                         "ms-settings:about",
            "Large System Cache":                       "ms-settings:about",
            "Processor Scheduling - Programs/Background": "ms-settings:about",
            "DEP (Data Execution Prevention)":          "ms-settings:about",
        }

        # ── Category-level fallbacks ───────────────────────────────────────
        from src.models.setting import SettingCategory
        CATEGORY_FALLBACK: dict = {
            SettingCategory.APPEARANCE:     "ms-settings:personalization",
            SettingCategory.FILE_EXPLORER:  "__explorer_options__",
            SettingCategory.TASKBAR:        "ms-settings:taskbar",
            SettingCategory.POWER:          "ms-settings:powersleep",
            SettingCategory.PRIVACY:        "ms-settings:privacy",
            SettingCategory.KEYBOARD:       "ms-settings:mousetouchpad",
            SettingCategory.SYSTEM:         "ms-settings:system",
            SettingCategory.NETWORK:        "ms-settings:network",
        }

        setting_name = setting.name
        settings_uri: str | None = None

        # 1. Exact name match
        if setting_name in SETTINGS_MAP:
            settings_uri = SETTINGS_MAP[setting_name]
        else:
            # 2. Partial / substring match (case-insensitive)
            name_lower = setting_name.lower()
            for key, uri in SETTINGS_MAP.items():
                if key.lower() in name_lower or name_lower in key.lower():
                    settings_uri = uri
                    break

        # 3. Category fallback
        if not settings_uri:
            settings_uri = CATEGORY_FALLBACK.get(
                getattr(setting, "category", None), "ms-settings:"
            )

        # ── Special case: File Explorer Folder Options ─────────────────────
        if settings_uri == "__explorer_options__":
            # Open Explorer Folder Options dialog directly (no ms-settings page)
            try:
                import subprocess as _sp
                _sp.Popen(
                    ["rundll32.exe", "shell32.dll,Options_RunDLL", "0"],
                    shell=False
                )
            except Exception:
                pass
            return

        # ── Open the ms-settings URI ───────────────────────────────────────
        try:
            import subprocess as _sp
            _sp.Popen(f"start {settings_uri}", shell=True)
        except Exception:
            pass

    def _create_system_settings_link(self, parent, setting):
        """Create a link button to open system settings for this setting"""
        import tkinter as tk
        from tkinter import ttk
        
        # Create a link-style button
        link_btn = ttk.Button(
            parent,
            text="⚙️ Open in System Settings",
            command=lambda: self._launch_system_settings(setting),
            style="Link.TButton"
        )
        
        # Configure link style if not already done
        try:
            style = ttk.Style()
            style.configure("Link.TButton", 
                          foreground="blue", 
                          font=("Segoe UI", 9, "underline"),
                          relief="flat",
                          background="white")
        except:
            pass
        
        return link_btn

    def _should_show_system_settings_link(self, setting):
        """Determine if a setting should show system settings link"""
        # Show system settings link for ALL settings in Manual Configuration
        # This provides users with direct access to native Windows settings
        return True

    def _create_tooltip(self, widget, text):
        """Create a tooltip for a widget"""
        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            label = tk.Label(tooltip, text=text, background="#ffffe0", 
                           relief=tk.SOLID, borderwidth=1, font=("Segoe UI", 9))
            label.pack()
            
            widget.tooltip = tooltip

        def on_leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip

        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def _create_button_controls(self, parent, setting, current_val, options, setting_id, callback=None):
        """Create button controls for multi-option settings"""
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(control_frame, text="Select:", style="Description.TLabel").pack(side=tk.LEFT, padx=(0, 10))
        
        buttons_frame = ttk.Frame(control_frame)
        buttons_frame.pack(side=tk.LEFT)
        
        buttons = []
        for option_name, option_value in options.items():
            # Compare current value with option value for selection state
            is_selected = str(current_val) == str(option_value)
            
            option_btn = ttk.Button(buttons_frame, text=option_name,
                                style=("Accent.TButton" if is_selected else "TButton"),
                                width=12)
            option_btn.pack(side=tk.LEFT, padx=(0, 5))
            
            # Add tooltip if hint is available
            hint = self._get_option_hint(setting, option_name)
            if hint:
                self._create_tooltip(option_btn, hint)
            
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
            # Compare current value with option value for selection state
            is_selected = str(current_val) == str(option_value)
            
            option_btn = ttk.Button(control_frame, text=option_name,
                                style=("Accent.TButton" if is_selected else "TButton"),
                                width=15)
            option_btn.pack(side=tk.LEFT, padx=(0, 8))
            
            # Add tooltip if hint is available
            hint = self._get_option_hint(setting, option_name)
            if hint:
                self._create_tooltip(option_btn, hint)
            
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
            # Check if values is already a dictionary (parsed by setting_loader)
            if isinstance(setting.values, dict):
                # setting_loader parses as: registry_value -> user_friendly_label
                # But we need: user_friendly_label -> registry_value for display
                for reg_val, label in setting.values.items():
                    options[label] = reg_val
                return options
            
            # Otherwise, parse from string format
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
                value = left.strip()  # Registry value (e.g., "0", "1")
                name = right.strip()  # User-friendly label (e.g., "Left", "Center")
                if value and name:
                    options[name] = value  # Map label -> registry value
        else:
            # Default boolean options with user-friendly labels
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
        """Update the visual selection state of setting controls"""
        row = self.manual_row_widgets.get(setting_id)
        if not row:
            return

        setting = row.get("setting")
        if not setting:
            return

        # Read current registry value
        current_val = self.registry_handler.read_value(setting.hive, setting.key_path, setting.value_name)
        
        # Update current value display
        friendly_label = self._get_friendly_label_for_value(setting, current_val)
        current_value_label = row.get("current_value_label")
        if current_value_label:
            current_value_label.config(text=f"Current: {friendly_label}")
        
        # Update button selection states
        buttons = row.get("buttons", [])
        if buttons:
            for btn, opt_val in buttons:
                # Compare registry value with option value
                is_selected = str(current_val) == str(opt_val)
                btn.configure(style=("Accent.TButton" if is_selected else "TButton"))
        
        # Update slider value if present
        slider_var = row.get("slider_var")
        if slider_var and current_val is not None:
            try:
                slider_var.set(float(current_val))
            except (TypeError, ValueError):
                pass  # Keep current slider value if conversion fails

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
