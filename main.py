import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import requests
import time
from geopy.distance import geodesic
import json
import re
import threading
from queue import Queue
import os
from pathlib import Path

# Configure CustomTkinter
ctk.set_appearance_mode("dark")  # "dark" or "light"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

# Cache file location
CACHE_FILE = Path.home() / ".address_distance_cache.json"

class GlassFrame(ctk.CTkFrame):
    """Custom glassmorphic frame with semi-transparent effect"""
    def __init__(self, master, **kwargs):
        # Extract custom parameters
        glass_color = kwargs.pop('glass_color', ("#e8ecf1", "#2b3e50"))
        border_color = kwargs.pop('border_color', ("#b8c5d6", "#3a4f6f"))
        
        super().__init__(
            master,
            fg_color=glass_color,
            border_width=2,
            border_color=border_color,
            corner_radius=12,
            **kwargs
        )

class AddressDistanceCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Address Distance Calculator")
        self.root.geometry("1200x700")
        self.root.minsize(1100, 650)
        
        # Theme state
        self.current_theme = "dark"
        
        # Modern glassmorphic color scheme
        self.colors = {
            'bg_primary': '#1a1a2e',
            'bg_secondary': '#16213e',
            'glass_bg': ('#e8ecf1', '#2b3e50'),
            'glass_border': ('#b8c5d6', '#3a4f6f'),
            'accent': '#0f4c75',
            'accent_hover': '#3282b8',
            'success': '#2ecc71',
            'warning': '#f39c12',
            'error': '#e74c3c',
            'text_primary': '#ecf0f1',
            'text_secondary': '#95a5a6'
        }
        
        # Store site addresses with geocoding cache
        self.site_addresses = []
        self.geocode_cache = {}
        self.all_results = []
        
        # Load persistent cache
        self.load_cache()
        
        # Filter states
        self.filter_vars = {
            'success': ctk.BooleanVar(value=True),
            'cached': ctk.BooleanVar(value=True),
            'broad': ctk.BooleanVar(value=True),
            'not_found': ctk.BooleanVar(value=True)
        }
        
        # Threading control
        self.calculation_thread = None
        self.stop_calculation = False
        self.result_queue = Queue()
        
        # Column selection for copying
        self.selected_columns = set()
        
        # Table cell selection for Excel-like behavior
        self.selected_cells = set()  # Set of (row, col) tuples
        self.selection_start = None  # Starting cell for drag selection
        self.is_dragging = False
        self.last_clicked_cell = None  # For shift-click range selection
        
        self.create_ui()
        
        # Bind keyboard shortcuts for results table
        self.root.bind('<Control-a>', lambda e: self.select_all_results())
        self.root.bind('<Control-c>', lambda e: self.copy_results_smart())
        self.root.bind('<Escape>', lambda e: self.clear_cell_selection())
        
        # Save cache on window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Start result queue processor
        self.process_queue()
    
    def create_ui(self):
        """Create the modern glassmorphic user interface"""
        # Main container with glassmorphic effect
        main_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        main_frame.grid(row=0, column=0, sticky="nsew", padx=15, pady=15)
        
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)  # Left column (input section)
        main_frame.grid_columnconfigure(1, weight=3)  # Right column (results) - 3x more space
        
        # Title Section - compact with theme toggle
        title_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        title_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        title_frame.grid_columnconfigure(0, weight=1)
        
        # Left side - titles
        title_left = ctk.CTkFrame(title_frame, fg_color="transparent")
        title_left.grid(row=0, column=0, sticky="w")
        
        title_label = ctk.CTkLabel(
            title_left, 
            text="Address Distance Calculator",
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
            text_color="#3282b8"
        )
        title_label.pack(anchor="w")
        
        subtitle_label = ctk.CTkLabel(
            title_left,
            text="Powered by Support AV Services - Operations & Integrations",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#95a5a6"
        )
        subtitle_label.pack(anchor="w")
        
        # Right side - theme toggle
        self.theme_btn = ctk.CTkButton(
            title_frame,
            text="☀️ Light Mode",
            command=self.toggle_theme,
            width=130,
            height=34,
            corner_radius=8,
            fg_color="#34495e",
            hover_color="#2c3e50",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.theme_btn.grid(row=0, column=1, sticky="e", padx=(10, 0))
        
        # LEFT COLUMN - Input Section
        left_column = ctk.CTkFrame(main_frame, fg_color="transparent")
        left_column.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        left_column.grid_columnconfigure(0, weight=1)
        left_column.grid_rowconfigure(1, weight=1)
        
        # Technician Address Section - compact
        tech_frame = GlassFrame(left_column, glass_color=self.colors['glass_bg'])
        tech_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        tech_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            tech_frame,
            text="Technician Address",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 4))
        
        self.tech_address = ctk.CTkTextbox(
            tech_frame, 
            height=35,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=6,
            border_width=2,
            border_color=self.colors['glass_border']
        )
        self.tech_address.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))
        
        # Bind paste event to auto-format technician address
        self.tech_address.bind('<<Paste>>', self.handle_tech_address_paste)
        self.tech_address.bind('<KeyRelease>', self.auto_format_tech_address)
        
        # Site Addresses Section
        site_frame = GlassFrame(left_column, glass_color=self.colors['glass_bg'])
        site_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        site_frame.grid_columnconfigure(0, weight=1)
        site_frame.grid_rowconfigure(1, weight=1)
        
        # Header with compact legend
        header_frame = ctk.CTkFrame(site_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        header_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            header_frame,
            text="Site Addresses",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            anchor="w"
        ).grid(row=0, column=0, sticky="w")
        
        # Compact legend below title
        legend_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        legend_frame.grid(row=1, column=0, sticky="w", pady=(4, 0))
        
        ctk.CTkLabel(
            legend_frame,
            text="Legend:",
            font=ctk.CTkFont(family="Segoe UI", size=9, weight="bold"),
            text_color="#95a5a6"
        ).pack(side="left", padx=(0, 6))
        
        self.create_legend_item(legend_frame, "✓", "#2ecc71")
        self.create_legend_item(legend_frame, "💾", "#3498db")
        self.create_legend_item(legend_frame, "⚠", "#f39c12")
        self.create_legend_item(legend_frame, "✗", "#e74c3c")
        
        # Input table container with scrollable frame
        self.input_scroll = ctk.CTkScrollableFrame(
            site_frame,
            height=180,
            corner_radius=6,
            border_width=2,
            border_color=self.colors['glass_border'],
            fg_color=("#f5f7fa", "#1c2833")
        )
        self.input_scroll.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 6))
        self.input_scroll.grid_columnconfigure(2, weight=1)
        
        # Compact headers for input table with checkbox column
        headers = ["", "Status", "Address", "Suburb", "State"]
        widths = [30, 90, 280, 120, 60]
        for idx, (header, width) in enumerate(zip(headers, widths)):
            anchor = "w" if idx == 2 else "center"
            ctk.CTkLabel(
                self.input_scroll, 
                text=header,
                width=width,
                font=ctk.CTkFont(weight="bold", size=11),
                anchor=anchor
            ).grid(row=0, column=idx, padx=3, pady=3, sticky="w" if idx == 2 else "")
        
        self.input_rows = []
        
        # Compact paste area
        paste_frame = ctk.CTkFrame(site_frame, fg_color="transparent")
        paste_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 6))
        paste_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(
            paste_frame, 
            text="Paste:",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#95a5a6"
        ).grid(row=0, column=0, padx=(0, 8), sticky="w")
        
        self.paste_entry = ctk.CTkTextbox(
            paste_frame,
            height=45,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=6,
            border_width=2,
            border_color=self.colors['glass_border']
        )
        self.paste_entry.grid(row=0, column=1, sticky="ew")
        
        # Buttons - improved layout and colors (responsive)
        button_frame = ctk.CTkFrame(site_frame, fg_color="transparent")
        button_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10), padx=12)
        button_frame.grid_columnconfigure(0, weight=1)
        
        # Single row - all buttons together
        btn_row_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
        btn_row_frame.grid(row=0, column=0, sticky="w")
        
        self.add_btn = ctk.CTkButton(
            btn_row_frame,
            text="➕ Add",
            command=self.process_pasted_data,
            width=100,
            height=34,
            corner_radius=8,
            fg_color="#27ae60",
            hover_color="#229954",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.add_btn.pack(side="left", padx=2)
        
        self.remove_btn = ctk.CTkButton(
            btn_row_frame,
            text="✖ Remove",
            command=self.remove_selected,
            width=110,
            height=34,
            corner_radius=8,
            fg_color="#e74c3c",
            hover_color="#c0392b",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.remove_btn.pack(side="left", padx=2)
        
        self.clear_btn = ctk.CTkButton(
            btn_row_frame,
            text="🗑 Clear All",
            command=self.clear_sites,
            width=110,
            height=34,
            corner_radius=8,
            fg_color="#e67e22",
            hover_color="#d35400",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.clear_btn.pack(side="left", padx=2)
        
        self.calc_btn = ctk.CTkButton(
            btn_row_frame,
            text="🚀 Calculate Distances",
            command=self.calculate_distances,
            width=180,
            height=34,
            corner_radius=8,
            fg_color="#3282b8",
            hover_color="#0f4c75",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.calc_btn.pack(side="left", padx=2)
        
        # RIGHT COLUMN - Results Section
        right_column = ctk.CTkFrame(main_frame, fg_color="transparent")
        right_column.grid(row=1, column=1, sticky="nsew", padx=(8, 0))
        right_column.grid_columnconfigure(0, weight=1)
        right_column.grid_rowconfigure(0, weight=1)
        
        # Results Section
        results_frame = GlassFrame(right_column, glass_color=self.colors['glass_bg'])
        results_frame.grid(row=0, column=0, sticky="nsew")
        results_frame.grid_columnconfigure(0, weight=1)
        results_frame.grid_rowconfigure(2, weight=1)
        
        # Header with Copy button
        results_header_frame = ctk.CTkFrame(results_frame, fg_color="transparent")
        results_header_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        results_header_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            results_header_frame,
            text="Results",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            anchor="w"
        ).grid(row=0, column=0, sticky="w")
        
        # Selection info label
        self.selection_info_label = ctk.CTkLabel(
            results_header_frame,
            text="",
            font=ctk.CTkFont(family="Segoe UI", size=9),
            text_color="#7f8c8d",
            anchor="w"
        )
        self.selection_info_label.grid(row=1, column=0, sticky="w", pady=(2, 0))
        
        self.copy_btn = ctk.CTkButton(
            results_header_frame,
            text="📋 Copy Results",
            command=self.copy_results_smart,
            width=140,
            height=32,
            corner_radius=8,
            fg_color="#16a085",
            hover_color="#138d75",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold")
        )
        self.copy_btn.grid(row=0, column=1, sticky="e", padx=(10, 0))
        
        # Compact filter controls
        filter_frame = ctk.CTkFrame(results_frame, fg_color="transparent")
        filter_frame.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 6))
        
        ctk.CTkLabel(
            filter_frame,
            text="Show:",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color="#95a5a6"
        ).pack(side="left", padx=(0, 8))
        
        ctk.CTkCheckBox(
            filter_frame,
            text="Found",
            variable=self.filter_vars['success'],
            command=self.apply_filters,
            fg_color="#27ae60",
            hover_color="#229954",
            corner_radius=5,
            font=ctk.CTkFont(size=11),
            checkbox_width=18,
            checkbox_height=18
        ).pack(side="left", padx=3)
        
        ctk.CTkCheckBox(
            filter_frame,
            text="Cached",
            variable=self.filter_vars['cached'],
            command=self.apply_filters,
            fg_color="#3498db",
            hover_color="#2980b9",
            corner_radius=5,
            font=ctk.CTkFont(size=11),
            checkbox_width=18,
            checkbox_height=18
        ).pack(side="left", padx=3)
        
        ctk.CTkCheckBox(
            filter_frame,
            text="Broad",
            variable=self.filter_vars['broad'],
            command=self.apply_filters,
            fg_color="#f39c12",
            hover_color="#e67e22",
            corner_radius=5,
            font=ctk.CTkFont(size=11),
            checkbox_width=18,
            checkbox_height=18
        ).pack(side="left", padx=3)
        
        ctk.CTkCheckBox(
            filter_frame,
            text="Not Found",
            variable=self.filter_vars['not_found'],
            command=self.apply_filters,
            fg_color="#e74c3c",
            hover_color="#c0392b",
            corner_radius=5,
            font=ctk.CTkFont(size=11),
            checkbox_width=18,
            checkbox_height=18
        ).pack(side="left", padx=3)
        
        # Results table
        self.results_scroll = ctk.CTkScrollableFrame(
            results_frame,
            corner_radius=6,
            border_width=2,
            border_color=self.colors['glass_border'],
            fg_color=("#f5f7fa", "#1c2833")
        )
        self.results_scroll.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 10))
        self.results_scroll.grid_columnconfigure(1, weight=1)
        
        # Compact headers for results table with click selection
        result_headers = ["Rank", "Address", "Suburb", "State", "Dist(km)", "Distance(min)", "Status"]
        result_widths = [45, 280, 110, 60, 75, 100, 140]
        self.result_header_labels = []
        
        for idx, (header, width) in enumerate(zip(result_headers, result_widths)):
            anchor = "w"  # Align all headers to the left
            header_label = ctk.CTkLabel(
                self.results_scroll, 
                text=header,
                width=width,
                font=ctk.CTkFont(weight="bold", size=11),
                anchor=anchor,
                cursor="hand2"
            )
            header_label.grid(row=0, column=idx, padx=2, pady=3, sticky="w")
            header_label.bind("<Button-1>", lambda e, col=idx: self.select_column(col))
            self.result_header_labels.append(header_label)
        
        self.results_rows = []
        
        # Status bar - full width at bottom
        status_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        status_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        status_frame.grid_columnconfigure(0, weight=1)
        
        self.status_var = tk.StringVar(value="✓ Ready to calculate distances")
        
        self.status_label = ctk.CTkLabel(
            status_frame,
            textvariable=self.status_var,
            anchor="w",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=("#d5dce4", "#34495e"),
            corner_radius=6,
            height=32
        )
        self.status_label.grid(row=0, column=0, sticky="ew", pady=(0, 6), padx=2)
        
        self.progress = ctk.CTkProgressBar(
            status_frame,
            corner_radius=6,
            height=10,
            progress_color="#27ae60",
            fg_color=("#d5dce4", "#34495e")
        )
        self.progress.grid(row=1, column=0, sticky="ew", padx=2)
        self.progress.set(0)
        self.progress.grid_remove()
        
        # Configure weights for responsive layout
        main_frame.grid_rowconfigure(1, weight=1)
    
    def toggle_theme(self):
        """Toggle between light and dark themes"""
        if self.current_theme == "dark":
            # Switch to light mode
            self.current_theme = "light"
            ctk.set_appearance_mode("light")
            self.theme_btn.configure(
                text="🌙 Dark Mode",
                fg_color="#2c3e50",
                hover_color="#34495e"
            )
            # Update color scheme for light mode
            self.update_theme_colors("light")
        else:
            # Switch to dark mode
            self.current_theme = "dark"
            ctk.set_appearance_mode("dark")
            self.theme_btn.configure(
                text="☀️ Light Mode",
                fg_color="#34495e",
                hover_color="#2c3e50"
            )
            # Update color scheme for dark mode
            self.update_theme_colors("dark")
    
    def update_theme_colors(self, theme):
        """Update UI elements based on theme"""
        # This method can be extended to update specific elements if needed
        pass
    
    def load_cache(self):
        """Load geocoding cache from persistent storage"""
        try:
            if CACHE_FILE.exists():
                with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                    self.geocode_cache = json.load(f)
                print(f"✓ Loaded {len(self.geocode_cache)} cached addresses from {CACHE_FILE}")
            else:
                print("ℹ No cache file found, starting with empty cache")
        except Exception as e:
            print(f"⚠ Error loading cache: {e}")
            self.geocode_cache = {}
    
    def save_cache(self):
        """Save geocoding cache to persistent storage"""
        try:
            with open(CACHE_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.geocode_cache, f, indent=2, ensure_ascii=False)
            print(f"✓ Saved {len(self.geocode_cache)} addresses to cache")
        except Exception as e:
            print(f"⚠ Error saving cache: {e}")
    
    def on_closing(self):
        """Handle application closing - save cache and cleanup"""
        self.save_cache()
        self.root.destroy()
    
    def create_legend_item(self, parent, text, color):
        """Create a compact squared legend item"""
        frame = ctk.CTkFrame(parent, fg_color=color, corner_radius=3, height=20, width=20)
        frame.pack(side="left", padx=2)
        frame.pack_propagate(False)
        
        label = ctk.CTkLabel(
            frame,
            text=text,
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color="white"
        )
        label.pack(expand=True)
    
    def add_input_row(self, status, address, suburb, state, tag='pending'):
        """Add a row to the input table"""
        row_num = len(self.input_rows) + 1
        
        # Color coding based on tag - improved for light mode
        colors = {
            'pending': ('#7f8c8d', '#95a5a6'),  # darker for light mode, lighter for dark mode
            'cached': '#3498db',
            'success': '#27ae60',
            'warning': '#e67e22',
            'error': '#c0392b'
        }
        
        fg_color = colors.get(tag, ('#7f8c8d', '#95a5a6'))
        
        # Create row frame with hover effect
        row_frame = ctk.CTkFrame(self.input_scroll, fg_color="transparent")
        row_frame.grid(row=row_num, column=0, columnspan=5, sticky="ew", pady=2)
        row_frame.grid_columnconfigure(2, weight=1)
        
        # Checkbox for selection
        var = ctk.BooleanVar(value=False)
        checkbox = ctk.CTkCheckBox(
            row_frame,
            text="",
            variable=var,
            width=30,
            checkbox_width=18,
            checkbox_height=18,
            corner_radius=4,
            fg_color="#e74c3c",
            hover_color="#c0392b"
        )
        checkbox.grid(row=0, column=0, padx=(3, 0))
        
        # Function to toggle checkbox when clicking on the row
        def toggle_checkbox(event=None):
            var.set(not var.get())
        
        # Create status label and keep reference for updates
        status_label = ctk.CTkLabel(row_frame, text=status, width=90, font=ctk.CTkFont(size=10), text_color=fg_color)
        status_label.grid(row=0, column=1, padx=3)
        
        address_label = ctk.CTkLabel(row_frame, text=address, width=280, font=ctk.CTkFont(size=10), anchor="w")
        address_label.grid(row=0, column=2, padx=3, sticky="w")
        
        suburb_label = ctk.CTkLabel(row_frame, text=suburb, width=120, font=ctk.CTkFont(size=10))
        suburb_label.grid(row=0, column=3, padx=3)
        
        state_label = ctk.CTkLabel(row_frame, text=state, width=60, font=ctk.CTkFont(size=10))
        state_label.grid(row=0, column=4, padx=3)
        
        # Bind click events to row frame and all labels (but not checkbox)
        row_frame.bind("<Button-1>", toggle_checkbox)
        status_label.bind("<Button-1>", toggle_checkbox)
        address_label.bind("<Button-1>", toggle_checkbox)
        suburb_label.bind("<Button-1>", toggle_checkbox)
        state_label.bind("<Button-1>", toggle_checkbox)
        
        self.input_rows.append({
            'frame': row_frame, 
            'address': address, 
            'suburb': suburb, 
            'state': state,
            'checkbox_var': var,
            'checkbox': checkbox,
            'status_label': status_label  # Store reference to status label
        })
    
    def update_input_row_status(self, row_index, status, tag):
        """Update the status of an input row"""
        if row_index >= len(self.input_rows):
            return
        
        # Color coding based on tag
        colors = {
            'pending': ('#7f8c8d', '#95a5a6'),
            'cached': '#3498db',
            'success': '#27ae60',
            'warning': '#e67e22',
            'error': '#c0392b'
        }
        
        fg_color = colors.get(tag, ('#7f8c8d', '#95a5a6'))
        
        # Update the status label
        status_label = self.input_rows[row_index]['status_label']
        status_label.configure(text=status, text_color=fg_color)
    
    def format_duration(self, minutes):
        """Format duration from minutes to 'X hr Y min' or 'Y min' format"""
        if isinstance(minutes, str):
            return minutes  # Return as-is if it's already a string (like 'N/A')
        
        total_minutes = int(round(minutes))
        hours = total_minutes // 60
        mins = total_minutes % 60
        
        if hours > 0:
            if mins > 0:
                return f"{hours} hr {mins} min"
            else:
                return f"{hours} hr"
        else:
            return f"{mins} min"
    
    def add_result_row(self, rank, address, suburb, state, distance, duration, status, tag='success'):
        """Add a row to the results table with selectable cells"""
        row_num = len(self.results_rows) + 1
        
        # Color coding based on tag
        colors = {
            'success': '#27ae60',
            'cached': '#3498db',
            'warning': '#e67e22',
            'error': '#c0392b'
        }
        
        fg_color = colors.get(tag, ('#7f8c8d', '#95a5a6'))
        
        # Create row frame
        row_frame = ctk.CTkFrame(self.results_scroll, fg_color="transparent")
        row_frame.grid(row=row_num, column=0, columnspan=7, sticky="ew", pady=2)
        row_frame.grid_columnconfigure(1, weight=1)
        
        # Cell data and configuration - all aligned to the left
        cell_data = [
            (str(rank), 45, ctk.CTkFont(size=10, weight="bold"), ("#2c3e50", "#ecf0f1"), "w"),
            (address, 280, ctk.CTkFont(size=10), None, "w"),
            (suburb, 110, ctk.CTkFont(size=10), None, "w"),
            (state, 60, ctk.CTkFont(size=10), None, "w"),
            (distance, 75, ctk.CTkFont(size=10), "#3498db", "w"),
            (duration, 100, ctk.CTkFont(size=10), "#3498db", "w"),
            (status, 140, ctk.CTkFont(size=10, weight="bold"), fg_color, "w")
        ]
        
        cells = []
        for col_idx, (text, width, font, text_color, anchor) in enumerate(cell_data):
            cell_label = ctk.CTkLabel(
                row_frame, 
                text=text, 
                width=width, 
                font=font, 
                anchor=anchor,
                cursor="hand2",
                fg_color="transparent"
            )
            if text_color:
                cell_label.configure(text_color=text_color)
            
            cell_label.grid(row=0, column=col_idx, padx=2, sticky="w")
            
            # Bind mouse events for selection
            cell_label.bind("<Button-1>", lambda e, r=row_num, c=col_idx: self.on_cell_click(e, r, c))
            cell_label.bind("<Shift-Button-1>", lambda e, r=row_num, c=col_idx: self.on_cell_shift_click(e, r, c))
            cell_label.bind("<Control-Button-1>", lambda e, r=row_num, c=col_idx: self.on_cell_ctrl_click(e, r, c))
            cell_label.bind("<B1-Motion>", lambda e, r=row_num, c=col_idx: self.on_cell_drag(e, r, c))
            cell_label.bind("<ButtonRelease-1>", lambda e: self.on_cell_release(e))
            cell_label.bind("<Button-3>", lambda e, r=row_num, c=col_idx: self.show_cell_context_menu(e, r, c))
            
            cells.append(cell_label)
        
        self.results_rows.append({'frame': row_frame, 'cells': cells})
    
    def clear_input_rows(self):
        """Clear all input rows"""
        for row in self.input_rows:
            row['frame'].destroy()
        self.input_rows = []
    
    def clear_results_rows(self):
        """Clear all result rows"""
        for row in self.results_rows:
            row['frame'].destroy()
        self.results_rows = []
        self.selected_cells.clear()
        self.selection_start = None
        self.last_clicked_cell = None
    
    def on_cell_click(self, event, row, col):
        """Handle single cell click - clear previous selection and select this cell"""
        self.clear_cell_selection()
        self.selected_cells.add((row, col))
        self.last_clicked_cell = (row, col)
        self.selection_start = (row, col)
        self.is_dragging = False
        self.highlight_selected_cells()
        
    def on_cell_shift_click(self, event, row, col):
        """Handle shift-click - select range from last clicked cell to this cell"""
        if self.last_clicked_cell:
            self.clear_cell_selection()
            start_row, start_col = self.last_clicked_cell
            end_row, end_col = row, col
            
            # Select rectangular range
            for r in range(min(start_row, end_row), max(start_row, end_row) + 1):
                for c in range(min(start_col, end_col), max(start_col, end_col) + 1):
                    self.selected_cells.add((r, c))
            
            self.highlight_selected_cells()
        else:
            self.on_cell_click(event, row, col)
    
    def on_cell_ctrl_click(self, event, row, col):
        """Handle ctrl-click - toggle cell selection"""
        if (row, col) in self.selected_cells:
            self.selected_cells.remove((row, col))
        else:
            self.selected_cells.add((row, col))
        
        self.last_clicked_cell = (row, col)
        self.highlight_selected_cells()
    
    def on_cell_drag(self, event, row, col):
        """Handle drag selection - select range from start to current cell"""
        if self.selection_start:
            self.is_dragging = True
            self.clear_cell_selection()
            
            start_row, start_col = self.selection_start
            
            # Select rectangular range from start to current position
            for r in range(min(start_row, row), max(start_row, row) + 1):
                for c in range(min(start_col, col), max(start_col, col) + 1):
                    self.selected_cells.add((r, c))
            
            self.highlight_selected_cells()
    
    def on_cell_release(self, event):
        """Handle mouse release - finish drag selection"""
        self.is_dragging = False
    
    def select_column(self, col):
        """Select entire column when header is clicked"""
        self.clear_cell_selection()
        
        # Select all cells in this column
        for row_idx in range(1, len(self.results_rows) + 1):
            self.selected_cells.add((row_idx, col))
        
        self.last_clicked_cell = (1, col)
        self.highlight_selected_cells()
    
    def clear_cell_selection(self):
        """Clear all cell selections and remove highlights"""
        for row_idx, row in enumerate(self.results_rows, start=1):
            for col_idx, cell in enumerate(row['cells']):
                cell.configure(fg_color="transparent")
        self.selected_cells.clear()
        
        # Update selection info label
        if hasattr(self, 'selection_info_label'):
            self.selection_info_label.configure(text="")
    
    def highlight_selected_cells(self):
        """Highlight all selected cells"""
        # First clear all highlights
        for row_idx, row in enumerate(self.results_rows, start=1):
            for col_idx, cell in enumerate(row['cells']):
                cell.configure(fg_color="transparent")
        
        # Then highlight selected cells
        selection_color = ("#b3d9ff", "#2d5f8f")  # Light blue for light mode, darker blue for dark mode
        
        for (row, col) in self.selected_cells:
            if 1 <= row <= len(self.results_rows):
                row_data = self.results_rows[row - 1]
                if 0 <= col < len(row_data['cells']):
                    row_data['cells'][col].configure(fg_color=selection_color)
        
        # Update selection info label
        if self.selected_cells:
            cell_count = len(self.selected_cells)
            self.selection_info_label.configure(text=f"📌 {cell_count} cell(s) selected")
        else:
            self.selection_info_label.configure(text="")
    
    def copy_selected_cells(self):
        """Copy selected cells to clipboard"""
        if not self.selected_cells:
            messagebox.showinfo("No Selection", "Please select cells to copy")
            return
        
        # Sort selected cells by row then column
        sorted_cells = sorted(list(self.selected_cells))
        
        if not sorted_cells:
            return
        
        # Group cells by row
        rows_dict = {}
        for row, col in sorted_cells:
            if row not in rows_dict:
                rows_dict[row] = []
            rows_dict[row].append(col)
        
        # Build clipboard text
        lines = []
        for row in sorted(rows_dict.keys()):
            cols = sorted(rows_dict[row])
            row_data = self.results_rows[row - 1]
            
            # Get cell text values
            cell_values = []
            for col in cols:
                if 0 <= col < len(row_data['cells']):
                    cell_values.append(row_data['cells'][col].cget("text"))
            
            lines.append('\t'.join(cell_values))
        
        text = '\n'.join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        
        cell_count = len(self.selected_cells)
        messagebox.showinfo("Copied", f"Copied {cell_count} cell(s) to clipboard!")
        self.status_var.set(f"✓ Copied {cell_count} cell(s) to clipboard")
    
    def select_all_results(self):
        """Select all cells in the results table"""
        self.clear_cell_selection()
        
        for row_idx in range(1, len(self.results_rows) + 1):
            for col_idx in range(7):  # 7 columns
                self.selected_cells.add((row_idx, col_idx))
        
        self.highlight_selected_cells()
    
    def show_cell_context_menu(self, event, row, col):
        """Show context menu on right-click"""
        # If the right-clicked cell is not in selection, select it
        if (row, col) not in self.selected_cells:
            self.on_cell_click(event, row, col)
        
        # Create context menu
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="Copy Selected Cells", command=self.copy_selected_cells)
        context_menu.add_command(label="Copy All Results", command=self.copy_all_results)
        context_menu.add_separator()
        context_menu.add_command(label="Select All (Ctrl+A)", command=self.select_all_results)
        context_menu.add_command(label="Clear Selection (Esc)", command=self.clear_cell_selection)
        
        # Display menu at cursor position
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    
    def handle_paste(self, event=None):
        """Handle paste event"""
        self.root.after(100, self.auto_process_hint)
        return None
    
    def handle_tech_address_paste(self, event=None):
        """Handle paste event for technician address with auto-formatting"""
        self.root.after(100, self.auto_format_tech_address)
        return None
    
    def auto_format_tech_address(self, event=None):
        """Auto-format the technician address by cleaning up whitespace, newlines, and tabs"""
        current_text = self.tech_address.get("1.0", tk.END).strip()
        
        if not current_text:
            return
        
        # Replace tabs with commas (for Excel column data)
        formatted_text = current_text.replace('\t', ', ')
        
        # Replace newlines with commas
        formatted_text = formatted_text.replace('\n', ', ')
        formatted_text = formatted_text.replace('\r', '')
        
        # Remove double commas
        while ', , ' in formatted_text:
            formatted_text = formatted_text.replace(', , ', ', ')
        
        # Remove multiple consecutive whitespaces
        formatted_text = re.sub(r'\s+', ' ', formatted_text)
        
        # Remove spaces before commas
        formatted_text = re.sub(r'\s+,', ',', formatted_text)
        
        # Ensure single space after commas
        formatted_text = re.sub(r',\s*', ', ', formatted_text)
        
        # Remove leading/trailing commas and whitespace
        formatted_text = formatted_text.strip(', ').strip()
        
        # Only update if the text changed to avoid infinite loops
        if formatted_text != current_text:
            self.tech_address.delete("1.0", tk.END)
            self.tech_address.insert("1.0", formatted_text)
            self.status_var.set("✓ Technician address auto-formatted")
    
    def auto_process_hint(self):
        """Show hint after pasting"""
        pasted_text = self.paste_entry.get("1.0", tk.END).strip()
        if pasted_text:
            self.status_var.set("📋 Data pasted! Click 'Add' to process")
    
    def parse_address_line(self, line):
        """Parse a single line of address data"""
        line = line.strip()
        if not line:
            return None
        
        # Try tab-separated first (Excel default)
        if '\t' in line:
            parts = [p.strip() for p in line.split('\t') if p.strip()]
            if len(parts) >= 3:
                return {'address': parts[0], 'suburb': parts[1], 'state': parts[2]}
        
        # Try comma-separated
        if ',' in line:
            parts = [p.strip() for p in line.split(',') if p.strip()]
            if len(parts) >= 3:
                return {
                    'address': ', '.join(parts[:-2]),
                    'suburb': parts[-2],
                    'state': parts[-1]
                }
            elif len(parts) == 2:
                return {'address': '', 'suburb': parts[0], 'state': parts[1]}
        
        # Try pipe-separated
        if '|' in line:
            parts = [p.strip() for p in line.split('|') if p.strip()]
            if len(parts) >= 3:
                return {'address': parts[0], 'suburb': parts[1], 'state': parts[2]}
        
        # Single line - try to detect full address
        parts = [p.strip() for p in line.split(',')]
        if len(parts) >= 2:
            last_part = parts[-1].strip()
            state_match = re.match(r'^([A-Z]{2,3})(\s+\d{4})?$', last_part)
            if state_match:
                state = state_match.group(1)
                suburb = parts[-2] if len(parts) >= 2 else ''
                address = ', '.join(parts[:-2]) if len(parts) > 2 else ''
                return {'address': address, 'suburb': suburb, 'state': state}
        
        return None
    
    def process_pasted_data(self):
        """Process pasted data"""
        pasted_text = self.paste_entry.get("1.0", tk.END).strip()
        
        if not pasted_text:
            messagebox.showinfo("No Data", "Please paste address data first")
            return
        
        lines = pasted_text.split('\n')
        added_count = 0
        skipped_count = 0
        
        for line in lines:
            parsed = self.parse_address_line(line)
            
            if parsed and (parsed['address'] or parsed['suburb']) and parsed['state']:
                address = parsed['address']
                suburb = parsed['suburb']
                state = parsed['state']
                
                is_duplicate = any(
                    addr['address'] == address and 
                    addr['suburb'] == suburb and 
                    addr['state'] == state
                    for addr in self.site_addresses
                )
                
                if not is_duplicate:
                    cache_key = f"{address}, {suburb}, {state}".lower()
                    is_cached = cache_key in self.geocode_cache
                    
                    status = '💾 Cached' if is_cached else 'Pending'
                    tag = 'cached' if is_cached else 'pending'
                    
                    self.add_input_row(status, address, suburb, state, tag)
                    
                    site_data = {
                        'address': address,
                        'suburb': suburb,
                        'state': state,
                        'status': 'pending'
                    }
                    
                    if is_cached:
                        cached_data = self.geocode_cache[cache_key]
                        site_data.update(cached_data)
                        site_data['status'] = 'cached'
                    
                    self.site_addresses.append(site_data)
                    added_count += 1
                else:
                    skipped_count += 1
        
        self.paste_entry.delete("1.0", tk.END)
        
        if added_count > 0:
            status_msg = f"✓ Added {added_count} address(es)"
            if skipped_count > 0:
                status_msg += f" ({skipped_count} duplicate(s) skipped)"
            status_msg += f". Total: {len(self.site_addresses)}"
            self.status_var.set(status_msg)
        else:
            if skipped_count > 0:
                messagebox.showinfo("Duplicates", 
                                  f"All {skipped_count} address(es) already exist in the list.")
                self.status_var.set(f"ℹ {skipped_count} duplicate(s) skipped")
            else:
                messagebox.showwarning("Invalid Format", 
                                     "No valid addresses found.\n\n" +
                                     "Supported formats:\n" +
                                     "• Excel columns (tab-separated): Address | Suburb | State\n" +
                                     "• Comma-separated: Address, Suburb, State\n" +
                                     "• Full address: Street, Suburb, STATE 1234")
                self.status_var.set("⚠ No valid addresses found")
    
    def remove_selected(self):
        """Remove selected addresses"""
        # Find all selected rows
        selected_indices = []
        for idx, row_data in enumerate(self.input_rows):
            if row_data['checkbox_var'].get():
                selected_indices.append(idx)
        
        if not selected_indices:
            messagebox.showinfo("No Selection", "Please select at least one address to remove by checking the checkbox(es).")
            return
        
        # Confirm removal
        count = len(selected_indices)
        if not messagebox.askyesno("Confirm Removal", 
                                   f"Remove {count} selected address(es)?"):
            return
        
        # Remove in reverse order to maintain correct indices
        for idx in reversed(selected_indices):
            # Destroy the frame
            self.input_rows[idx]['frame'].destroy()
            # Remove from lists
            del self.input_rows[idx]
            del self.site_addresses[idx]
        
        self.status_var.set(f"✓ Removed {count} address(es). Total: {len(self.site_addresses)}")
    
    def clear_sites(self):
        """Clear all site addresses"""
        if self.site_addresses:
            if messagebox.askyesno("Confirm Clear", 
                                  f"Clear all {len(self.site_addresses)} addresses?"):
                self.site_addresses = []
                self.clear_input_rows()
                self.status_var.set("✓ All addresses cleared")
        else:
            self.status_var.set("ℹ No addresses to clear")
    
    def geocode_address_incremental(self, address, max_retries=4):
        """Geocode address with incremental broader search strategy"""
        base_url = "https://nominatim.openstreetmap.org/search"
        headers = {'User-Agent': 'AddressDistanceCalculator/3.0'}
        
        parts = [p.strip() for p in address.split(',')]
        search_attempts = []
        
        # Attempt 1: Full address
        search_attempts.append({
            'query': address,
            'level': 0,
            'description': 'exact address'
        })
        
        # Attempt 2: Remove shop/unit numbers
        if len(parts) >= 3:
            street_part = parts[0]
            cleaned_street = re.sub(r'^(Shop|Unit|Suite|Level|T/a|Tenancy|Lot)\s*\d+[A-Za-z]?,?\s*', '', 
                                   street_part, flags=re.IGNORECASE)
            
            if cleaned_street != street_part and cleaned_street.strip():
                search_attempts.append({
                    'query': ', '.join([cleaned_street] + parts[1:]),
                    'level': 1,
                    'description': 'without shop/unit'
                })
        
        # Attempt 3: Suburb + State only
        if len(parts) >= 2:
            search_attempts.append({
                'query': f"{parts[-2]}, {parts[-1]}",
                'level': 3,
                'description': 'suburb and state'
            })
        
        # Try each search attempt
        for attempt_num, attempt in enumerate(search_attempts[:max_retries], 1):
            if attempt_num > 1:
                time.sleep(0.5)
            
            params = {
                'q': attempt['query'],
                'format': 'json',
                'limit': 3,
                'countrycodes': 'au',
                'addressdetails': 1
            }
            
            try:
                response = requests.get(base_url, params=params, headers=headers, timeout=10)
                response.raise_for_status()
                data = response.json()
                
                if data:
                    return (float(data[0]['lat']), 
                           float(data[0]['lon']), 
                           attempt['level'],
                           attempt['description'])
            except Exception as e:
                print(f"Attempt {attempt_num} failed: {e}")
                continue
        
        return None, None, None, None
    
    def get_osrm_route(self, coord1, coord2):
        """Get actual route distance and duration using OSRM"""
        try:
            lon1, lat1 = coord1[1], coord1[0]
            lon2, lat2 = coord2[1], coord2[0]
            
            url = f"http://router.project-osrm.org/route/v1/driving/{lon1},{lat1};{lon2},{lat2}"
            params = {'overview': 'false', 'steps': 'false'}
            
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            data = response.json()
            
            if data['code'] == 'Ok' and data['routes']:
                route = data['routes'][0]
                distance_km = route['distance'] / 1000
                duration_min = route['duration'] / 60
                return distance_km, duration_min
            else:
                distance_km = geodesic(coord1, coord2).kilometers
                duration_min = (distance_km / 50) * 60
                return distance_km, duration_min
        except Exception as e:
            print(f"OSRM error: {e}")
            distance_km = geodesic(coord1, coord2).kilometers
            duration_min = (distance_km / 50) * 60
            return distance_km, duration_min
    
    def calculate_distances_worker(self, tech_addr):
        """Worker function for calculating distances in a separate thread"""
        try:
            self.result_queue.put(('status', "Geocoding technician address..."))
            tech_lat, tech_lon, tech_level, tech_desc = self.geocode_address_incremental(tech_addr)
            
            if not tech_lat:
                self.result_queue.put(('error', "Could not geocode technician address"))
                return
            
            tech_coords = (tech_lat, tech_lon)
            self.result_queue.put(('progress', 0.05))
            time.sleep(0.5)
            
            results = []
            total = len(self.site_addresses)
            
            for i, site in enumerate(self.site_addresses):
                if self.stop_calculation:
                    self.result_queue.put(('status', "❌ Calculation cancelled"))
                    return
                
                progress_value = 0.05 + (i + 1) / total * 0.95
                full_address = f"{site['address']}, {site['suburb']}, {site['state']}"
                cache_key = full_address.lower()
                
                if 'lat' in site and 'lon' in site and site['lat'] and site['lon']:
                    site_lat = site['lat']
                    site_lon = site['lon']
                    match_level = site.get('match_level', 0)
                    match_desc = site.get('match_desc', 'exact')
                    
                    self.result_queue.put(('status', f"💾 Using cached data: {site['suburb']}"))
                    
                    site_coords = (site_lat, site_lon)
                    distance_km, duration_min = self.get_osrm_route(tech_coords, site_coords)
                    
                    status = "💾 Cached"
                    tag = 'cached'
                else:
                    self.result_queue.put(('status', f"⏳ Geocoding: {site['suburb']}"))
                    
                    site_lat, site_lon, match_level, match_desc = self.geocode_address_incremental(full_address)
                    
                    if site_lat and site_lon:
                        site_coords = (site_lat, site_lon)
                        distance_km, duration_min = self.get_osrm_route(tech_coords, site_coords)
                        
                        self.geocode_cache[cache_key] = {
                            'lat': site_lat,
                            'lon': site_lon,
                            'match_level': match_level,
                            'match_desc': match_desc
                        }
                        
                        site['lat'] = site_lat
                        site['lon'] = site_lon
                        site['match_level'] = match_level
                        site['match_desc'] = match_desc
                        
                        if match_level == 0:
                            status = "✓ Found (exact)"
                            tag = 'success'
                        elif match_level <= 2:
                            status = f"✓ Found ({match_desc})"
                            tag = 'success'
                        else:
                            status = f"⚠ Broad ({match_desc})"
                            tag = 'warning'
                    else:
                        distance_km = float('inf')
                        duration_min = float('inf')
                        status = '✗ Not Found'
                        tag = 'error'
                        match_level = 999
                
                result = {
                    'address': site['address'],
                    'suburb': site['suburb'],
                    'state': site['state'],
                    'distance': distance_km,
                    'duration': duration_min,
                    'status': status,
                    'tag': tag,
                    'match_level': match_level
                }
                results.append(result)
                
                # Update the input row status in the UI
                self.result_queue.put(('update_row', (i, status, tag)))
                
                self.result_queue.put(('progress', progress_value))
                time.sleep(0.5)
            
            results.sort(key=lambda x: x['distance'])
            self.result_queue.put(('results', results))
            self.result_queue.put(('complete', None))
            
        except Exception as e:
            self.result_queue.put(('error', f"Calculation error: {str(e)}"))
    
    def process_queue(self):
        """Process messages from the worker thread"""
        try:
            while not self.result_queue.empty():
                msg_type, data = self.result_queue.get_nowait()
                
                if msg_type == 'status':
                    self.status_var.set(data)
                elif msg_type == 'progress':
                    self.progress.set(data)
                elif msg_type == 'results':
                    self.all_results = data
                    self.apply_filters()
                elif msg_type == 'complete':
                    self.calculation_complete()
                elif msg_type == 'error':
                    self.handle_calculation_error(data)
                elif msg_type == 'update_row':
                    # Update the status of an input row
                    row_index, status, tag = data
                    self.update_input_row_status(row_index, status, tag)
        except:
            pass
        
        self.root.after(100, self.process_queue)
    
    def calculation_complete(self):
        """Handle calculation completion"""
        self.progress.grid_remove()
        self.calc_btn.configure(state="normal")
        
        # Save cache after calculation completes
        self.save_cache()
        
        counts = {
            'success': sum(1 for r in self.all_results if r['tag'] == 'success'),
            'cached': sum(1 for r in self.all_results if r['tag'] == 'cached'),
            'warning': sum(1 for r in self.all_results if r['tag'] == 'warning'),
            'error': sum(1 for r in self.all_results if r['tag'] == 'error')
        }
        
        summary_parts = []
        if counts['success'] > 0:
            summary_parts.append(f"{counts['success']} found")
        if counts['cached'] > 0:
            summary_parts.append(f"{counts['cached']} cached")
        if counts['warning'] > 0:
            summary_parts.append(f"{counts['warning']} broad")
        if counts['error'] > 0:
            summary_parts.append(f"{counts['error']} not found")
        
        summary = f"✓ Complete! " + ", ".join(summary_parts)
        self.status_var.set(summary)
    
    def handle_calculation_error(self, error_msg):
        """Handle calculation errors"""
        self.progress.grid_remove()
        self.calc_btn.configure(state="normal")
        messagebox.showerror("Calculation Error", error_msg)
        self.status_var.set(f"✗ Error: {error_msg}")
    
    def apply_filters(self):
        """Apply filters to results display"""
        if not self.all_results:
            return
        
        self.clear_results_rows()
        
        filtered_results = []
        for result in self.all_results:
            tag = result['tag']
            
            if tag == 'success' and self.filter_vars['success'].get():
                filtered_results.append(result)
            elif tag == 'cached' and self.filter_vars['cached'].get():
                filtered_results.append(result)
            elif tag == 'warning' and self.filter_vars['broad'].get():
                filtered_results.append(result)
            elif tag == 'error' and self.filter_vars['not_found'].get():
                filtered_results.append(result)
        
        for rank, result in enumerate(filtered_results, 1):
            if result['distance'] == float('inf'):
                self.add_result_row(
                    rank, result['address'], result['suburb'], result['state'],
                    'N/A', 'N/A', result['status'], result['tag']
                )
            else:
                # Format duration using the new format_duration method
                formatted_duration = self.format_duration(result['duration'])
                self.add_result_row(
                    rank, result['address'], result['suburb'], result['state'],
                    f"{result['distance']:.2f}", formatted_duration,
                    result['status'], result['tag']
                )
        
        total = len(self.all_results)
        showing = len(filtered_results)
        if showing < total:
            self.status_var.set(f"🔍 Showing {showing} of {total} results (filtered)")
        else:
            self.status_var.set(f"📊 Showing all {total} results")
    
    def calculate_distances(self):
        """Start distance calculation in a separate thread"""
        tech_addr = self.tech_address.get("1.0", tk.END).strip()
        
        if not tech_addr:
            messagebox.showwarning("Input Required", "Please enter the technician's address")
            return
        
        if not self.site_addresses:
            messagebox.showwarning("Input Required", "Please add at least one site address")
            return
        
        if self.calculation_thread and self.calculation_thread.is_alive():
            messagebox.showwarning("Busy", "Calculation already in progress")
            return
        
        self.progress.grid()
        self.progress.set(0)
        self.calc_btn.configure(state="disabled")
        
        self.stop_calculation = False
        self.calculation_thread = threading.Thread(
            target=self.calculate_distances_worker,
            args=(tech_addr,),
            daemon=True
        )
        self.calculation_thread.start()
    
    def copy_results_smart(self):
        """Smart copy - copies selected cells if any, otherwise copies all results"""
        if self.selected_cells:
            self.copy_selected_cells()
        else:
            self.copy_all_results()
    
    def copy_all_results(self):
        """Copy all results to clipboard"""
        if not self.all_results:
            messagebox.showinfo("No Results", "No results to copy")
            return
        
        lines = ["Rank\tAddress\tSuburb\tState\tDistance (km)\tDuration (min)\tStatus"]
        
        for rank, result in enumerate(self.all_results, 1):
            if result['distance'] == float('inf'):
                line = f"{rank}\t{result['address']}\t{result['suburb']}\t{result['state']}\tN/A\tN/A\t{result['status']}"
            else:
                line = f"{rank}\t{result['address']}\t{result['suburb']}\t{result['state']}\t{result['distance']:.2f}\t{result['duration']:.0f}\t{result['status']}"
            lines.append(line)
        
        text = '\n'.join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        
        messagebox.showinfo("Copied", f"Copied {len(self.all_results)} results to clipboard!")
        self.status_var.set(f"✓ Copied {len(self.all_results)} results to clipboard")

if __name__ == "__main__":
    app = ctk.CTk()
    calculator = AddressDistanceCalculator(app)
    app.mainloop()
