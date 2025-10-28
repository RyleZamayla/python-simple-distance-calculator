import tkinter as tk
from tkinter import ttk, messagebox
import requests
import time
from geopy.distance import geodesic
import json
import re
import threading
from queue import Queue

class AddressDistanceCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Address Distance Calculator")
        self.root.geometry("1200x800")
        
        # Configure modern color scheme
        self.colors = {
            'bg_gradient_start': '#0f0c29',
            'bg_gradient_end': '#24243e',
            'glass_bg': '#ffffff',
            'glass_border': '#e0e0e0',
            'accent': '#6366f1',
            'accent_hover': '#4f46e5',
            'success': '#10b981',
            'warning': '#f59e0b',
            'error': '#ef4444',
            'text_primary': '#1f2937',
            'text_secondary': '#6b7280',
            'shadow': '#00000020'
        }
        
        self.root.configure(bg='#1a1a2e')
        
        # Store site addresses with geocoding cache
        self.site_addresses = []
        self.geocode_cache = {}
        self.all_results = []  # Store all results for filtering
        
        # Filter states
        self.filter_vars = {
            'success': tk.BooleanVar(value=True),
            'cached': tk.BooleanVar(value=True),
            'broad': tk.BooleanVar(value=True),
            'not_found': tk.BooleanVar(value=True)
        }
        
        # Threading control
        self.calculation_thread = None
        self.stop_calculation = False
        self.result_queue = Queue()
        
        # Column selection for copying
        self.selected_columns = set()
        self.is_selecting_columns = False
        
        # Cell selection for Excel-like copying
        self.cell_selection = {
            'active': False,
            'start_item': None,
            'start_col': None,
            'end_item': None,
            'end_col': None,
            'selected_cells': set()
        }
        
        self.setup_modern_styles()
        self.create_ui()
        
        # Start result queue processor
        self.process_queue()
    
    def setup_modern_styles(self):
        """Setup modern visual styles"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure Treeview colors
        style.configure("Modern.Treeview",
                       background="white",
                       foreground=self.colors['text_primary'],
                       fieldbackground="white",
                       borderwidth=0,
                       font=('Segoe UI', 9))
        style.configure("Modern.Treeview.Heading",
                       background=self.colors['accent'],
                       foreground="white",
                       borderwidth=0,
                       font=('Segoe UI', 9, 'bold'))
        style.map('Modern.Treeview.Heading',
                 background=[('active', self.colors['accent_hover'])])
        
        # Selected column heading style
        style.configure("Selected.Treeview.Heading",
                       background='#10b981',
                       foreground="white",
                       borderwidth=0,
                       font=('Segoe UI', 9, 'bold'))
        
        style.configure("Modern.Horizontal.TProgressbar",
                       troughcolor='#2d3748',
                       background=self.colors['accent'],
                       borderwidth=0,
                       thickness=6)
        
        # Checkbutton style
        style.configure("Modern.TCheckbutton",
                       background='white',
                       foreground=self.colors['text_primary'],
                       font=('Segoe UI', 9))
    
    def create_ui(self):
        """Create the user interface"""
        main_frame = tk.Frame(self.root, bg='#1a1a2e')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=20, pady=20)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Title
        title_frame = tk.Frame(main_frame, bg='#1a1a2e')
        title_frame.grid(row=0, column=0, sticky=tk.W, pady=(0, 20))
        
        title_label = tk.Label(title_frame, text="📍 Address Distance Calculator", 
                               font=('Segoe UI', 20, 'bold'),
                               bg='#1a1a2e', fg='white')
        title_label.pack(anchor=tk.W)
        
        subtitle_label = tk.Label(title_frame, text="Calculate optimal routes with precision • Powered by OSRM", 
                                 font=('Segoe UI', 10),
                                 bg='#1a1a2e', fg='#9ca3af')
        subtitle_label.pack(anchor=tk.W)
        
        # Technician Address Section
        tech_container = self.create_glass_frame(main_frame, "🏠 Technician Base Address")
        tech_container.container.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.tech_address = tk.Text(tech_container, height=2, width=80, 
                                    font=('Segoe UI', 10),
                                    bg='white', fg=self.colors['text_primary'],
                                    relief=tk.FLAT, borderwidth=0,
                                    insertbackground=self.colors['accent'])
        self.tech_address.grid(row=0, column=0, pady=8, padx=15, sticky=(tk.W, tk.E))
        
        # Site Addresses Section
        site_container = self.create_glass_frame(main_frame, "📋 Site Addresses")
        site_container.container.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        
        instruction_text = "💡 Paste addresses in any format: Excel columns, comma-separated, or line-by-line"
        instruction_label = tk.Label(site_container, text=instruction_text, 
                                    font=('Segoe UI', 9),
                                    bg='white', fg=self.colors['accent'])
        instruction_label.grid(row=0, column=0, sticky=tk.W, pady=(8, 8), padx=15)
        
        # Input table container
        input_container = tk.Frame(site_container, bg='white')
        input_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=15, pady=(0, 10))
        input_container.columnconfigure(0, weight=1)
        input_container.rowconfigure(0, weight=1)
        
        # Create input table
        input_columns = ('Status', 'Address', 'Suburb', 'State')
        self.input_tree = ttk.Treeview(input_container, columns=input_columns, 
                                       show='headings', height=6, style="Modern.Treeview")
        
        self.input_tree.heading('Status', text='Status')
        self.input_tree.heading('Address', text='Address')
        self.input_tree.heading('Suburb', text='Suburb')
        self.input_tree.heading('State', text='State')
        
        self.input_tree.column('Status', width=100, anchor='center')
        self.input_tree.column('Address', width=450)
        self.input_tree.column('Suburb', width=180)
        self.input_tree.column('State', width=120)
        
        self.input_tree.tag_configure('normal', background='white')
        self.input_tree.tag_configure('pending', background='#f3f4f6')
        self.input_tree.tag_configure('cached', background='#dbeafe', foreground='#1e40af')
        self.input_tree.tag_configure('retried', background='#fef3c7', foreground='#92400e')
        self.input_tree.tag_configure('error', background='#fee2e2', foreground='#991b1b')
        
        self.input_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        input_scrollbar = ttk.Scrollbar(input_container, orient=tk.VERTICAL, 
                                       command=self.input_tree.yview)
        input_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.input_tree.configure(yscrollcommand=input_scrollbar.set)
        
        # Paste area
        paste_container = tk.Frame(
            site_container,
            bg='#f3f4f6',
            highlightbackground=self.colors['accent'],
            highlightthickness=2,
            bd=0,
            relief=tk.FLAT
        )
        paste_container.grid(row=2, column=0, pady=(0, 10), padx=15, sticky=(tk.W, tk.E))
        paste_container.columnconfigure(1, weight=1)
        
        paste_label = tk.Label(paste_container, text="📋 Paste Area:", 
                              font=('Segoe UI', 9, 'bold'),
                              bg='white', fg=self.colors['text_primary'])
        paste_label.grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        
        self.paste_entry = tk.Text(paste_container, height=3, 
                                   font=('Segoe UI', 10),
                                   bg='#f9fafb', fg=self.colors['text_primary'],
                                   relief=tk.FLAT, borderwidth=1,
                                   insertbackground=self.colors['accent'])
        self.paste_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        self.paste_entry.bind('<Control-v>', self.handle_paste)
        self.paste_entry.bind('<Command-v>', self.handle_paste)
        
        # Buttons
        button_frame = tk.Frame(main_frame, bg='#1a1a2e')
        button_frame.grid(row=3, column=0, pady=(0, 15))
        
        self.add_btn = self.create_modern_button(button_frame, "➕ Add Addresses", 
                                                 self.process_pasted_data, 'accent')
        self.add_btn.pack(side=tk.LEFT, padx=5)
        
        self.remove_btn = self.create_modern_button(button_frame, "❌ Remove Selected", 
                                                     self.remove_selected, 'secondary')
        self.remove_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_btn = self.create_modern_button(button_frame, "🗑️ Clear All", 
                                                    self.clear_sites, 'secondary')
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        self.calc_btn = self.create_modern_button(button_frame, "🚀 Calculate Distances", 
                                                   self.calculate_distances, 'primary')
        self.calc_btn.pack(side=tk.LEFT, padx=15)
        
        self.copy_btn = self.create_modern_button(button_frame, "📄 Copy Selected Columns", 
                                                   self.copy_selected_columns_data, 'secondary')
        self.copy_btn.pack(side=tk.LEFT, padx=5)
        
        # Legend
        legend_frame = tk.Frame(main_frame, bg='#1a1a2e')
        legend_frame.grid(row=4, column=0, pady=(0, 15))
        
        legend_title = tk.Label(legend_frame, text="Legend:", 
                               font=('Segoe UI', 9, 'bold'),
                               bg='#1a1a2e', fg='white')
        legend_title.pack(side=tk.LEFT, padx=5)
        
        self.create_legend_item(legend_frame, " ✓ Success ", '#d1fae5', '#065f46')
        self.create_legend_item(legend_frame, " 💾 Cached ", '#dbeafe', '#1e40af')
        self.create_legend_item(legend_frame, " ⚠ Broad Match ", '#fef3c7', '#92400e')
        self.create_legend_item(legend_frame, " ✗ Not Found ", '#fee2e2', '#991b1b')
        
        # Results Table with Filters
        results_container = self.create_glass_frame(main_frame, "📊 Results (Ranked by Distance)")
        results_container.container.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        results_container.columnconfigure(0, weight=1)
        
        # Filter controls
        filter_frame = tk.Frame(results_container, bg='white')
        filter_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=15, pady=(8, 8))
        
        filter_label = tk.Label(filter_frame, text="🔍 Show:", 
                               font=('Segoe UI', 9, 'bold'),
                               bg='white', fg=self.colors['text_primary'])
        filter_label.pack(side=tk.LEFT, padx=(0, 10))
        
        # Filter checkboxes
        ttk.Checkbutton(filter_frame, text="✓ Found", 
                       variable=self.filter_vars['success'],
                       command=self.apply_filters,
                       style="Modern.TCheckbutton").pack(side=tk.LEFT, padx=5)
        
        ttk.Checkbutton(filter_frame, text="💾 Cached", 
                       variable=self.filter_vars['cached'],
                       command=self.apply_filters,
                       style="Modern.TCheckbutton").pack(side=tk.LEFT, padx=5)
        
        ttk.Checkbutton(filter_frame, text="⚠ Broad Match", 
                       variable=self.filter_vars['broad'],
                       command=self.apply_filters,
                       style="Modern.TCheckbutton").pack(side=tk.LEFT, padx=5)
        
        ttk.Checkbutton(filter_frame, text="✗ Not Found", 
                       variable=self.filter_vars['not_found'],
                       command=self.apply_filters,
                       style="Modern.TCheckbutton").pack(side=tk.LEFT, padx=5)
        
        # Column selection hint
        hint_label = tk.Label(filter_frame, text="💡 Drag to select cells, Ctrl+C to copy • Click headers for column copy", 
                             font=('Segoe UI', 8, 'italic'),
                             bg='white', fg=self.colors['text_secondary'])
        hint_label.pack(side=tk.RIGHT, padx=10)
        
        # Results treeview
        results_tree_container = tk.Frame(results_container, bg='white')
        results_tree_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=15, pady=(0, 10))
        results_tree_container.columnconfigure(0, weight=1)
        results_tree_container.rowconfigure(0, weight=1)
        
        result_columns = ('Rank', 'Address', 'Suburb', 'State', 'Distance (km)', 
                         'Duration (min)', 'Status')
        self.results_tree = ttk.Treeview(results_tree_container, columns=result_columns, 
                                        show='headings', height=8, style="Modern.Treeview")
        
        # Bind column header clicks for selection
        for col in result_columns:
            self.results_tree.heading(col, text=col, 
                                     command=lambda c=col: self.toggle_column_selection(c))
        
        self.results_tree.column('Rank', width=50, anchor='center')
        self.results_tree.column('Address', width=350)
        self.results_tree.column('Suburb', width=150)
        self.results_tree.column('State', width=100)
        self.results_tree.column('Distance (km)', width=120, anchor='center')
        self.results_tree.column('Duration (min)', width=120, anchor='center')
        self.results_tree.column('Status', width=150, anchor='center')
        
        self.results_tree.tag_configure('success', background='#d1fae5', foreground='#065f46')
        self.results_tree.tag_configure('cached', background='#dbeafe', foreground='#1e40af')
        self.results_tree.tag_configure('warning', background='#fef3c7', foreground='#92400e')
        self.results_tree.tag_configure('error', background='#fee2e2', foreground='#991b1b')
        
        self.results_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Enable Excel-like cell selection
        self.results_tree.bind('<Button-1>', self.on_cell_click)
        self.results_tree.bind('<B1-Motion>', self.on_cell_drag)
        self.results_tree.bind('<ButtonRelease-1>', self.on_cell_release)
        
        # Enable copying from results tree
        self.results_tree.bind('<Control-c>', self.copy_selected_cells)
        self.results_tree.bind('<Command-c>', self.copy_selected_cells)
        
        results_scrollbar = ttk.Scrollbar(results_tree_container, orient=tk.VERTICAL, 
                                         command=self.results_tree.yview)
        results_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.results_tree.configure(yscrollcommand=results_scrollbar.set)
        
        # Status bar
        status_frame = tk.Frame(main_frame, bg='#1a1a2e')
        status_frame.grid(row=6, column=0, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(0, weight=1)
        
        self.status_var = tk.StringVar()
        self.status_var.set("✓ Ready - Paste addresses and click 'Add Addresses'")
        
        status_bar = tk.Label(status_frame, textvariable=self.status_var, 
                             bg='#2d3748', fg='white',
                             anchor=tk.W, font=('Segoe UI', 9),
                             padx=15, pady=8)
        status_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        
        self.progress = ttk.Progressbar(status_frame, mode='determinate', 
                                       style="Modern.Horizontal.TProgressbar")
        self.progress.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.progress.grid_remove()
        
        # Configure row weights
        main_frame.rowconfigure(2, weight=1)
        main_frame.rowconfigure(5, weight=2)
        results_container.rowconfigure(1, weight=1)
    
    def create_glass_frame(self, parent, title):
        """Create a frame with glass morphism effect"""
        container = tk.Frame(parent, bg='#1a1a2e')
        container.configure(highlightbackground='#e5e7eb', highlightthickness=1)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)
        
        title_bar = tk.Frame(container, bg='white', height=35)
        title_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))
        title_bar.grid_propagate(False)
        
        title_label = tk.Label(title_bar, text=title, 
                              font=('Segoe UI', 11, 'bold'),
                              bg='white', fg=self.colors['text_primary'],
                              anchor=tk.W, padx=15)
        title_label.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=8)
        
        content = tk.Frame(container, bg='white', relief=tk.FLAT, borderwidth=0)
        content.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)
        content.container = container
        
        return content
    
    def create_modern_button(self, parent, text, command, style_type='secondary'):
        """Create a modern styled button"""
        if style_type == 'primary':
            bg_color = self.colors['accent']
            fg_color = 'white'
            hover_bg = self.colors['accent_hover']
        elif style_type == 'accent':
            bg_color = self.colors['success']
            fg_color = 'white'
            hover_bg = '#059669'
        else:
            bg_color = '#6b7280'
            fg_color = 'white'
            hover_bg = '#4b5563'
        
        button = tk.Button(parent, text=text, command=command,
                          bg=bg_color, fg=fg_color,
                          font=('Segoe UI', 9, 'bold'),
                          relief=tk.FLAT, borderwidth=0,
                          padx=20, pady=10,
                          cursor='hand2')
        
        button.bind('<Enter>', lambda e: button.config(bg=hover_bg))
        button.bind('<Leave>', lambda e: button.config(bg=bg_color))
        
        return button
    
    def create_legend_item(self, parent, text, bg, fg):
        """Create a legend item"""
        label = tk.Label(parent, text=text, bg=bg, fg=fg,
                        font=('Segoe UI', 8, 'bold'),
                        padx=8, pady=4)
        label.pack(side=tk.LEFT, padx=3)
    
    def on_cell_click(self, event):
        """Handle cell click for Excel-like selection"""
        region = self.results_tree.identify_region(event.x, event.y)
        
        if region == "cell":
            item = self.results_tree.identify_row(event.y)
            column = self.results_tree.identify_column(event.x)
            
            if item and column:
                # Clear previous selection
                self.clear_cell_highlights()
                
                # Start new selection
                self.cell_selection['active'] = True
                self.cell_selection['start_item'] = item
                self.cell_selection['start_col'] = column
                self.cell_selection['end_item'] = item
                self.cell_selection['end_col'] = column
                self.cell_selection['selected_cells'] = {(item, column)}
                
                # Highlight selected cell
                self.highlight_selected_cells()
    
    def on_cell_drag(self, event):
        """Handle cell drag for selecting multiple cells"""
        if not self.cell_selection['active']:
            return
        
        region = self.results_tree.identify_region(event.x, event.y)
        
        if region == "cell":
            item = self.results_tree.identify_row(event.y)
            column = self.results_tree.identify_column(event.x)
            
            if item and column:
                self.cell_selection['end_item'] = item
                self.cell_selection['end_col'] = column
                
                # Update selected cells
                self.update_cell_selection()
                self.highlight_selected_cells()
    
    def on_cell_release(self, event):
        """Handle mouse release to finalize selection"""
        if self.cell_selection['active']:
            self.cell_selection['active'] = False
    
    def update_cell_selection(self):
        """Update the set of selected cells based on start and end"""
        self.cell_selection['selected_cells'] = set()
        
        # Get all items and columns
        all_items = self.results_tree.get_children()
        all_columns = self.results_tree['columns']
        
        # Find indices
        start_item = self.cell_selection['start_item']
        end_item = self.cell_selection['end_item']
        start_col = self.cell_selection['start_col']
        end_col = self.cell_selection['end_col']
        
        if start_item not in all_items or end_item not in all_items:
            return
        
        start_row_idx = all_items.index(start_item)
        end_row_idx = all_items.index(end_item)
        
        # Normalize column identifiers (they're like '#1', '#2', etc.)
        start_col_idx = int(start_col.replace('#', '')) - 1
        end_col_idx = int(end_col.replace('#', '')) - 1
        
        # Ensure start is before end
        min_row = min(start_row_idx, end_row_idx)
        max_row = max(start_row_idx, end_row_idx)
        min_col = min(start_col_idx, end_col_idx)
        max_col = max(start_col_idx, end_col_idx)
        
        # Select all cells in the range
        for row_idx in range(min_row, max_row + 1):
            for col_idx in range(min_col, max_col + 1):
                item = all_items[row_idx]
                col = f'#{col_idx + 1}'
                self.cell_selection['selected_cells'].add((item, col))
    
    def highlight_selected_cells(self):
        """Highlight selected cells (visual feedback)"""
        # Clear previous highlights
        self.clear_cell_highlights()
        
        # We can't directly highlight individual cells in ttk.Treeview,
        # but we can select rows
        selected_items = {item for item, col in self.cell_selection['selected_cells']}
        self.results_tree.selection_set(list(selected_items))
    
    def clear_cell_highlights(self):
        """Clear cell highlights"""
        self.results_tree.selection_remove(self.results_tree.selection())
    
    def copy_selected_cells(self, event=None):
        """Copy selected cells in Excel format"""
        if not self.cell_selection['selected_cells']:
            # Fallback to copying selected rows
            return self.copy_selected_rows(event)
        
        # Get all items and columns
        all_items = self.results_tree.get_children()
        all_columns = self.results_tree['columns']
        
        # Organize selected cells by row and column
        cells_by_row = {}
        for item, col in self.cell_selection['selected_cells']:
            if item not in all_items:
                continue
            
            row_idx = all_items.index(item)
            col_idx = int(col.replace('#', '')) - 1
            
            if row_idx not in cells_by_row:
                cells_by_row[row_idx] = {}
            
            values = self.results_tree.item(item, 'values')
            if col_idx < len(values):
                cells_by_row[row_idx][col_idx] = str(values[col_idx])
        
        # Build text output
        lines = []
        for row_idx in sorted(cells_by_row.keys()):
            cols = cells_by_row[row_idx]
            # Get column range
            if cols:
                min_col = min(cols.keys())
                max_col = max(cols.keys())
                
                # Build row with all columns in range (empty if not selected)
                row_values = []
                for col_idx in range(min_col, max_col + 1):
                    row_values.append(cols.get(col_idx, ''))
                
                lines.append('\t'.join(row_values))
        
        if lines:
            text = '\n'.join(lines)
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            
            num_cells = len(self.cell_selection['selected_cells'])
            num_rows = len(cells_by_row)
            num_cols = max(len(cells) for cells in cells_by_row.values()) if cells_by_row else 0
            
            self.status_var.set(f"✓ Copied {num_rows} row(s) × {num_cols} column(s) ({num_cells} cells)")
    
    def copy_selected_rows(self, event=None):
        """Copy selected rows with all columns"""
        selected_items = self.results_tree.selection()
        if not selected_items:
            return
        
        lines = []
        for item in selected_items:
            values = self.results_tree.item(item, 'values')
            lines.append('\t'.join(str(v) for v in values))
        
        text = '\n'.join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        
        self.status_var.set(f"✓ Copied {len(selected_items)} row(s) to clipboard")
    
    def toggle_column_selection(self, column):
        """Toggle column selection for copying"""
        if column in self.selected_columns:
            self.selected_columns.remove(column)
            # Reset to normal style
            self.results_tree.heading(column, text=column)
        else:
            self.selected_columns.add(column)
            # Highlight selected column
            self.results_tree.heading(column, text=f"✓ {column}")
        
        # Update status
        if self.selected_columns:
            cols = ', '.join(self.selected_columns)
            self.status_var.set(f"📋 Selected columns: {cols}")
        else:
            self.status_var.set("💡 Click column headers to select for copying")
    
    def copy_selected_columns_data(self):
        """Copy only selected columns from all visible results"""
        if not self.selected_columns:
            messagebox.showinfo("No Columns Selected", 
                              "Please click on column headers to select which columns to copy.")
            return
        
        if not self.results_tree.get_children():
            messagebox.showinfo("No Results", "No results to copy")
            return
        
        # Get all columns
        all_columns = self.results_tree['columns']
        selected_indices = [all_columns.index(col) for col in self.selected_columns]
        selected_indices.sort()
        
        # Build header
        headers = [all_columns[i] for i in selected_indices]
        lines = ['\t'.join(headers)]
        
        # Add data rows
        for item in self.results_tree.get_children():
            values = self.results_tree.item(item, 'values')
            selected_values = [str(values[i]) for i in selected_indices]
            lines.append('\t'.join(selected_values))
        
        text = '\n'.join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        
        col_names = ', '.join(headers)
        messagebox.showinfo("Copied", 
                          f"Copied {len(lines)-1} rows with columns:\n{col_names}\n\nYou can paste directly into Excel or other spreadsheets.")
        self.status_var.set(f"✓ Copied {len(self.selected_columns)} columns × {len(lines)-1} rows")
    
    def copy_selected_rows(self, event=None):
        """Copy selected rows with all columns"""
        selected_items = self.results_tree.selection()
        if not selected_items:
            return
        
        lines = []
        for item in selected_items:
            values = self.results_tree.item(item, 'values')
            lines.append('\t'.join(str(v) for v in values))
        
        text = '\n'.join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        
        self.status_var.set(f"✓ Copied {len(selected_items)} row(s) to clipboard")
    
    def apply_filters(self):
        """Apply filters to results display"""
        if not self.all_results:
            return
        
        # Clear current display
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        # Filter results based on checkboxes
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
        
        # Display filtered results
        for rank, result in enumerate(filtered_results, 1):
            if result['distance'] == float('inf'):
                self.results_tree.insert('', tk.END, values=(
                    rank,
                    result['address'],
                    result['suburb'],
                    result['state'],
                    'N/A',
                    'N/A',
                    result['status']
                ), tags=(result['tag'],))
            else:
                tag = result['tag']
                if tag == 'cached':
                    display_tag = 'cached'
                elif tag == 'normal' or tag == 'success':
                    display_tag = 'success'
                else:
                    display_tag = 'warning'
                
                self.results_tree.insert('', tk.END, values=(
                    rank,
                    result['address'],
                    result['suburb'],
                    result['state'],
                    f"{result['distance']:.2f}",
                    f"{result['duration']:.0f}",
                    result['status']
                ), tags=(display_tag,))
        
        # Clear cell selection when filters change
        self.cell_selection['selected_cells'] = set()
        
        # Update status
        total = len(self.all_results)
        showing = len(filtered_results)
        if showing < total:
            self.status_var.set(f"🔍 Showing {showing} of {total} results (filtered)")
        else:
            self.status_var.set(f"📊 Showing all {total} results")
    
    def handle_paste(self, event=None):
        """Handle paste event"""
        self.root.after(100, self.auto_process_hint)
        return None
    
    def auto_process_hint(self):
        """Show hint after pasting"""
        pasted_text = self.paste_entry.get("1.0", tk.END).strip()
        if pasted_text:
            self.status_var.set("📋 Data pasted! Click 'Add Addresses' to process")
    
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
                    
                    self.input_tree.insert('', tk.END, 
                                         values=(status, address, suburb, state),
                                         tags=(tag,))
                    
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
                                     "• Full address: Street, Suburb, STATE 1234\n" +
                                     "• Line-by-line in any above format")
                self.status_var.set("⚠ No valid addresses found")
    
    def remove_selected(self):
        """Remove selected addresses"""
        selected_items = self.input_tree.selection()
        if not selected_items:
            messagebox.showinfo("No Selection", "Please select an address to remove")
            return
        
        removed_count = 0
        for item in selected_items:
            values = self.input_tree.item(item, 'values')
            self.site_addresses = [addr for addr in self.site_addresses 
                                   if not (addr['address'] == values[1] and 
                                          addr['suburb'] == values[2] and 
                                          addr['state'] == values[3])]
            self.input_tree.delete(item)
            removed_count += 1
        
        self.status_var.set(f"✓ Removed {removed_count} address(es). Total: {len(self.site_addresses)}")
    
    def clear_sites(self):
        """Clear all site addresses"""
        if self.site_addresses:
            if messagebox.askyesno("Confirm Clear", 
                                  f"Clear all {len(self.site_addresses)} addresses?"):
                self.site_addresses = []
                for item in self.input_tree.get_children():
                    self.input_tree.delete(item)
                self.status_var.set("✓ All addresses cleared")
        else:
            self.status_var.set("ℹ No addresses to clear")
    
    def geocode_address_incremental(self, address, max_retries=4):
        """
        Geocode address with incremental broader search strategy
        
        Strategy:
        1. Try exact address (e.g., "Shop 6, 5 Halley Street, Chisholm, ACT")
        2. Remove shop/unit (e.g., "5 Halley Street, Chisholm, ACT")
        3. Try suburb + state only (e.g., "Chisholm, ACT")
        4. Try major landmarks if applicable (e.g., "Shopping Centre Name, Suburb, State")
        5. Try state only as last resort (e.g., "ACT, Australia")
        """
        base_url = "https://nominatim.openstreetmap.org/search"
        headers = {'User-Agent': 'AddressDistanceCalculator/3.0'}
        
        parts = [p.strip() for p in address.split(',')]
        
        # Build search attempts from most specific to broader
        search_attempts = []
        
        # Attempt 1: Full address
        search_attempts.append({
            'query': address,
            'level': 0,
            'description': 'exact address'
        })
        
        # Attempt 2: Remove shop/unit numbers and descriptors
        if len(parts) >= 3:
            street_part = parts[0]
            # Remove common prefixes
            cleaned_street = re.sub(r'^(Shop|Unit|Suite|Level|T/a|Tenancy|Lot)\s*\d+[A-Za-z]?,?\s*', '', 
                                   street_part, flags=re.IGNORECASE)
            cleaned_street = re.sub(r'^(Shops?|Units?)\s+\d+\s*[&-]\s*\d+,?\s*', '', 
                                   cleaned_street, flags=re.IGNORECASE)
            
            if cleaned_street != street_part and cleaned_street.strip():
                search_attempts.append({
                    'query': ', '.join([cleaned_street] + parts[1:]),
                    'level': 1,
                    'description': 'without shop/unit number'
                })
        
        # Attempt 3: Extract shopping centre name if present
        shopping_centre_match = re.search(r'([\w\s]+(?:Shopping Centre|Market|Plaza|Fair|Marketplace))', 
                                         address, re.IGNORECASE)
        if shopping_centre_match and len(parts) >= 2:
            centre_name = shopping_centre_match.group(1).strip()
            search_attempts.append({
                'query': f"{centre_name}, {parts[-2]}, {parts[-1]}",
                'level': 1,
                'description': 'shopping centre with suburb'
            })
        
        # Attempt 4: Street name + suburb + state (remove building number)
        if len(parts) >= 3:
            street_name = re.sub(r'^\d+[A-Za-z]?\s*', '', parts[0])  # Remove leading number
            street_name = re.sub(r'^(Shop|Unit|Suite|Level|T/a|Tenancy)\s*.*?,?\s*', '', 
                                street_name, flags=re.IGNORECASE)
            if street_name.strip() and len(street_name.strip()) > 3:
                search_attempts.append({
                    'query': ', '.join([street_name.strip()] + parts[1:]),
                    'level': 2,
                    'description': 'street name with suburb'
                })
        
        # Attempt 5: Suburb + State only (most common fallback)
        if len(parts) >= 2:
            search_attempts.append({
                'query': f"{parts[-2]}, {parts[-1]}",
                'level': 3,
                'description': 'suburb and state'
            })
        
        # Attempt 6: State + Australia (last resort)
        if len(parts) >= 1:
            search_attempts.append({
                'query': f"{parts[-1]}, Australia",
                'level': 4,
                'description': 'state only'
            })
        
        # Try each search attempt
        for attempt_num, attempt in enumerate(search_attempts[:max_retries], 1):
            if attempt_num > 1:
                time.sleep(0.5)  # Small delay between retries
            
            params = {
                'q': attempt['query'],
                'format': 'json',
                'limit': 3,  # Get top 3 results for better matching
                'countrycodes': 'au',
                'addressdetails': 1
            }
            
            try:
                response = requests.get(base_url, params=params, headers=headers, timeout=10)
                response.raise_for_status()
                data = response.json()
                
                if data:
                    # For suburb-level searches, try to match the actual suburb
                    if attempt['level'] >= 2 and len(parts) >= 2:
                        suburb_to_match = parts[-2].lower().strip()
                        
                        # Look for best match in results
                        for result in data:
                            result_address = result.get('address', {})
                            result_suburb = (result_address.get('suburb') or 
                                           result_address.get('city') or 
                                           result_address.get('town') or '').lower()
                            
                            if suburb_to_match in result_suburb or result_suburb in suburb_to_match:
                                return (float(result['lat']), 
                                       float(result['lon']), 
                                       attempt['level'],
                                       attempt['description'])
                    
                    # Return first result if no specific match found
                    return (float(data[0]['lat']), 
                           float(data[0]['lon']), 
                           attempt['level'],
                           attempt['description'])
                
            except Exception as e:
                print(f"Attempt {attempt_num} failed for '{attempt['query']}': {e}")
                continue
        
        # All attempts failed
        return None, None, None, None
    
    def get_osrm_route(self, coord1, coord2):
        """Get actual route distance and duration using OSRM"""
        try:
            lon1, lat1 = coord1[1], coord1[0]
            lon2, lat2 = coord2[1], coord2[0]
            
            url = f"http://router.project-osrm.org/route/v1/driving/{lon1},{lat1};{lon2},{lat2}"
            params = {
                'overview': 'false',
                'steps': 'false'
            }
            
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
    
    def update_input_tree_status(self, address, suburb, state, status, tag):
        """Update status in input tree"""
        for item in self.input_tree.get_children():
            values = self.input_tree.item(item, 'values')
            if values[1] == address and values[2] == suburb and values[3] == state:
                self.input_tree.item(item, values=(status, address, suburb, state), tags=(tag,))
                break
    
    def calculate_distances_worker(self, tech_addr):
        """Worker function for calculating distances in a separate thread"""
        try:
            self.result_queue.put(('status', "🔍 Geocoding technician address..."))
            tech_lat, tech_lon, tech_level, tech_desc = self.geocode_address_incremental(tech_addr)
            
            if not tech_lat:
                self.result_queue.put(('error', "Could not geocode technician address"))
                return
            
            tech_coords = (tech_lat, tech_lon)
            self.result_queue.put(('progress', 1))
            time.sleep(0.5)
            
            results = []
            total = len(self.site_addresses)
            
            for i, site in enumerate(self.site_addresses):
                if self.stop_calculation:
                    self.result_queue.put(('status', "❌ Calculation cancelled"))
                    return
                
                progress_percent = int(((i + 1) / total) * 100)
                full_address = f"{site['address']}, {site['suburb']}, {site['state']}"
                cache_key = full_address.lower()
                
                if 'lat' in site and 'lon' in site and site['lat'] and site['lon']:
                    site_lat = site['lat']
                    site_lon = site['lon']
                    match_level = site.get('match_level', 0)
                    match_desc = site.get('match_desc', 'exact')
                    
                    self.result_queue.put(('status', f"💾 Using cached data ({progress_percent}%): {site['suburb']}"))
                    
                    site_coords = (site_lat, site_lon)
                    distance_km, duration_min = self.get_osrm_route(tech_coords, site_coords)
                    
                    # Determine status based on match level
                    if match_level == 0:
                        status = "💾 Cached"
                        tag = 'cached'
                    else:
                        status = f"💾 Cached ({match_desc})"
                        tag = 'cached'
                    
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
                    self.result_queue.put(('update_input', (site['address'], site['suburb'], site['state'], status, tag)))
                    self.result_queue.put(('progress', i + 2))
                    continue
                
                self.result_queue.put(('status', f"⏳ Geocoding ({progress_percent}%): {site['suburb']}"))
                
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
                    site['status'] = 'geocoded'
                    
                    # Determine status based on match quality
                    if match_level == 0:
                        status = "✓ Found (exact)"
                        tag = 'success'
                    elif match_level == 1:
                        status = f"✓ Found ({match_desc})"
                        tag = 'success'
                    elif match_level == 2:
                        status = f"⚠ Approximate ({match_desc})"
                        tag = 'warning'
                    else:
                        status = f"⚠ Broad ({match_desc})"
                        tag = 'warning'
                    
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
                    self.result_queue.put(('update_input', (site['address'], site['suburb'], site['state'], status, tag)))
                else:
                    result = {
                        'address': site['address'],
                        'suburb': site['suburb'],
                        'state': site['state'],
                        'distance': float('inf'),
                        'duration': float('inf'),
                        'status': '✗ Not Found',
                        'tag': 'error',
                        'match_level': 999
                    }
                    results.append(result)
                    self.result_queue.put(('update_input', (site['address'], site['suburb'], site['state'], '✗ Not Found', 'error')))
                
                self.result_queue.put(('progress', i + 2))
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
                    self.progress['value'] = data
                elif msg_type == 'update_input':
                    self.update_input_tree_status(*data)
                elif msg_type == 'results':
                    self.all_results = data
                    self.apply_filters()
                elif msg_type == 'complete':
                    self.calculation_complete()
                elif msg_type == 'error':
                    self.handle_calculation_error(data)
        except:
            pass
        
        self.root.after(100, self.process_queue)
    
    def calculation_complete(self):
        """Handle calculation completion"""
        self.progress.grid_remove()
        self.calc_btn.config(state=tk.NORMAL)
        
        # Count by category
        counts = {
            'success': sum(1 for r in self.all_results if r['tag'] == 'success'),
            'cached': sum(1 for r in self.all_results if r['tag'] == 'cached'),
            'retry': sum(1 for r in self.all_results if r['tag'] == 'warning'),
            'error': sum(1 for r in self.all_results if r['tag'] == 'error')
        }
        
        summary_parts = []
        if counts['success'] > 0:
            summary_parts.append(f"{counts['success']} found")
        if counts['cached'] > 0:
            summary_parts.append(f"{counts['cached']} cached")
        if counts['retry'] > 0:
            summary_parts.append(f"{counts['retry']} broad matches")
        if counts['error'] > 0:
            summary_parts.append(f"{counts['error']} not found")
        
        summary = f"✓ Complete! " + ", ".join(summary_parts)
        self.status_var.set(summary)
        
        if counts['error'] > 0 or counts['cached'] > 0:
            msg_parts = ["Calculation complete!\n"]
            if counts['success'] > 0:
                msg_parts.append(f"✓ Successfully geocoded: {counts['success']}")
            if counts['cached'] > 0:
                msg_parts.append(f"💾 Used cached data: {counts['cached']}")
            if counts['retry'] > 0:
                msg_parts.append(f"⚠ Broad matches: {counts['retry']}")
            if counts['error'] > 0:
                msg_parts.append(f"✗ Not found: {counts['error']}")
            msg_parts.append("\nUse filters to show/hide specific result types.")
            
            messagebox.showinfo("Results Summary", "\n".join(msg_parts))
    
    def handle_calculation_error(self, error_msg):
        """Handle calculation errors"""
        self.progress.grid_remove()
        self.calc_btn.config(state=tk.NORMAL)
        messagebox.showerror("Calculation Error", error_msg)
        self.status_var.set(f"✗ Error: {error_msg}")
    
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
        
        addresses_to_geocode = [s for s in self.site_addresses 
                                if 'lat' not in s or 'lon' not in s or not s.get('lat') or not s.get('lon')]
        
        self.progress.grid()
        total_steps = len(addresses_to_geocode) + 1
        self.progress['maximum'] = total_steps
        self.progress['value'] = 0
        
        self.calc_btn.config(state=tk.DISABLED)
        
        self.stop_calculation = False
        self.calculation_thread = threading.Thread(
            target=self.calculate_distances_worker,
            args=(tech_addr,),
            daemon=True
        )
        self.calculation_thread.start()

if __name__ == "__main__":
    root = tk.Tk()
    app = AddressDistanceCalculator(root)
    root.mainloop()