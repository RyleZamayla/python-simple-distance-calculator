import tkinter as tk
from tkinter import ttk, messagebox
import requests
import time
from geopy.distance import geodesic
import json
import re

class AddressDistanceCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Address Distance Calculator")
        self.root.geometry("1200x750")
        
        # Configure modern color scheme with glass effect
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
        
        # Set background with gradient effect
        self.root.configure(bg='#1a1a2e')
        
        # Store site addresses with geocoding cache
        self.site_addresses = []
        self.geocode_cache = {}  # Cache for already geocoded addresses
        
        # Configure modern styles
        self.setup_modern_styles()
        
        # Create main frame with glass effect
        main_frame = tk.Frame(root, bg='#1a1a2e')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=20, pady=20)
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Title with modern styling
        title_frame = tk.Frame(main_frame, bg='#1a1a2e')
        title_frame.grid(row=0, column=0, sticky=tk.W, pady=(0, 20))
        
        title_label = tk.Label(title_frame, text="📍 Address Distance Calculator", 
                               font=('Segoe UI', 20, 'bold'),
                               bg='#1a1a2e', fg='white')
        title_label.pack(anchor=tk.W)
        
        subtitle_label = tk.Label(title_frame, text="Calculate optimal routes with precision", 
                                 font=('Segoe UI', 10),
                                 bg='#1a1a2e', fg='#9ca3af')
        subtitle_label.pack(anchor=tk.W)
        
        # Technician Address Section with glass effect
        tech_container = self.create_glass_frame(main_frame, "🏠 Technician Base Address")
        tech_container.container.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.tech_address = tk.Text(tech_container, height=2, width=80, 
                                    font=('Segoe UI', 10),
                                    bg='white', fg=self.colors['text_primary'],
                                    relief=tk.FLAT, borderwidth=0,
                                    insertbackground=self.colors['accent'])
        self.tech_address.grid(row=0, column=0, pady=8, padx=15, sticky=(tk.W, tk.E))
        
        # Site Addresses Section with glass effect
        site_container = self.create_glass_frame(main_frame, "📋 Site Addresses")
        site_container.container.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        
        # Instructions with modern styling
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
        
        # Create input table with modern styling
        input_columns = ('Status', 'Address', 'Suburb', 'State')
        
        # Custom style for treeview
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
        
        self.input_tree = ttk.Treeview(input_container, columns=input_columns, 
                                       show='headings', height=6, style="Modern.Treeview")
        
        # Define headings and columns
        self.input_tree.heading('Status', text='Status')
        self.input_tree.heading('Address', text='Address')
        self.input_tree.heading('Suburb', text='Suburb')
        self.input_tree.heading('State', text='State')
        
        self.input_tree.column('Status', width=100, anchor='center')
        self.input_tree.column('Address', width=450)
        self.input_tree.column('Suburb', width=180)
        self.input_tree.column('State', width=120)
        
        # Configure tags with modern colors
        self.input_tree.tag_configure('normal', background='white')
        self.input_tree.tag_configure('pending', background='#f3f4f6')
        self.input_tree.tag_configure('cached', background='#dbeafe', foreground='#1e40af')
        self.input_tree.tag_configure('retried', background='#fef3c7', foreground='#92400e')
        self.input_tree.tag_configure('error', background='#fee2e2', foreground='#991b1b')
        
        self.input_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for input table with modern styling
        input_scrollbar = ttk.Scrollbar(input_container, orient=tk.VERTICAL, 
                                       command=self.input_tree.yview)
        input_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.input_tree.configure(yscrollcommand=input_scrollbar.set)
        
        # Paste area with modern styling and enhanced visibility
        paste_container = tk.Frame(
            site_container,
            bg='#f3f4f6',  # subtle light gray for contrast
            highlightbackground=self.colors['accent'],  # accent border
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
        
        # Buttons with modern styling
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
        
        # Legend with modern styling
        legend_frame = tk.Frame(main_frame, bg='#1a1a2e')
        legend_frame.grid(row=4, column=0, pady=(0, 15))
        
        legend_title = tk.Label(legend_frame, text="Legend:", 
                               font=('Segoe UI', 9, 'bold'),
                               bg='#1a1a2e', fg='white')
        legend_title.pack(side=tk.LEFT, padx=5)
        
        # Color indicators with glass effect
        self.create_legend_item(legend_frame, " ✓ Success ", '#d1fae5', '#065f46')
        self.create_legend_item(legend_frame, " 💾 Cached ", '#dbeafe', '#1e40af')
        self.create_legend_item(legend_frame, " ⚠ Broad Match ", '#fef3c7', '#92400e')
        self.create_legend_item(legend_frame, " ✗ Not Found ", '#fee2e2', '#991b1b')
        
        # Results Table with glass effect
        results_container = self.create_glass_frame(main_frame, "📊 Results (Ranked by Distance)")
        results_container.container.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        results_container.columnconfigure(0, weight=1)
        results_container.columnconfigure(1, minsize=20)
        
        # Create Results Treeview with modern styling
        result_columns = ('Rank', 'Address', 'Suburb', 'State', 'Distance (km)', 
                         'Duration (min)', 'Status')
        self.results_tree = ttk.Treeview(results_container, columns=result_columns, 
                                        show='headings', height=8, style="Modern.Treeview")
        
        # Define headings
        self.results_tree.heading('Rank', text='#')
        self.results_tree.heading('Address', text='Address')
        self.results_tree.heading('Suburb', text='Suburb')
        self.results_tree.heading('State', text='State')
        self.results_tree.heading('Distance (km)', text='Distance (km)')
        self.results_tree.heading('Duration (min)', text='Duration (min)')
        self.results_tree.heading('Status', text='Status')
        
        self.results_tree.column('Rank', width=50, anchor='center')
        self.results_tree.column('Address', width=350)
        self.results_tree.column('Suburb', width=150)
        self.results_tree.column('State', width=100)
        self.results_tree.column('Distance (km)', width=120, anchor='center')
        self.results_tree.column('Duration (min)', width=120, anchor='center')
        self.results_tree.column('Status', width=150, anchor='center')
        
        # Configure result tags with modern colors
        self.results_tree.tag_configure('success', background='#d1fae5', foreground='#065f46')
        self.results_tree.tag_configure('cached', background='#dbeafe', foreground='#1e40af')
        self.results_tree.tag_configure('warning', background='#fef3c7', foreground='#92400e')
        self.results_tree.tag_configure('error', background='#fee2e2', foreground='#991b1b')
        
        self.results_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=15, pady=10)
        
        # Scrollbar for results
        results_scrollbar = ttk.Scrollbar(results_container, orient=tk.VERTICAL, 
                                         command=self.results_tree.yview)
        results_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S), pady=10)
        self.results_tree.configure(yscrollcommand=results_scrollbar.set)
        
        # Status bar with modern styling
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
        
        # Progress bar with modern styling
        style.configure("Modern.Horizontal.TProgressbar",
                       troughcolor='#2d3748',
                       background=self.colors['accent'],
                       borderwidth=0,
                       thickness=6)
        
        self.progress = ttk.Progressbar(status_frame, mode='determinate', 
                                       style="Modern.Horizontal.TProgressbar")
        self.progress.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.progress.grid_remove()  # Hide initially
        
        # Configure row weights for resizing
        main_frame.rowconfigure(2, weight=1)
        main_frame.rowconfigure(5, weight=2)
    
    def setup_modern_styles(self):
        """Setup modern visual styles"""
        pass  # Styles are now configured inline
    
    def create_glass_frame(self, parent, title):
        """Create a frame with glass morphism effect"""
        # Main container
        container = tk.Frame(parent, bg='#1a1a2e')
        container.configure(highlightbackground='#e5e7eb', highlightthickness=1)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)
        
        # Title bar
        title_bar = tk.Frame(container, bg='white', height=35)
        title_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))
        title_bar.grid_propagate(False)
        
        title_label = tk.Label(title_bar, text=title, 
                              font=('Segoe UI', 11, 'bold'),
                              bg='white', fg=self.colors['text_primary'],
                              anchor=tk.W, padx=15)
        title_label.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=8)
        
        # Content frame with white background (THIS is what we return and use)
        content = tk.Frame(container, bg='white', relief=tk.FLAT, borderwidth=0)
        content.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)  # Row 1 for expandable content
        
        # Store reference to container in the content frame
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
        
        # Hover effects
        button.bind('<Enter>', lambda e: button.config(bg=hover_bg))
        button.bind('<Leave>', lambda e: button.config(bg=bg_color))
        
        return button
    
    def create_legend_item(self, parent, text, bg, fg):
        """Create a legend item with modern styling"""
        label = tk.Label(parent, text=text, bg=bg, fg=fg,
                        font=('Segoe UI', 8, 'bold'),
                        padx=8, pady=4)
        label.pack(side=tk.LEFT, padx=3)
    
    def handle_paste(self, event=None):
        """Handle paste event to auto-process after a short delay"""
        self.root.after(100, self.auto_process_hint)
        return None
    
    def auto_process_hint(self):
        """Show hint after pasting"""
        pasted_text = self.paste_entry.get("1.0", tk.END).strip()
        if pasted_text:
            self.status_var.set("📋 Data pasted! Click 'Add Addresses' to process")
    
    
    def parse_address_line(self, line):
        """Parse a single line of address data in various formats"""
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
                # Last part is likely state, second-to-last is suburb
                return {
                    'address': ', '.join(parts[:-2]),
                    'suburb': parts[-2],
                    'state': parts[-1]
                }
            elif len(parts) == 2:
                # Suburb, State format (address missing)
                return {'address': '', 'suburb': parts[0], 'state': parts[1]}
        
        # Try pipe-separated
        if '|' in line:
            parts = [p.strip() for p in line.split('|') if p.strip()]
            if len(parts) >= 3:
                return {'address': parts[0], 'suburb': parts[1], 'state': parts[2]}
        
        # Single line - try to detect if it's a full address
        # Pattern: street, suburb, state postcode
        parts = [p.strip() for p in line.split(',')]
        if len(parts) >= 2:
            # Check if last part looks like "STATE POSTCODE" or just "STATE"
            last_part = parts[-1].strip()
            state_match = re.match(r'^([A-Z]{2,3})(\s+\d{4})?$', last_part)
            if state_match:
                state = state_match.group(1)
                suburb = parts[-2] if len(parts) >= 2 else ''
                address = ', '.join(parts[:-2]) if len(parts) > 2 else ''
                return {'address': address, 'suburb': suburb, 'state': state}
        
        return None
    
    def process_pasted_data(self):
        """Process pasted data from multiple formats"""
        pasted_text = self.paste_entry.get("1.0", tk.END).strip()
        
        if not pasted_text:
            messagebox.showinfo("No Data", "Please paste address data first")
            return
        
        # Split by newlines
        lines = pasted_text.split('\n')
        added_count = 0
        skipped_count = 0
        
        for line in lines:
            parsed = self.parse_address_line(line)
            
            if parsed and (parsed['address'] or parsed['suburb']) and parsed['state']:
                address = parsed['address']
                suburb = parsed['suburb']
                state = parsed['state']
                
                # Check for duplicates
                is_duplicate = any(
                    addr['address'] == address and 
                    addr['suburb'] == suburb and 
                    addr['state'] == state
                    for addr in self.site_addresses
                )
                
                if not is_duplicate:
                    # Check if we have cached geocoding data
                    cache_key = f"{address}, {suburb}, {state}".lower()
                    is_cached = cache_key in self.geocode_cache
                    
                    status = '💾 Cached' if is_cached else 'Pending'
                    tag = 'cached' if is_cached else 'pending'
                    
                    # Add to tree
                    self.input_tree.insert('', tk.END, 
                                         values=(status, address, suburb, state),
                                         tags=(tag,))
                    
                    # Add to list
                    site_data = {
                        'address': address,
                        'suburb': suburb,
                        'state': state,
                        'status': 'pending'
                    }
                    
                    # If cached, add the cached data
                    if is_cached:
                        cached_data = self.geocode_cache[cache_key]
                        site_data.update(cached_data)
                        site_data['status'] = 'cached'
                    
                    self.site_addresses.append(site_data)
                    added_count += 1
                else:
                    skipped_count += 1
        
        # Clear paste area
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
        """Remove selected address from the list"""
        selected_items = self.input_tree.selection()
        if not selected_items:
            messagebox.showinfo("No Selection", "Please select an address to remove")
            return
        
        removed_count = 0
        for item in selected_items:
            values = self.input_tree.item(item, 'values')
            # Remove from list
            self.site_addresses = [addr for addr in self.site_addresses 
                                   if not (addr['address'] == values[1] and 
                                          addr['suburb'] == values[2] and 
                                          addr['state'] == values[3])]
            # Remove from tree
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
    
    def geocode_address(self, address, retry_broader=False):
        """Geocode an address using Nominatim API with retry logic"""
        base_url = "https://nominatim.openstreetmap.org/search"
        
        # If retry, make the search broader by removing specific details
        search_address = address
        if retry_broader:
            # Remove street numbers and unit details for broader search
            parts = address.split(',')
            if len(parts) >= 2:
                # Keep suburb and state, remove specific address
                search_address = ', '.join(parts[1:])
        
        params = {
            'q': search_address,
            'format': 'json',
            'limit': 1,
            'countrycodes': 'au'  # Focus on Australia for better results
        }
        headers = {
            'User-Agent': 'AddressDistanceCalculator/2.0'
        }
        
        try:
            response = requests.get(base_url, params=params, headers=headers)
            response.raise_for_status()
            data = response.json()
            
            if data:
                return float(data[0]['lat']), float(data[0]['lon']), retry_broader
            return None, None, retry_broader
        except Exception as e:
            print(f"Error geocoding address: {e}")
            return None, None, retry_broader
    
    def update_input_tree_status(self, address, suburb, state, status, tag):
        """Update the status in the input tree"""
        for item in self.input_tree.get_children():
            values = self.input_tree.item(item, 'values')
            if values[1] == address and values[2] == suburb and values[3] == state:
                self.input_tree.item(item, values=(status, address, suburb, state), tags=(tag,))
                break
    
    def calculate_distances(self):
        """Calculate distances from technician to all sites with smart retry and caching"""
        tech_addr = self.tech_address.get("1.0", tk.END).strip()
        
        if not tech_addr:
            messagebox.showwarning("Input Required", "Please enter the technician's address")
            return
        
        if not self.site_addresses:
            messagebox.showwarning("Input Required", "Please add at least one site address")
            return
        
        # Clear previous results
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        # Count addresses that need geocoding (those without lat/lon data)
        addresses_to_geocode = [s for s in self.site_addresses 
                                if 'lat' not in s or 'lon' not in s or not s.get('lat') or not s.get('lon')]
        cached_count = len(self.site_addresses) - len(addresses_to_geocode)
        
        # Show and initialize progress bar
        self.progress.grid()
        total_steps = len(addresses_to_geocode) + 1  # +1 for technician address
        self.progress['maximum'] = total_steps
        self.progress['value'] = 0
        
        self.status_var.set("🔍 Geocoding technician address...")
        self.root.update()
        
        # Geocode technician address
        tech_lat, tech_lon, _ = self.geocode_address(tech_addr)
        if not tech_lat:
            self.progress.grid_remove()
            messagebox.showerror("Geocoding Error", 
                               "Could not find technician address.\n\n" +
                               "Please check the address and try again.")
            self.status_var.set("✗ Error: Could not geocode technician address")
            return
        
        # Update progress for technician address
        self.progress['value'] = 1
        self.root.update()
        
        tech_coords = (tech_lat, tech_lon)
        time.sleep(1)  # Respect API rate limit
        
        results = []
        total = len(self.site_addresses)
        processed = 0
        
        for i, site in enumerate(self.site_addresses):
            # Update progress bar
            current_step = i + 2 - cached_count if i >= cached_count else 1
            if 'lat' not in site or 'lon' not in site:
                self.progress['value'] = current_step
            
            progress_percent = int(((i + 1) / total) * 100)
            
            # Build full address
            full_address = f"{site['address']}, {site['suburb']}, {site['state']}"
            cache_key = full_address.lower()
            
            # Check if already cached (either from previous calculation or from cache)
            if 'lat' in site and 'lon' in site and site['lat'] and site['lon']:
                # Use cached coordinates
                site_lat = site['lat']
                site_lon = site['lon']
                was_retried = site.get('was_retried', False)
                
                self.status_var.set(f"💾 Using cached data ({progress_percent}%): {site['suburb']}")
                self.root.update()
                
                site_coords = (site_lat, site_lon)
                distance_km = geodesic(tech_coords, site_coords).kilometers
                duration_min = (distance_km / 50) * 60
                
                status = "💾 Cached" if not was_retried else "💾 Cached (Broad)"
                tag = 'cached'
                
                results.append({
                    'address': site['address'],
                    'suburb': site['suburb'],
                    'state': site['state'],
                    'distance': distance_km,
                    'duration': duration_min,
                    'status': status,
                    'tag': tag
                })
                
                self.update_input_tree_status(site['address'], site['suburb'], 
                                             site['state'], status, tag)
                continue
            
            # Need to geocode this address
            processed += 1
            progress = f"⏳ Processing {processed}/{len(addresses_to_geocode)} ({progress_percent}%): {site['suburb']}"
            self.status_var.set(progress)
            self.root.update()
            
            # Try geocoding with original address
            site_lat, site_lon, was_retried = self.geocode_address(full_address, False)
            
            # If failed, try broader search
            if not site_lat:
                time.sleep(1)
                self.status_var.set(f"⚠ Retrying with broader search ({progress_percent}%): {site['suburb']}")
                self.root.update()
                site_lat, site_lon, was_retried = self.geocode_address(full_address, True)
            
            if site_lat and site_lon:
                site_coords = (site_lat, site_lon)
                # Calculate distance
                distance_km = geodesic(tech_coords, site_coords).kilometers
                
                # Estimate duration (50 km/h average)
                duration_min = (distance_km / 50) * 60
                
                # Cache the geocoding result
                self.geocode_cache[cache_key] = {
                    'lat': site_lat,
                    'lon': site_lon,
                    'was_retried': was_retried
                }
                
                # Update site data with geocoding info
                site['lat'] = site_lat
                site['lon'] = site_lon
                site['was_retried'] = was_retried
                site['status'] = 'geocoded'
                
                status = "✓ Found" if not was_retried else "⚠ Broad Match"
                tag = 'normal' if not was_retried else 'retried'
                
                results.append({
                    'address': site['address'],
                    'suburb': site['suburb'],
                    'state': site['state'],
                    'distance': distance_km,
                    'duration': duration_min,
                    'status': status,
                    'tag': tag
                })
                
                self.update_input_tree_status(site['address'], site['suburb'], 
                                             site['state'], status, tag)
            else:
                results.append({
                    'address': site['address'],
                    'suburb': site['suburb'],
                    'state': site['state'],
                    'distance': float('inf'),
                    'duration': float('inf'),
                    'status': '✗ Not Found',
                    'tag': 'error'
                })
                
                self.update_input_tree_status(site['address'], site['suburb'], 
                                             site['state'], '✗ Not Found', 'error')
            
            time.sleep(1)  # Respect API rate limit
        
        # Complete progress bar
        self.progress['value'] = total_steps
        self.root.update()
        time.sleep(0.3)  # Brief pause to show completion
        
        # Sort by distance
        results.sort(key=lambda x: x['distance'])
        
        # Display results
        success_count = 0
        retry_count = 0
        error_count = 0
        cached_result_count = 0
        
        for rank, result in enumerate(results, 1):
            if result['distance'] == float('inf'):
                self.results_tree.insert('', tk.END, values=(
                    rank,
                    result['address'],
                    result['suburb'],
                    result['state'],
                    'N/A',
                    'N/A',
                    result['status']
                ), tags=('error',))
                error_count += 1
            else:
                if result['tag'] == 'cached':
                    tag = 'cached'
                    cached_result_count += 1
                elif result['tag'] == 'normal':
                    tag = 'success'
                    success_count += 1
                else:
                    tag = 'warning'
                    retry_count += 1
                    
                self.results_tree.insert('', tk.END, values=(
                    rank,
                    result['address'],
                    result['suburb'],
                    result['state'],
                    f"{result['distance']:.2f}",
                    f"{result['duration']:.0f}",
                    result['status']
                ), tags=(tag,))
        
        # Hide progress bar
        self.progress.grid_remove()
        
        summary_parts = []
        if success_count > 0:
            summary_parts.append(f"{success_count} found")
        if cached_result_count > 0:
            summary_parts.append(f"{cached_result_count} cached")
        if retry_count > 0:
            summary_parts.append(f"{retry_count} broad matches")
        if error_count > 0:
            summary_parts.append(f"{error_count} not found")
        
        summary = f"✓ Complete! " + ", ".join(summary_parts)
        self.status_var.set(summary)
        
        if error_count > 0 or cached_result_count > 0:
            msg_parts = ["Calculation complete!\n"]
            if success_count > 0:
                msg_parts.append(f"✓ Successfully geocoded: {success_count}")
            if cached_result_count > 0:
                msg_parts.append(f"💾 Used cached data: {cached_result_count}")
            if retry_count > 0:
                msg_parts.append(f"⚠ Broad matches: {retry_count}")
            if error_count > 0:
                msg_parts.append(f"✗ Not found: {error_count}")
            msg_parts.append("\nCheck the Status column for details.")
            
            messagebox.showinfo("Results Summary", "\n".join(msg_parts))

if __name__ == "__main__":
    root = tk.Tk()
    app = AddressDistanceCalculator(root)
    root.mainloop()