# 📍 Address Distance Calculator

A modern, user-friendly desktop application for calculating distances between a technician's base location and multiple site addresses. Built with Python and Tkinter, featuring a sleek glass-morphism UI design and intelligent geocoding with caching.

![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)

## ✨ Features

### 🎨 Modern CustomTkinter Design
- Sleek dark/light theme toggle with glass-effect panels
- Professional color scheme with modern blue accents
- Smooth hover effects and modern typography
- Responsive layout that adapts to window resizing
- Built with CustomTkinter for enhanced UI components

### 📋 Flexible Address Input
Support for multiple paste formats:
- **Excel columns** (tab-separated): `Address | Suburb | State`
- **Comma-separated**: `Street Address, Suburb, State`
- **Full addresses**: `123 Main St, Sydney, NSW 2000`
- **Pipe-separated**: `Address | Suburb | State`
- **Line-by-line** in any of the above formats
- **Mixed formats** in a single paste operation

### ⚡ Smart Geocoding with Caching
- Intelligent address geocoding using OpenStreetMap Nominatim API
- Automatic retry with broader search if exact match fails
- **Persistent cache storage** - geocoded addresses saved between sessions
- Built-in caching system - previously geocoded addresses are processed instantly
- Visual indicators for cached vs. newly geocoded addresses
- Respects API rate limits automatically
- Cache stored in user's home directory: `~/.address_distance_cache.json`

### 📊 Comprehensive Results
- **Actual route distances** using OSRM routing engine
- **Accurate travel duration** based on real road networks
- Ranked results sorted by distance
- **Excel-like table** with selectable cells, rows, and columns
- **Smart copy** - Copy selected cells or entire results to clipboard
- **Filtering options** - Show/hide results by status type
- Color-coded status indicators:
  - 🟢 **Green** - Successfully geocoded
  - 🔵 **Blue** - Cached (instant results)
  - 🟡 **Yellow** - Broad match (approximate location)
  - 🔴 **Red** - Address not found

### 🛠️ Additional Features
- **Dark/Light theme toggle** - Switch between themes with a single click
- **Excel-like cell selection** - Select individual cells, rows, or columns in results
- **Smart copy functionality** - Copy selected cells or all results with Ctrl+C
- **Keyboard shortcuts** - Ctrl+A (select all), Ctrl+C (copy), Esc (clear selection)
- **Result filtering** - Filter results by status (Found, Cached, Broad, Not Found)
- **Auto-formatting** - Technician address auto-formats on paste
- **Real-time routing** - Uses OSRM for accurate route distances and durations
- **Duration formatting** - Smart formatting (e.g., "1 hr 30 min" or "45 min")
- **Threaded calculations** - Non-blocking UI during geocoding operations
- **Duplicate address detection**
- **Bulk address management** (add, remove, clear all)
- **Real-time progress tracking**
- **Context menu** - Right-click for quick actions on results
- **Responsive status messages**

## 🚀 Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package installer)

### Required Dependencies
```bash
pip install requests geopy customtkinter
```

### Clone the Repository
```bash
git clone https://github.com/YOUR_USERNAME/python-simple-distance-calculator.git
cd python-simple-distance-calculator
```

### Install Dependencies
```bash
pip install -r requirements.txt
```

## 💻 Usage

### Running the Application
```bash
python main.py
```

### Quick Start Guide

1. **Enter Technician Address**
   - Type or paste the base address in the top field

2. **Add Site Addresses**
   - Paste addresses in the paste area (supports multiple formats)
   - Click "➕ Add Addresses" to process
   - Addresses are automatically parsed and added to the list

3. **Calculate Distances**
   - Click "🚀 Calculate Distances"
   - Watch the progress bar as addresses are geocoded
   - View ranked results sorted by distance

4. **Manage Addresses**
   - Select and remove individual addresses with "✖ Remove"
   - Clear all addresses with "🗑 Clear All"
   - Previously calculated addresses use cached data for instant recalculation

5. **Work with Results**
   - Click on individual cells, drag to select ranges, or click column headers
   - Use Ctrl+A to select all results
   - Press Ctrl+C to copy selected cells or all results
   - Right-click for context menu with copy options
   - Filter results by clicking checkboxes (Found, Cached, Broad, Not Found)

6. **Toggle Theme**
   - Click the theme toggle button in the top-right corner
   - Switch between light and dark modes for comfortable viewing

### Example Input Formats

**Excel (tab-separated):**
```
123 Main St    Sydney    NSW
456 High St    Melbourne    VIC
```

**Comma-separated:**
```
123 Main St, Sydney, NSW
456 High St, Melbourne, VIC
```

**Full address:**
```
123 Main St, Sydney, NSW 2000
456 High St, Melbourne, VIC 3000
```

**Mixed formats:**
```
123 Main St    Sydney    NSW
456 High St, Melbourne, VIC
789 Low St | Brisbane | QLD
```

## 🏗️ Architecture

### Core Components
- **AddressDistanceCalculator** - Main application class
- **Geocoding Engine** - OpenStreetMap Nominatim integration
- **Cache System** - In-memory geocoding cache for performance
- **UI Framework** - Modern Tkinter with custom styling

### Key Technologies
- **CustomTkinter** - Modern Python GUI framework with native look
- **geopy** - Geocoding and distance calculations
- **requests** - HTTP client for API calls
- **OpenStreetMap Nominatim** - Free geocoding service
- **OSRM** - Open Source Routing Machine for accurate route calculations
- **Threading** - Non-blocking UI during long operations

## 🎨 UI Design Philosophy

The application uses a modern CustomTkinter design with:
- Dark/light theme support with smooth transitions
- Glass-effect panels with semi-transparent backgrounds
- Modern blue accent colors
- Professional Segoe UI typography
- Smooth hover transitions and animations
- Color-coded status indicators
- Excel-like table with cell selection capabilities

## 🔧 Configuration

### Geocoding Settings
The application uses OpenStreetMap's Nominatim API with these defaults:
- **Country**: Australia (countrycodes: 'au')
- **Rate Limit**: 1 second between requests
- **Retry Logic**: Automatic broader search on failure

### Distance Calculation
- **Method**: OSRM (Open Source Routing Machine) for actual route distances
- **Fallback**: Geodesic distance (great-circle distance) if OSRM unavailable
- **Duration**: Calculated from actual route data via OSRM API

## 📝 Development

### Project Structure
```
python-simple-distance-calculator/
├── main.py              # Main application file
├── requirements.txt     # Python dependencies
├── .gitignore          # Git ignore rules
├── README.md           # This file
└── build/              # PyInstaller build artifacts (ignored)
```

### Code Highlights
- Modern Python with threading for non-blocking operations
- CustomTkinter for enhanced UI components
- Comprehensive error handling
- Clean separation of concerns
- Queue-based thread communication
- Extensive inline documentation
- User-friendly status messages
- Excel-like table interaction model

## 🤝 Contributing

Contributions are welcome! Here are some ways you can contribute:
- Report bugs and issues
- Suggest new features
- Submit pull requests
- Improve documentation
- Share the project

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- **OpenStreetMap** - For providing free geocoding services
- **OSRM** - For providing free routing and distance calculations
- **CustomTkinter** - For modern UI components
- **geopy** - For distance calculation utilities
- **Python Community** - For excellent documentation and support

## 🐛 Known Limitations

- Geocoding depends on OpenStreetMap data quality
- Rate limited to 1 request per second
- Designed primarily for Australian addresses (can be configured for other countries)
- Requires internet connection for geocoding

## 🔮 Future Enhancements

- [ ] Export results to CSV/Excel files
- [ ] Import addresses from CSV/Excel files
- [ ] Customizable speed estimates for manual calculation mode
- [ ] Interactive map visualization of routes
- [ ] Route optimization (traveling salesman problem)
- [ ] Support for multiple technicians
- [x] Persistent cache storage
- [ ] Configuration file support for settings
- [x] Dark/light theme toggle
- [x] Excel-like cell selection and copying
- [x] Result filtering by status
- [x] OSRM integration for accurate routing
- [x] Threaded calculations
- [x] Smart duration formatting
- [x] Context menu for results

## 📧 Contact

For questions, suggestions, or issues, please open an issue on GitHub.

---

**Made with ❤️ using Python and Tkinter**
