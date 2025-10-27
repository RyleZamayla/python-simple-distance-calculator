# 📍 Address Distance Calculator

A modern, user-friendly desktop application for calculating distances between a technician's base location and multiple site addresses. Built with Python and Tkinter, featuring a sleek glass-morphism UI design and intelligent geocoding with caching.

![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)

## ✨ Features

### 🎨 Modern Glass Morphism Design
- Sleek dark theme with glass-effect panels
- Professional color scheme with indigo accents
- Smooth hover effects and modern typography
- Responsive layout that adapts to window resizing

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
- Built-in caching system - previously geocoded addresses are processed instantly
- Visual indicators for cached vs. newly geocoded addresses
- Respects API rate limits automatically

### 📊 Comprehensive Results
- Distance calculation in kilometers
- Estimated travel duration (based on 50 km/h average)
- Ranked results sorted by distance
- Color-coded status indicators:
  - 🟢 **Green** - Successfully geocoded
  - 🔵 **Blue** - Cached (instant results)
  - 🟡 **Yellow** - Broad match (approximate location)
  - 🔴 **Red** - Address not found

### 🛠️ Additional Features
- Duplicate address detection
- Bulk address management (add, remove, clear all)
- Real-time progress tracking
- Export-ready results table
- Responsive status messages

## 🚀 Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package installer)

### Required Dependencies
```bash
pip install requests geopy
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
   - Select and remove individual addresses with "❌ Remove Selected"
   - Clear all addresses with "🗑️ Clear All"
   - Previously calculated addresses use cached data for instant recalculation

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
- **Tkinter** - Native Python GUI framework
- **geopy** - Geocoding and distance calculations
- **requests** - HTTP client for API calls
- **OpenStreetMap Nominatim** - Free geocoding service

## 🎨 UI Design Philosophy

The application uses a modern glass-morphism design with:
- Dark gradient background (#1a1a2e)
- White glass-effect panels
- Indigo accent colors (#6366f1)
- Professional Segoe UI typography
- Smooth hover transitions
- Color-coded status indicators

## 🔧 Configuration

### Geocoding Settings
The application uses OpenStreetMap's Nominatim API with these defaults:
- **Country**: Australia (countrycodes: 'au')
- **Rate Limit**: 1 second between requests
- **Retry Logic**: Automatic broader search on failure

### Distance Calculation
- **Method**: Geodesic distance (great-circle distance)
- **Speed Estimate**: 50 km/h average for duration calculation

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
- Modern Python with type hints ready
- Comprehensive error handling
- Clean separation of concerns
- Extensive inline documentation
- User-friendly status messages

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
- **geopy** - For distance calculation utilities
- **Python Community** - For excellent documentation and support

## 🐛 Known Limitations

- Geocoding depends on OpenStreetMap data quality
- Rate limited to 1 request per second
- Designed primarily for Australian addresses (can be configured for other countries)
- Requires internet connection for geocoding

## 🔮 Future Enhancements

- [ ] Export results to CSV/Excel
- [ ] Import addresses from files
- [ ] Customizable speed estimates
- [ ] Map visualization
- [ ] Route optimization
- [ ] Support for multiple technicians
- [ ] Persistent cache storage
- [ ] Configuration file support
- [ ] Dark/light theme toggle

## 📧 Contact

For questions, suggestions, or issues, please open an issue on GitHub.

---

**Made with ❤️ using Python and Tkinter**
