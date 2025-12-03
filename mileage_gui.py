#!/usr/bin/env python3
"""
Mileage Analysis Tool - Desktop GUI
D'Ewart Representatives, L.L.C.

A PyQt6 desktop application with embedded Google Maps for analyzing
and categorizing business vs personal mileage.
"""

import sys
import os
import json
import urllib.parse
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QTableWidget, QTableWidgetItem, QPushButton, QLabel,
    QFileDialog, QComboBox, QDateEdit, QGroupBox, QFormLayout,
    QTabWidget, QTextEdit, QMessageBox, QProgressBar, QStatusBar,
    QHeaderView, QMenu, QLineEdit, QCheckBox, QFrame, QStyle
)
from PyQt6.QtCore import Qt, QDate, QThread, pyqtSignal, QUrl
from PyQt6.QtGui import QAction, QColor, QFont, QIcon
from PyQt6.QtWebEngineWidgets import QWebEngineView

# Import the analysis module
import analyze_mileage as analyzer


def get_app_dir():
    """Get the application directory (works for both script and exe)"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(__file__)


class AnalysisWorker(QThread):
    """Background worker for running mileage analysis"""
    finished = pyqtSignal(dict)
    progress = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, file_path: str, enable_lookup: bool = False,
                 start_date: str = None, end_date: str = None):
        super().__init__()
        self.file_path = file_path
        self.enable_lookup = enable_lookup
        self.start_date = start_date
        self.end_date = end_date

    def run(self):
        # Redirect stdout/stderr to prevent issues when no console (pythonw.exe)
        import io
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        if sys.stdout is None:
            sys.stdout = io.StringIO()
        if sys.stderr is None:
            sys.stderr = io.StringIO()

        try:
            self.progress.emit(f"Loading: {self.file_path}")

            # Load configuration
            analyzer.load_config()

            # Read trips from file
            self.progress.emit("Reading trips...")
            trips = analyzer.read_trips(self.file_path)
            if not trips:
                self.error.emit(f"No trips found in file: {self.file_path}")
                return

            self.progress.emit(f"Loaded {len(trips)} trips. Categorizing...")

            # Apply date filtering if specified
            if self.start_date:
                start = datetime.strptime(self.start_date, '%Y-%m-%d')
                trips = [t for t in trips if t['started'] >= start]
            if self.end_date:
                end = datetime.strptime(self.end_date, '%Y-%m-%d')
                end = end.replace(hour=23, minute=59, second=59)
                trips = [t for t in trips if t['started'] <= end]

            # Categorize trips using the existing module's function
            categorized_trips = []
            for i, trip in enumerate(trips):
                if i % 50 == 0:
                    self.progress.emit(f"Processing trip {i+1} of {len(trips)}...")

                # categorize_trip returns (category, business_name) tuple
                category, business_name = analyzer.categorize_trip(
                    trip,
                    enable_lookup=self.enable_lookup
                )

                trip['auto_category'] = category  # lowercase: 'business', 'personal', 'commute'
                trip['computed_category'] = category.upper()  # uppercase for display
                trip['business_name'] = business_name or ''
                categorized_trips.append(trip)

            # Calculate statistics using same logic as analyze_mileage.py
            self.progress.emit("Calculating statistics...")

            from collections import defaultdict
            weekly_stats = defaultdict(lambda: {
                'commute': 0.0,
                'business': 0.0,
                'personal': 0.0,
                'total': 0.0,
                'portland_miles': 0.0,
                'spokane_miles': 0.0,
                'weekend_miles': 0.0,
                'trips': []
            })

            for trip in categorized_trips:
                week_key = analyzer.get_week_key(trip['started'])
                weekly_stats[week_key][trip['auto_category']] += trip['distance']
                weekly_stats[week_key]['total'] += trip['distance']
                weekly_stats[week_key]['trips'].append(trip)

                # Track Portland trips
                if analyzer.is_portland_area(trip['start_address']) or analyzer.is_portland_area(trip['end_address']):
                    weekly_stats[week_key]['portland_miles'] += trip['distance']

                # Track Spokane trips
                if analyzer.is_spokane_area(trip['start_address']) or analyzer.is_spokane_area(trip['end_address']):
                    weekly_stats[week_key]['spokane_miles'] += trip['distance']

                # Track weekend miles
                day_of_week = trip['started'].weekday()
                hour = trip['started'].hour
                is_weekend = (day_of_week == 4 and hour >= 17) or \
                             (day_of_week == 5) or \
                             (day_of_week == 6) or \
                             (day_of_week == 0 and hour < 6)
                if is_weekend:
                    weekly_stats[week_key]['weekend_miles'] += trip['distance']

            # Calculate totals
            total_commute = sum(stats['commute'] for stats in weekly_stats.values())
            total_business = sum(stats['business'] for stats in weekly_stats.values())
            total_personal = sum(stats['personal'] for stats in weekly_stats.values())
            total_all = sum(stats['total'] for stats in weekly_stats.values())

            result = {
                'trips': categorized_trips,
                'weekly_stats': dict(weekly_stats),
                'totals': {
                    'total_miles': total_all,
                    'business_miles': total_business,
                    'personal_miles': total_personal,
                    'commute_miles': total_commute,
                    'business_pct': (total_business / total_all * 100) if total_all > 0 else 0,
                    'personal_pct': ((total_personal + total_commute) / total_all * 100) if total_all > 0 else 0
                }
            }

            self.finished.emit(result)

        except Exception as e:
            import traceback
            self.error.emit(f"{str(e)}\n{traceback.format_exc()}")
        finally:
            # Restore stdout/stderr
            sys.stdout = old_stdout
            sys.stderr = old_stderr


class MapView(QWebEngineView):
    """Embedded Google Maps view for displaying trip locations"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.trips_data = []
        self.selected_trip = None
        self.api_key = self._load_api_key()
        self._load_base_map()

    def _load_api_key(self):
        """Load Google Maps API key from config.json"""
        config_file = os.path.join(get_app_dir(), 'config.json')
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get('google_places_api_key', '').strip()
            except:
                pass
        return ''

    def _load_base_map(self):
        """Load the base Google Maps HTML"""
        html = self._generate_map_html()
        self.setHtml(html)

    def _generate_map_html(self, center_lat=47.7511, center_lng=-122.2076, zoom=10):
        """Generate the Google Maps HTML with JavaScript API"""
        # Build the Maps API URL with key if available
        api_url = "https://maps.googleapis.com/maps/api/js?libraries=geometry,places&callback=initMap&loading=async"
        if self.api_key:
            api_url += f"&key={self.api_key}"

        html = f'''
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mileage Map</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        html, body {{ height: 100%; width: 100%; }}
        #map {{ height: 100%; width: 100%; }}
        .info-window {{
            font-family: Arial, sans-serif;
            max-width: 300px;
        }}
        .info-window h3 {{
            margin-bottom: 8px;
            color: #1a73e8;
        }}
        .info-window p {{
            margin: 4px 0;
            font-size: 13px;
        }}
        .info-window .category {{
            display: inline-block;
            padding: 2px 8px;
            border-radius: 4px;
            font-weight: bold;
            font-size: 11px;
        }}
        .info-window .business {{ background: #e8f5e9; color: #2e7d32; }}
        .info-window .personal {{ background: #fff3e0; color: #e65100; }}
        .info-window .commute {{ background: #e3f2fd; color: #1565c0; }}
        .legend {{
            background: white;
            padding: 10px;
            margin: 10px;
            border-radius: 4px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.3);
        }}
        .legend-item {{
            display: flex;
            align-items: center;
            margin: 4px 0;
        }}
        .legend-color {{
            width: 16px;
            height: 16px;
            border-radius: 50%;
            margin-right: 8px;
        }}
    </style>
    <script src="{api_url}" async defer></script>
    <script>
        let map;
        let markers = [];
        let polylines = [];
        let infoWindow;

        function initMap() {{
            map = new google.maps.Map(document.getElementById('map'), {{
                center: {{ lat: {center_lat}, lng: {center_lng} }},
                zoom: {zoom},
                styles: [
                    {{
                        featureType: "poi",
                        elementType: "labels",
                        stylers: [{{ visibility: "off" }}]
                    }}
                ]
            }});

            infoWindow = new google.maps.InfoWindow();

            // Add legend
            const legend = document.createElement('div');
            legend.className = 'legend';
            legend.innerHTML = `
                <div class="legend-item">
                    <div class="legend-color" style="background: #4CAF50;"></div>
                    <span>Business</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: #FF9800;"></div>
                    <span>Personal</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color" style="background: #2196F3;"></div>
                    <span>Commute</span>
                </div>
            `;
            map.controls[google.maps.ControlPosition.LEFT_BOTTOM].push(legend);
        }}

        function clearMap() {{
            markers.forEach(m => m.setMap(null));
            polylines.forEach(p => p.setMap(null));
            markers = [];
            polylines = [];
        }}

        function getCategoryColor(category) {{
            switch(category) {{
                case 'BUSINESS': return '#4CAF50';
                case 'PERSONAL': return '#FF9800';
                case 'COMMUTE': return '#2196F3';
                default: return '#9E9E9E';
            }}
        }}

        function addTrip(tripData) {{
            const color = getCategoryColor(tripData.category);

            // Add start marker (smaller, hollow)
            if (tripData.startLat && tripData.startLng) {{
                const startMarker = new google.maps.Marker({{
                    position: {{ lat: tripData.startLat, lng: tripData.startLng }},
                    map: map,
                    icon: {{
                        path: google.maps.SymbolPath.CIRCLE,
                        scale: 6,
                        fillColor: 'white',
                        fillOpacity: 1,
                        strokeColor: color,
                        strokeWeight: 2
                    }},
                    title: 'Start: ' + tripData.startAddress
                }});
                markers.push(startMarker);
            }}

            // Add end marker (larger, filled)
            if (tripData.endLat && tripData.endLng) {{
                const endMarker = new google.maps.Marker({{
                    position: {{ lat: tripData.endLat, lng: tripData.endLng }},
                    map: map,
                    icon: {{
                        path: google.maps.SymbolPath.CIRCLE,
                        scale: 8,
                        fillColor: color,
                        fillOpacity: 1,
                        strokeColor: 'white',
                        strokeWeight: 2
                    }},
                    title: tripData.businessName || tripData.endAddress
                }});

                endMarker.addListener('click', () => {{
                    const categoryClass = tripData.category.toLowerCase();
                    infoWindow.setContent(`
                        <div class="info-window">
                            <h3>${{tripData.businessName || 'Unknown Location'}}</h3>
                            <p><span class="category ${{categoryClass}}">${{tripData.category}}</span></p>
                            <p><strong>Date:</strong> ${{tripData.date}}</p>
                            <p><strong>Distance:</strong> ${{tripData.distance.toFixed(1)}} miles</p>
                            <p><strong>From:</strong> ${{tripData.startAddress}}</p>
                            <p><strong>To:</strong> ${{tripData.endAddress}}</p>
                        </div>
                    `);
                    infoWindow.open(map, endMarker);
                }});

                markers.push(endMarker);
            }}

            // Draw route line if we have both coordinates
            if (tripData.startLat && tripData.startLng && tripData.endLat && tripData.endLng) {{
                const path = [
                    {{ lat: tripData.startLat, lng: tripData.startLng }},
                    {{ lat: tripData.endLat, lng: tripData.endLng }}
                ];

                const polyline = new google.maps.Polyline({{
                    path: path,
                    geodesic: true,
                    strokeColor: color,
                    strokeOpacity: 0.6,
                    strokeWeight: 3
                }});
                polyline.setMap(map);
                polylines.push(polyline);
            }}
        }}

        async function showAddress(address) {{
            console.log('showAddress called with:', address);
            if (!map) {{
                console.error('Map not initialized!');
                return 'Map not initialized';
            }}
            try {{
                // Use new Places API (Place.searchByText)
                console.log('Searching for:', address);
                const {{ Place }} = await google.maps.importLibrary("places");
                const {{ places }} = await Place.searchByText({{
                    textQuery: address,
                    fields: ['displayName', 'location'],
                    maxResultCount: 1
                }});
                
                console.log('Search results:', places ? places.length : 0);
                
                if (places && places.length > 0) {{
                    const place = places[0];
                    console.log('Found:', place.displayName, place.location.toString());
                    
                    map.setCenter(place.location);
                    map.setZoom(16);

                    // Clear any existing search markers
                    if (window.searchMarker) {{
                        window.searchMarker.setMap(null);
                    }}
                    if (window.searchInfoWindow) {{
                        window.searchInfoWindow.close();
                    }}

                    window.searchMarker = new google.maps.Marker({{
                        position: place.location,
                        map: map,
                        animation: google.maps.Animation.DROP,
                        title: address
                    }});
                    
                    window.searchInfoWindow = new google.maps.InfoWindow({{
                        content: '<div style="padding:5px;"><b>' + address + '</b></div>'
                    }});
                    window.searchInfoWindow.open(map, window.searchMarker);
                    
                    console.log('Marker added successfully');
                    return 'Success';
                }} else {{
                    console.error('No results found');
                    return 'No results';
                }}
            }} catch (e) {{
                console.error('Error in showAddress:', e);
                return 'Error: ' + e.message;
            }}
        }}

        function showLocation(lat, lng, label) {{
            console.log('showLocation called:', lat, lng, label);
            if (!map) {{
                console.error('Map not initialized!');
                return 'Map not initialized';
            }}
            try {{
                const position = {{ lat: lat, lng: lng }};
                map.setCenter(position);
                map.setZoom(15);

                const marker = new google.maps.Marker({{
                    position: position,
                    map: map,
                    animation: google.maps.Animation.DROP,
                    title: label
                }});

                if (label) {{
                    const infoWin = new google.maps.InfoWindow({{
                        content: '<div style="font-weight:bold;">' + label + '</div>'
                    }});
                    infoWin.open(map, marker);
                }}

                console.log('Marker added at', lat, lng);
                return 'Location shown';
            }} catch (e) {{
                console.error('Error in showLocation:', e);
                return 'Error: ' + e.message;
            }}
        }}

        function fitBounds(trips) {{
            if (trips.length === 0) return;

            const bounds = new google.maps.LatLngBounds();
            trips.forEach(t => {{
                if (t.startLat && t.startLng) {{
                    bounds.extend({{ lat: t.startLat, lng: t.startLng }});
                }}
                if (t.endLat && t.endLng) {{
                    bounds.extend({{ lat: t.endLat, lng: t.endLng }});
                }}
            }});
            map.fitBounds(bounds);
        }}
    </script>
</head>
<body>
    <div id="map"></div>
</body>
</html>
'''
        return html

    def show_trips(self, trips: List[Dict]):
        """Display multiple trips on the map"""
        self.trips_data = trips

        # We need to geocode addresses to get coordinates
        # For now, use a placeholder - in production, you'd use the Google Geocoding API
        js_code = "clearMap();\n"

        for trip in trips[:100]:  # Limit to 100 trips for performance
            # Create a simplified trip data object for JavaScript
            trip_js = {
                'date': trip['started'].strftime('%Y-%m-%d %H:%M'),
                'category': trip.get('computed_category', 'PERSONAL'),
                'distance': trip.get('distance', 0),
                'startAddress': trip.get('start_address', ''),
                'endAddress': trip.get('end_address', ''),
                'businessName': trip.get('business_name', ''),
                'startLat': None,
                'startLng': None,
                'endLat': None,
                'endLng': None
            }

            # Convert to JSON for JavaScript
            trip_json = json.dumps(trip_js)
            js_code += f"addTrip({trip_json});\n"

        self.page().runJavaScript(js_code)

    def show_address(self, address: str):
        """Center map on a specific address"""
        escaped_address = address.replace("'", "\'")
        js = f"showAddress('{escaped_address}');"
        self.page().runJavaScript(js)

    def open_in_google_maps(self, address: str):
        """Open address in Google Maps (external browser)"""
        import webbrowser
        url = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(address)}"
        webbrowser.open(url)


class SummaryWidget(QWidget):
    """Widget displaying summary statistics"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # Totals group
        totals_group = QGroupBox("Mileage Totals")
        totals_layout = QFormLayout(totals_group)

        self.total_label = QLabel("0.0 miles")
        self.business_label = QLabel("0.0 miles (0%)")
        self.personal_label = QLabel("0.0 miles (0%)")
        self.commute_label = QLabel("0.0 miles")

        # Style the labels
        font = QFont()
        font.setPointSize(11)
        font.setBold(True)

        for label in [self.total_label, self.business_label,
                     self.personal_label, self.commute_label]:
            label.setFont(font)

        self.business_label.setStyleSheet("color: #2e7d32;")
        self.personal_label.setStyleSheet("color: #e65100;")
        self.commute_label.setStyleSheet("color: #1565c0;")

        totals_layout.addRow("Total Miles:", self.total_label)
        totals_layout.addRow("Business:", self.business_label)
        totals_layout.addRow("Personal:", self.personal_label)
        totals_layout.addRow("Commute:", self.commute_label)

        layout.addWidget(totals_group)

        # Trip counts group
        counts_group = QGroupBox("Trip Counts")
        counts_layout = QFormLayout(counts_group)

        self.total_trips_label = QLabel("0")
        self.business_trips_label = QLabel("0")
        self.personal_trips_label = QLabel("0")
        self.commute_trips_label = QLabel("0")

        counts_layout.addRow("Total Trips:", self.total_trips_label)
        counts_layout.addRow("Business Trips:", self.business_trips_label)
        counts_layout.addRow("Personal Trips:", self.personal_trips_label)
        counts_layout.addRow("Commute Trips:", self.commute_trips_label)

        layout.addWidget(counts_group)
        layout.addStretch()

    def update_stats(self, data: dict):
        """Update the summary display with new data"""
        totals = data.get('totals', {})
        trips = data.get('trips', [])

        self.total_label.setText(f"{totals.get('total_miles', 0):.1f} miles")
        self.business_label.setText(
            f"{totals.get('business_miles', 0):.1f} miles ({totals.get('business_pct', 0):.1f}%)"
        )
        self.personal_label.setText(
            f"{totals.get('personal_miles', 0):.1f} miles ({totals.get('personal_pct', 0):.1f}%)"
        )
        self.commute_label.setText(f"{totals.get('commute_miles', 0):.1f} miles")

        # Count trips by category
        business_count = sum(1 for t in trips if t.get('computed_category') == 'BUSINESS')
        personal_count = sum(1 for t in trips if t.get('computed_category') == 'PERSONAL')
        commute_count = sum(1 for t in trips if t.get('computed_category') == 'COMMUTE')

        self.total_trips_label.setText(str(len(trips)))
        self.business_trips_label.setText(str(business_count))
        self.personal_trips_label.setText(str(personal_count))
        self.commute_trips_label.setText(str(commute_count))


class UnresolvedAddressesWidget(QWidget):
    """Widget for viewing and resolving unresolved business addresses"""

    address_selected = pyqtSignal(str, float, float)  # Emits address, lat, lng when clicked
    mapping_saved = pyqtSignal()  # Emits when a mapping is saved

    def __init__(self, parent=None):
        super().__init__(parent)
        self.addresses_data = []  # List of dicts with address info
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Header with count
        header_layout = QHBoxLayout()
        self.count_label = QLabel("No unresolved addresses")
        self.count_label.setStyleSheet("font-weight: bold; padding: 5px;")
        header_layout.addWidget(self.count_label)

        refresh_btn = QPushButton("Refresh")
        refresh_btn.setMaximumWidth(80)
        refresh_btn.clicked.connect(self._refresh_list)
        header_layout.addWidget(refresh_btn)

        layout.addLayout(header_layout)

        # Address list
        self.address_list = QTableWidget()
        self.address_list.setColumnCount(3)
        self.address_list.setHorizontalHeaderLabels(["Address", "Visits", "Total Miles"])
        self.address_list.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.address_list.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.address_list.setAlternatingRowColors(True)

        header = self.address_list.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.address_list.setColumnWidth(1, 60)
        self.address_list.setColumnWidth(2, 80)

        self.address_list.itemSelectionChanged.connect(self._on_address_selected)
        layout.addWidget(self.address_list)

        # Business name input area
        input_group = QGroupBox("Assign Business Name")
        input_layout = QVBoxLayout(input_group)

        self.selected_address_label = QLabel("Select an address above")
        self.selected_address_label.setWordWrap(True)
        self.selected_address_label.setStyleSheet("color: #666; padding: 5px; background: #f5f5f5;")
        input_layout.addWidget(self.selected_address_label)

        name_layout = QHBoxLayout()
        self.business_name_input = QLineEdit()
        self.business_name_input.setPlaceholderText("Enter business name...")
        self.business_name_input.returnPressed.connect(self._save_mapping)
        name_layout.addWidget(self.business_name_input)

        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self._save_mapping)
        self.save_btn.setEnabled(False)
        name_layout.addWidget(self.save_btn)

        self.personal_btn = QPushButton("Mark Personal")
        self.personal_btn.clicked.connect(self._mark_personal)
        self.personal_btn.setEnabled(False)
        name_layout.addWidget(self.personal_btn)

        input_layout.addLayout(name_layout)

        # Quick action buttons
        action_layout = QHBoxLayout()
        self.view_map_btn = QPushButton("View on Map")
        self.view_map_btn.clicked.connect(self._view_on_map)
        self.view_map_btn.setEnabled(False)
        action_layout.addWidget(self.view_map_btn)

        self.open_gmaps_btn = QPushButton("Open in Google Maps")
        self.open_gmaps_btn.clicked.connect(self._open_in_google_maps)
        self.open_gmaps_btn.setEnabled(False)
        action_layout.addWidget(self.open_gmaps_btn)

        action_layout.addStretch()
        input_layout.addLayout(action_layout)

        layout.addWidget(input_group)

        self.current_address = None

    def load_unresolved(self, trips: List[Dict]):
        """Analyze trips and find addresses without business names"""
        # Load existing business mappings
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        business_mapping = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    business_mapping = json.load(f)
            except:
                pass

        # Find unique destination addresses without business names
        address_stats = {}
        for trip in trips:
            addr = trip.get('end_address', '').strip()
            if not addr:
                continue

            # Skip if already has a business name or is in mapping
            business_name = trip.get('business_name', '')
            if business_name and business_name not in ['', 'Unknown', 'Home', 'Office']:
                continue

            # Skip home/work addresses
            if analyzer.is_home_address(addr) or analyzer.is_work_address(addr):
                continue

            # Skip if already mapped
            addr_lower = addr.lower()
            already_mapped = False
            for mapped_addr in business_mapping.keys():
                if mapped_addr.lower() in addr_lower or addr_lower in mapped_addr.lower():
                    already_mapped = True
                    break
            if already_mapped:
                continue

            # Track stats and coordinates
            if addr not in address_stats:
                address_stats[addr] = {
                    'visits': 0,
                    'miles': 0.0,
                    'lat': trip.get('end_lat'),
                    'lng': trip.get('end_lng')
                }
            address_stats[addr]['visits'] += 1
            address_stats[addr]['miles'] += trip.get('distance', 0)
            # Capture lat/lng if not already set
            if not address_stats[addr].get('lat'):
                address_stats[addr]['lat'] = trip.get('end_lat')
                address_stats[addr]['lng'] = trip.get('end_lng')

        # Sort by visit count (most visited first)
        self.addresses_data = [
            {'address': addr, 'visits': stats['visits'], 'miles': stats['miles'],
             'lat': stats.get('lat'), 'lng': stats.get('lng')}
            for addr, stats in sorted(address_stats.items(), key=lambda x: -x[1]['visits'])
        ]

        self._populate_list()

    def _populate_list(self):
        """Populate the address list table"""
        self.address_list.setRowCount(len(self.addresses_data))

        for row, data in enumerate(self.addresses_data):
            self.address_list.setItem(row, 0, QTableWidgetItem(data['address']))
            self.address_list.setItem(row, 1, QTableWidgetItem(str(data['visits'])))
            self.address_list.setItem(row, 2, QTableWidgetItem(f"{data['miles']:.1f}"))

        count = len(self.addresses_data)
        self.count_label.setText(f"{count} unresolved address{'es' if count != 1 else ''}")

    def _on_address_selected(self):
        """Handle address selection"""
        rows = self.address_list.selectionModel().selectedRows()
        if rows and rows[0].row() < len(self.addresses_data):
            data = self.addresses_data[rows[0].row()]
            self.current_address = data['address']
            self.selected_address_label.setText(self.current_address)
            self.selected_address_label.setStyleSheet("color: #000; padding: 5px; background: #e3f2fd;")

            # Enable buttons
            self.save_btn.setEnabled(True)
            self.personal_btn.setEnabled(True)
            self.view_map_btn.setEnabled(True)
            self.open_gmaps_btn.setEnabled(True)

            # Emit signal to show on map with coordinates
            data = self.addresses_data[rows[0].row()]
            lat = data.get('lat') or 0.0
            lng = data.get('lng') or 0.0
            self.address_selected.emit(self.current_address, lat, lng)
        else:
            self.current_address = None
            self.selected_address_label.setText("Select an address above")
            self.selected_address_label.setStyleSheet("color: #666; padding: 5px; background: #f5f5f5;")
            self.save_btn.setEnabled(False)
            self.personal_btn.setEnabled(False)
            self.view_map_btn.setEnabled(False)
            self.open_gmaps_btn.setEnabled(False)

    def _save_mapping(self):
        """Save the business name mapping"""
        if not self.current_address:
            return

        name = self.business_name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "No Name", "Please enter a business name.")
            return

        self._save_to_mapping_file(self.current_address, name)
        self.business_name_input.clear()
        self._remove_current_from_list()
        self.mapping_saved.emit()

    def _mark_personal(self):
        """Mark the address as personal"""
        if not self.current_address:
            return

        self._save_to_mapping_file(self.current_address, "[PERSONAL]")
        self._remove_current_from_list()
        self.mapping_saved.emit()

    def _save_to_mapping_file(self, address: str, name: str):
        """Save a mapping to the business_mapping.json file"""
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')

        # Load existing
        mappings = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except:
                pass

        # Add new mapping
        mappings[address] = name

        # Save
        try:
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, indent=2)
            QMessageBox.information(self, "Saved", f"Saved: {address[:50]}... = {name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save: {e}")

    def _remove_current_from_list(self):
        """Remove the current address from the list"""
        if self.current_address:
            self.addresses_data = [d for d in self.addresses_data if d['address'] != self.current_address]
            self._populate_list()
            self.current_address = None
            self._on_address_selected()

    def _view_on_map(self):
        """Emit signal to view address on embedded map"""
        if self.current_address:
            # Find the data for current address
            lat, lng = 0.0, 0.0
            for data in self.addresses_data:
                if data['address'] == self.current_address:
                    lat = data.get('lat') or 0.0
                    lng = data.get('lng') or 0.0
                    break
            self.address_selected.emit(self.current_address, lat, lng)

    def _open_in_google_maps(self):
        """Open address in external Google Maps"""
        if self.current_address:
            import webbrowser
            url = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(self.current_address)}"
            webbrowser.open(url)

    def _refresh_list(self):
        """Signal to parent to refresh the list"""
        self.mapping_saved.emit()


class TripTableWidget(QTableWidget):
    """Table widget for displaying trip data"""

    trip_selected = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
        self.trips_data = []

    def _setup_ui(self):
        self.setColumnCount(7)
        self.setHorizontalHeaderLabels([
            "Date", "Day", "Category", "Distance",
            "From", "To", "Business Name"
        ])

        # Set column widths
        header = self.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)

        self.setColumnWidth(0, 130)
        self.setColumnWidth(1, 60)
        self.setColumnWidth(2, 80)
        self.setColumnWidth(3, 70)

        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)

        # Context menu
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_context_menu)

        # Selection changed
        self.itemSelectionChanged.connect(self._on_selection_changed)

    def _show_context_menu(self, pos):
        """Show context menu on right-click"""
        row = self.rowAt(pos.y())
        if row < 0 or row >= len(self.trips_data):
            return

        menu = QMenu(self)

        view_map_action = menu.addAction("View on Map")
        open_gmaps_action = menu.addAction("Open in Google Maps")
        menu.addSeparator()
        change_category = menu.addMenu("Change Category")
        business_action = change_category.addAction("Business")
        personal_action = change_category.addAction("Personal")
        commute_action = change_category.addAction("Commute")

        action = menu.exec(self.viewport().mapToGlobal(pos))

        trip = self.trips_data[row]
        if action == view_map_action:
            self.trip_selected.emit(trip)
        elif action == open_gmaps_action:
            import webbrowser
            addr = trip.get('end_address', '')
            url = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(addr)}"
            webbrowser.open(url)

    def _on_selection_changed(self):
        """Handle selection change"""
        rows = self.selectionModel().selectedRows()
        if rows and len(self.trips_data) > rows[0].row():
            trip = self.trips_data[rows[0].row()]
            self.trip_selected.emit(trip)

    def load_trips(self, trips: List[Dict]):
        """Load trip data into the table"""
        self.trips_data = trips
        self.setRowCount(len(trips))

        for row, trip in enumerate(trips):
            # Date
            date_str = trip['started'].strftime('%Y-%m-%d %H:%M')
            self.setItem(row, 0, QTableWidgetItem(date_str))

            # Day of week
            day_str = trip['started'].strftime('%a')
            self.setItem(row, 1, QTableWidgetItem(day_str))

            # Category
            category = trip.get('computed_category', 'PERSONAL')
            cat_item = QTableWidgetItem(category)
            if category == 'BUSINESS':
                cat_item.setBackground(QColor('#e8f5e9'))
                cat_item.setForeground(QColor('#2e7d32'))
            elif category == 'PERSONAL':
                cat_item.setBackground(QColor('#fff3e0'))
                cat_item.setForeground(QColor('#e65100'))
            elif category == 'COMMUTE':
                cat_item.setBackground(QColor('#e3f2fd'))
                cat_item.setForeground(QColor('#1565c0'))
            self.setItem(row, 2, cat_item)

            # Distance
            dist_str = f"{trip.get('distance', 0):.1f} mi"
            self.setItem(row, 3, QTableWidgetItem(dist_str))

            # From/To addresses
            self.setItem(row, 4, QTableWidgetItem(trip.get('start_address', '')))
            self.setItem(row, 5, QTableWidgetItem(trip.get('end_address', '')))

            # Business name
            self.setItem(row, 6, QTableWidgetItem(trip.get('business_name', '')))

    def filter_by_category(self, category: str):
        """Show only trips of a specific category (or all)"""
        for row in range(self.rowCount()):
            if category == "All":
                self.setRowHidden(row, False)
            else:
                trip = self.trips_data[row]
                show = trip.get('computed_category', '') == category.upper()
                self.setRowHidden(row, not show)


class MileageAnalyzerGUI(QMainWindow):
    """Main application window"""

    def __init__(self):
        super().__init__()
        self.current_file = None
        self.analysis_data = None
        self._setup_ui()
        self._setup_menu()

    def _setup_ui(self):
        self.setWindowTitle("Mileage Analyzer - D'Ewart Representatives, L.L.C.")
        self.setMinimumSize(1200, 800)

        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(4, 4, 4, 4)
        main_layout.setSpacing(4)

        # Toolbar area
        toolbar = self._create_toolbar()
        main_layout.addWidget(toolbar)

        # Main content area with splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Left side: Tabs for Trips and Unresolved Addresses
        left_tabs = QTabWidget()

        # === Tab 1: All Trips ===
        trips_tab = QWidget()
        trips_layout = QVBoxLayout(trips_tab)
        trips_layout.setContentsMargins(0, 0, 0, 0)

        # Filter bar
        filter_frame = QFrame()
        filter_frame.setFrameStyle(QFrame.Shape.StyledPanel)
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(8, 4, 8, 4)

        filter_layout.addWidget(QLabel("Category:"))
        self.category_filter = QComboBox()
        self.category_filter.addItems(["All", "Business", "Personal", "Commute"])
        self.category_filter.currentTextChanged.connect(self._on_category_filter_changed)
        filter_layout.addWidget(self.category_filter)

        filter_layout.addWidget(QLabel("Search:"))
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search addresses...")
        self.search_box.textChanged.connect(self._on_search_changed)
        filter_layout.addWidget(self.search_box)

        filter_layout.addStretch()
        trips_layout.addWidget(filter_frame)

        # Trip table
        self.trip_table = TripTableWidget()
        self.trip_table.trip_selected.connect(self._on_trip_selected)
        trips_layout.addWidget(self.trip_table)

        left_tabs.addTab(trips_tab, "All Trips")

        # === Tab 2: Unresolved Addresses ===
        self.unresolved_widget = UnresolvedAddressesWidget()
        self.unresolved_widget.address_selected.connect(self._on_address_selected_for_map)
        self.unresolved_widget.mapping_saved.connect(self._on_mapping_saved)
        left_tabs.addTab(self.unresolved_widget, "Resolve Addresses")

        splitter.addWidget(left_tabs)
        self.left_tabs = left_tabs

        # Right side: Map and Summary tabs
        self.right_tabs = QTabWidget()

        # Map tab
        self.map_view = MapView()
        self.right_tabs.addTab(self.map_view, "Map")

        # Summary tab
        self.summary_widget = SummaryWidget()
        self.right_tabs.addTab(self.summary_widget, "Summary")

        # Weekly breakdown tab
        self.weekly_text = QTextEdit()
        self.weekly_text.setReadOnly(True)
        self.weekly_text.setFont(QFont("Consolas", 10))
        self.right_tabs.addTab(self.weekly_text, "Weekly Breakdown")

        splitter.addWidget(self.right_tabs)

        # Set splitter sizes (60% left, 40% right)
        splitter.setSizes([700, 500])

        main_layout.addWidget(splitter)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # Progress bar (hidden by default)
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(200)
        self.progress_bar.hide()
        self.status_bar.addPermanentWidget(self.progress_bar)

        self.status_bar.showMessage("Ready. Open a trip log file to begin.")

    def _create_toolbar(self):
        """Create the toolbar with main actions"""
        toolbar = QFrame()
        toolbar.setFrameStyle(QFrame.Shape.StyledPanel)
        toolbar.setFixedHeight(45)
        layout = QHBoxLayout(toolbar)
        layout.setContentsMargins(8, 4, 8, 4)

        # Open file button
        open_btn = QPushButton("Open Trip Log")
        open_btn.clicked.connect(self._open_file)
        layout.addWidget(open_btn)

        # Date range
        layout.addWidget(QLabel("From:"))
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate().addMonths(-3))
        layout.addWidget(self.start_date)

        layout.addWidget(QLabel("To:"))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())
        layout.addWidget(self.end_date)

        # Business lookup checkbox
        self.lookup_checkbox = QCheckBox("Enable Business Lookup")
        self.lookup_checkbox.setToolTip("Use online services to identify businesses (slower)")
        layout.addWidget(self.lookup_checkbox)

        # Analyze button
        analyze_btn = QPushButton("Analyze")
        analyze_btn.clicked.connect(self._run_analysis)
        layout.addWidget(analyze_btn)

        layout.addStretch()

        # Export button
        export_btn = QPushButton("Export to Excel")
        export_btn.clicked.connect(self._export_excel)
        layout.addWidget(export_btn)

        return toolbar

    def _setup_menu(self):
        """Set up the application menu bar"""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("&File")

        open_action = QAction("&Open Trip Log...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self._open_file)
        file_menu.addAction(open_action)

        file_menu.addSeparator()

        export_action = QAction("&Export to Excel...", self)
        export_action.setShortcut("Ctrl+E")
        export_action.triggered.connect(self._export_excel)
        file_menu.addAction(export_action)

        file_menu.addSeparator()

        exit_action = QAction("E&xit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # View menu
        view_menu = menubar.addMenu("&View")

        refresh_action = QAction("&Refresh Analysis", self)
        refresh_action.setShortcut("F5")
        refresh_action.triggered.connect(self._run_analysis)
        view_menu.addAction(refresh_action)

        # Tools menu
        tools_menu = menubar.addMenu("&Tools")

        unresolved_action = QAction("Analyze &Unresolved Addresses...", self)
        unresolved_action.triggered.connect(self._analyze_unresolved)
        tools_menu.addAction(unresolved_action)

        mapping_action = QAction("Edit Business &Mappings...", self)
        mapping_action.triggered.connect(self._edit_mappings)
        tools_menu.addAction(mapping_action)

        # Help menu
        help_menu = menubar.addMenu("&Help")

        about_action = QAction("&About", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    def _open_file(self):
        """Open a trip log file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Trip Log",
            "",
            "Trip Files (*.csv *.xlsx);;CSV Files (*.csv);;Excel Files (*.xlsx);;All Files (*.*)"
        )

        if file_path:
            self.current_file = file_path
            self.status_bar.showMessage(f"Loaded: {os.path.basename(file_path)}")
            self._run_analysis()

    def _run_analysis(self):
        """Run the mileage analysis"""
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please open a trip log file first.")
            return

        # Get date range
        start_date = self.start_date.date().toString("yyyy-MM-dd")
        end_date = self.end_date.date().toString("yyyy-MM-dd")

        # Show progress
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.progress_bar.show()

        # Run analysis in background thread
        self.worker = AnalysisWorker(
            self.current_file,
            enable_lookup=self.lookup_checkbox.isChecked(),
            start_date=start_date,
            end_date=end_date
        )
        self.worker.finished.connect(self._on_analysis_complete)
        self.worker.progress.connect(self._on_analysis_progress)
        self.worker.error.connect(self._on_analysis_error)
        self.worker.start()

    def _on_analysis_progress(self, message: str):
        """Update progress message"""
        self.status_bar.showMessage(message)

    def _on_analysis_complete(self, data: dict):
        """Handle analysis completion"""
        self.progress_bar.hide()
        self.analysis_data = data

        # Update trip table
        trips = data.get('trips', [])
        self.trip_table.load_trips(trips)

        # Update summary
        self.summary_widget.update_stats(data)

        # Update weekly breakdown text
        self._update_weekly_text(data)

        # Update map with trips
        self.map_view.show_trips(trips)

        # Update unresolved addresses list
        self.unresolved_widget.load_unresolved(trips)

        # Update tab label with count
        unresolved_count = len(self.unresolved_widget.addresses_data)
        if unresolved_count > 0:
            self.left_tabs.setTabText(1, f"Resolve Addresses ({unresolved_count})")
        else:
            self.left_tabs.setTabText(1, "Resolve Addresses")

        self.status_bar.showMessage(f"Analysis complete. {len(trips)} trips, {unresolved_count} addresses to resolve.")

    def _on_analysis_error(self, error: str):
        """Handle analysis error"""
        self.progress_bar.hide()
        QMessageBox.critical(self, "Analysis Error", f"Error during analysis:\n{error}")
        self.status_bar.showMessage("Analysis failed.")

    def _update_weekly_text(self, data: dict):
        """Update the weekly breakdown text view"""
        weekly_stats = data.get('weekly_stats', {})

        text = "WEEKLY MILEAGE BREAKDOWN\n"
        text += "=" * 80 + "\n\n"
        text += f"{'Week Starting':<15} {'Commute':>10} {'Business':>10} {'Personal':>10} {'Total':>10}\n"
        text += "-" * 55 + "\n"

        for week in sorted(weekly_stats.keys()):
            stats = weekly_stats[week]
            text += f"{week:<15} {stats.get('commute', 0):>10.1f} {stats.get('business', 0):>10.1f} "
            text += f"{stats.get('personal', 0):>10.1f} {stats.get('total', 0):>10.1f}\n"

        self.weekly_text.setText(text)

    def _on_category_filter_changed(self, category: str):
        """Handle category filter change"""
        self.trip_table.filter_by_category(category)

    def _on_search_changed(self, text: str):
        """Handle search text change"""
        text = text.lower()
        for row in range(self.trip_table.rowCount()):
            show = True
            if text:
                # Search in address columns
                from_addr = self.trip_table.item(row, 4).text().lower() if self.trip_table.item(row, 4) else ""
                to_addr = self.trip_table.item(row, 5).text().lower() if self.trip_table.item(row, 5) else ""
                business = self.trip_table.item(row, 6).text().lower() if self.trip_table.item(row, 6) else ""
                show = text in from_addr or text in to_addr or text in business
            self.trip_table.setRowHidden(row, not show)

    def _on_trip_selected(self, trip: dict):
        """Handle trip selection - show on map"""
        address = trip.get('end_address', '')
        if address:
            self.right_tabs.setCurrentIndex(0)  # Switch to Map tab
            self.map_view.show_address(address)

    def _on_address_selected_for_map(self, address: str, lat: float, lng: float):
        """Handle address selection from unresolved list - show on embedded map"""
        if address:
            self.right_tabs.setCurrentIndex(0)  # Switch to Map tab
            self.map_view.show_address(address)

    def _on_mapping_saved(self):
        """Handle when a business mapping is saved - refresh analysis"""
        if self.current_file:
            self._run_analysis()

    def _export_excel(self):
        """Export analysis to Excel"""
        if not self.analysis_data:
            QMessageBox.warning(self, "No Data", "Please run an analysis first.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export to Excel",
            "mileage_analysis.xlsx",
            "Excel Files (*.xlsx)"
        )

        if file_path:
            try:
                # Use the existing export function from analyze_mileage
                trips = self.analysis_data.get('trips', [])
                weekly_stats = self.analysis_data.get('weekly_stats', {})
                totals = self.analysis_data.get('totals', {})

                # export_to_excel expects individual totals, not a dict
                analyzer.export_to_excel(
                    trips,
                    weekly_stats,
                    totals.get('commute_miles', 0),
                    totals.get('business_miles', 0),
                    totals.get('personal_miles', 0),
                    totals.get('total_miles', 0),
                    filename=file_path
                )

                QMessageBox.information(
                    self, "Export Complete",
                    f"Analysis exported to:\n{file_path}"
                )
            except Exception as e:
                QMessageBox.critical(
                    self, "Export Error",
                    f"Failed to export:\n{str(e)}"
                )

    def _analyze_unresolved(self):
        """Open dialog to analyze unresolved addresses"""
        QMessageBox.information(
            self, "Unresolved Addresses",
            "This feature will help identify addresses that need business name mappings.\n\n"
            "For now, please use the command line:\n"
            "MileageAnalyzer.exe --analyze-unresolved"
        )

    def _edit_mappings(self):
        """Open business mappings file for editing"""
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        if os.path.exists(mapping_file):
            os.startfile(mapping_file)
        else:
            QMessageBox.warning(
                self, "File Not Found",
                "business_mapping.json not found.\n"
                "Run an analysis first to create this file."
            )

    def _show_about(self):
        """Show about dialog"""
        QMessageBox.about(
            self, "About Mileage Analyzer",
            "<h2>Mileage Analyzer</h2>"
            "<p>Version 3.0</p>"
            "<p>D'Ewart Representatives, L.L.C.</p>"
            "<hr>"
            "<p>A desktop application for analyzing and categorizing "
            "vehicle mileage for business expense tracking.</p>"
            "<p>Features:</p>"
            "<ul>"
            "<li>Automatic trip categorization</li>"
            "<li>Business location identification</li>"
            "<li>Interactive map view</li>"
            "<li>Excel report generation</li>"
            "</ul>"
        )


def main():
    """Application entry point"""
    app = QApplication(sys.argv)

    # Set application style
    app.setStyle('Fusion')

    # Create and show main window
    window = MileageAnalyzerGUI()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
