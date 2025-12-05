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
    QHeaderView, QMenu, QLineEdit, QCheckBox, QFrame, QStyle,
    QTreeWidget, QTreeWidgetItem, QAbstractItemView, QDialog,
    QScrollArea, QDoubleSpinBox
)
from PyQt6.QtCore import Qt, QDate, QThread, pyqtSignal, QUrl, QSettings, QByteArray
from PyQt6.QtGui import QAction, QColor, QFont, QIcon
from PyQt6.QtWebEngineWidgets import QWebEngineView

# Import the analysis module
import analyze_mileage as analyzer


class NumericTableWidgetItem(QTableWidgetItem):
    """QTableWidgetItem that sorts numerically instead of alphabetically"""
    def __init__(self, text, value=None):
        super().__init__(text)
        self._value = value if value is not None else 0
    
    def __lt__(self, other):
        if isinstance(other, NumericTableWidgetItem):
            return self._value < other._value
        return super().__lt__(other)


def get_app_dir():
    """Get the application directory (works for both script and exe)"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(__file__)


def get_trip_key(trip: dict) -> str:
    """Generate a unique key for a trip based on start time and start address"""
    started = trip.get('started')
    if hasattr(started, 'isoformat'):
        start_str = started.isoformat()
    else:
        start_str = str(started)
    return f"{start_str}|{trip.get('start_address', '')}"


def load_trip_notes() -> dict:
    """Load trip notes from JSON file"""
    notes_file = os.path.join(get_app_dir(), 'trip_notes.json')
    if os.path.exists(notes_file):
        try:
            with open(notes_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {}


def save_trip_note(trip: dict, note: str) -> tuple[bool, str]:
    """Save a note for a specific trip.
    Returns (success, message) tuple."""
    notes_file = os.path.join(get_app_dir(), 'trip_notes.json')
    notes = load_trip_notes()
    trip_key = get_trip_key(trip)

    if note.strip():
        notes[trip_key] = note.strip()
    elif trip_key in notes:
        del notes[trip_key]  # Remove empty notes

    try:
        with open(notes_file, 'w', encoding='utf-8') as f:
            json.dump(notes, f, indent=2, ensure_ascii=False)
        return True, "Note saved"
    except PermissionError:
        return False, "Cannot save note - file is in use by another program"
    except Exception as e:
        return False, f"Failed to save note: {e}"


def get_category_reason(trip: dict, category: str, business_name: str) -> str:
    """Determine why a trip was categorized a certain way"""
    end_addr = trip.get('end_address', '')
    start_addr = trip.get('start_address', '')
    distance = trip.get('distance', 0)
    timestamp = trip.get('started')

    # Check for saved category mapping first
    if analyzer.get_mapping_category(end_addr):
        return "Saved mapping"

    # Business location detection
    if analyzer.is_business_location(start_addr) or analyzer.is_business_location(end_addr):
        return "Known business location"

    # Home/work commute
    if analyzer.is_home_address(start_addr) and analyzer.is_work_address(end_addr):
        return "Home to office commute"
    if analyzer.is_work_address(start_addr) and analyzer.is_home_address(end_addr):
        return "Office to home commute"

    # Whidbey trips
    if analyzer.is_whidbey_area(start_addr) or analyzer.is_whidbey_area(end_addr):
        return "Whidbey area (personal)"

    # Distance threshold
    if distance >= analyzer.BUSINESS_DISTANCE_THRESHOLD:
        if timestamp:
            day_of_week = timestamp.weekday()
            if day_of_week < 5:
                return f"Weekday trip >= {analyzer.BUSINESS_DISTANCE_THRESHOLD} mi"

    # Weekend detection
    if timestamp:
        day_of_week = timestamp.weekday()
        hour = timestamp.hour
        if day_of_week == 4 and hour >= 17:
            return "Friday after 5pm (weekend)"
        if day_of_week in [5, 6]:
            return "Weekend day"
        if day_of_week == 0 and hour < 7:
            return "Monday before 7am (weekend)"

    # Local area detection
    if analyzer.is_bothell_area(start_addr) and analyzer.is_bothell_area(end_addr):
        if distance < analyzer.BUSINESS_DISTANCE_THRESHOLD:
            if timestamp and timestamp.weekday() < 5:
                return "Local weekday trip"
            else:
                return "Local weekend trip"

    return "Default personal"


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

            # Load configuration and business mapping
            analyzer.load_config()
            analyzer.load_business_mapping()

            # Read trips from file
            self.progress.emit("Reading trips...")
            trips = analyzer.read_trips(self.file_path)
            if not trips:
                self.error.emit(f"No trips found in file: {self.file_path}")
                return

            # Merge false stops (red lights, traffic stops, etc.)
            self.progress.emit(f"Loaded {len(trips)} trips. Merging false stops...")
            trips, merge_count = analyzer.merge_short_stops(trips)
            if merge_count > 0:
                self.progress.emit(f"Merged {merge_count} false stops. Now {len(trips)} trips.")

            # Flag micro-trips (GPS drift, parking adjustments)
            self.progress.emit("Flagging micro-trips...")
            trips, micro_count = analyzer.flag_micro_trips(trips)
            if micro_count > 0:
                self.progress.emit(f"Found {micro_count} micro-trips. Categorizing...")
            else:
                self.progress.emit(f"{len(trips)} trips ready. Categorizing...")

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
                trip['category_reason'] = get_category_reason(trip, category, business_name)
                categorized_trips.append(trip)

            # Save business mapping if lookup was enabled (to persist new lookups)
            if self.enable_lookup:
                analyzer.save_business_mapping()

            # Detect duplicate trips
            self.progress.emit("Checking for duplicates...")
            duplicate_count = 0
            seen_trips = {}  # Key: (start_time, end_address) -> trip
            for trip in categorized_trips:
                start_time = trip.get('started')
                end_addr = trip.get('end_address', '').strip().lower()
                if start_time:
                    key = (start_time.strftime('%Y-%m-%d %H:%M'), end_addr)
                    if key in seen_trips:
                        # Mark both as potential duplicates
                        trip['is_duplicate'] = True
                        seen_trips[key]['is_duplicate'] = True
                        duplicate_count += 1
                    else:
                        seen_trips[key] = trip

            if duplicate_count > 0:
                self.progress.emit(f"Found {duplicate_count} potential duplicate trips.")

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

            # Find date range of all trips (before filtering)
            all_dates = [t['started'] for t in categorized_trips if t.get('started')]
            min_date = min(all_dates).strftime('%Y-%m-%d') if all_dates else None
            max_date = max(all_dates).strftime('%Y-%m-%d') if all_dates else None

            result = {
                'trips': categorized_trips,
                'weekly_stats': dict(weekly_stats),
                'date_range': {'min': min_date, 'max': max_date},
                'totals': {
                    'total_miles': total_all,
                    'business_miles': total_business,
                    'personal_miles': total_personal,
                    'commute_miles': total_commute,
                    'business_pct': (total_business / total_all * 100) if total_all > 0 else 0,
                    'personal_pct': (total_personal / total_all * 100) if total_all > 0 else 0,
                    'commute_pct': (total_commute / total_all * 100) if total_all > 0 else 0
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

    business_selected = pyqtSignal(str)  # Emitted when user selects a business from map

    def __init__(self, parent=None):
        super().__init__(parent)
        self.trips_data = []
        self.selected_trip = None
        self.api_key = self._load_api_key()
        
        self._load_base_map()

        # Timer to poll for business selection from map
        from PyQt6.QtCore import QTimer
        self._business_poll_timer = QTimer()
        self._business_poll_timer.timeout.connect(self._check_selected_business)
        self._business_poll_timer.start(500)  # Check every 500ms
        
        # Timer to poll for placeId lookup requests from JavaScript
        self._placeid_poll_timer = QTimer()
        self._placeid_poll_timer.timeout.connect(self._check_placeid_request)
        self._placeid_poll_timer.start(200)  # Check every 200ms
        

    def _check_selected_business(self):
        """Poll JavaScript for selected business name"""
        self.page().runJavaScript(
            "if (window.selectedBusinessName) { var name = window.selectedBusinessName; window.selectedBusinessName = null; name; } else { null; }",
            self._handle_business_selection
        )

    def _handle_business_selection(self, business_name):
        """Handle business name selected from map"""
        if business_name:
            self.business_selected.emit(business_name)

    def _check_placeid_request(self):
        """Poll JavaScript for placeId lookup requests"""
        self.page().runJavaScript(
            "if (window.pendingPlaceIdRequest) { var req = window.pendingPlaceIdRequest; window.pendingPlaceIdRequest = null; req; } else { null; }",
            self._handle_placeid_request
        )
    
    def _handle_placeid_request(self, request):
        """Fetch place details from Python (avoids CORS issues)"""
        if not request:
            return
        
        import requests
        place_id = request.get('placeId')
        lat = request.get('lat')
        lng = request.get('lng')
        
        if not place_id or not self.api_key:
            return
        
        try:
            url = f"https://maps.googleapis.com/maps/api/place/details/json?place_id={place_id}&fields=name,formatted_address,vicinity&key={self.api_key}"
            response = requests.get(url, timeout=10)
            data = response.json()
            
            if data.get('status') == 'OK' and data.get('result'):
                place = data['result']
                name = place.get('name', '')
                address = place.get('vicinity') or place.get('formatted_address', '')
                # Send result back to JavaScript
                js = f"window.placeDetailsResult = {{ name: '{self._js_escape(name)}', address: '{self._js_escape(address)}', lat: {lat}, lng: {lng} }};"
                self.page().runJavaScript(js)
            else:
                # Send error back
                error_msg = data.get('error_message', data.get('status', 'Unknown error'))
                js = f"window.placeDetailsResult = {{ error: '{self._js_escape(error_msg)}', lat: {lat}, lng: {lng} }};"
                self.page().runJavaScript(js)
        except Exception as e:
            js = f"window.placeDetailsResult = {{ error: '{self._js_escape(str(e))}', lat: {lat}, lng: {lng} }};"
            self.page().runJavaScript(js)
    
    def _js_escape(self, s):
        """Escape string for JavaScript"""
        s = str(s)
        s = s.replace(chr(92), chr(92)+chr(92))  # backslash
        s = s.replace(chr(39), chr(92)+chr(39))  # single quote
        s = s.replace(chr(10), " ")  # newline
        s = s.replace(chr(13), "")   # carriage return
        return s

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
        api_url = "https://maps.googleapis.com/maps/api/js?libraries=geometry,places,routes&callback=initMap&loading=async"
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
        // API key for Routes API calls
        const API_KEY = "{self.api_key}";

        let map;
        let markers = [];
        let polylines = [];
        let infoWindow;
        let placesService;
        let selectedPlaceCallback = null;

        function initMap() {{
            map = new google.maps.Map(document.getElementById('map'), {{
                center: {{ lat: {center_lat}, lng: {center_lng} }},
                zoom: {zoom},
                clickableIcons: true
            }});

            infoWindow = new google.maps.InfoWindow();

            // Initialize Places service for business lookups
            placesService = new google.maps.places.PlacesService(map);

            // Click handler - detect POI clicks which have placeId
            map.addListener('click', (e) => {{
                const hasPlaceId = e.placeId ? true : false;
                
                // MUST stop immediately to prevent default info window
                if (hasPlaceId) {{
                    e.stop();
                }}
                
                if (hasPlaceId) {{
                    getPlaceDetails(e.placeId, e.latLng);
                }} else {{
                    showAddressAtLocation(e.latLng);
                }}
            }});

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

        function getPlaceDetails(placeId, location) {{
            // Request place details from Python (avoids CORS issues)
            window.pendingPlaceIdRequest = {{
                placeId: placeId,
                lat: location.lat(),
                lng: location.lng()
            }};
            
            // Show loading message
            infoWindow.setContent('<div style="padding:10px;">Looking up business...</div>');
            infoWindow.setPosition(location);
            infoWindow.open(map);
            
            // Poll for result from Python
            let attempts = 0;
            const checkResult = setInterval(() => {{
                attempts++;
                if (window.placeDetailsResult) {{
                    clearInterval(checkResult);
                    const result = window.placeDetailsResult;
                    window.placeDetailsResult = null;
                    
                    if (result.error) {{
                        infoWindow.setContent('<div style="padding:10px;">Error: ' + result.error + '</div>');
                        setTimeout(() => {{ showAddressAtLocation(location); }}, 2000);
                    }} else {{
                        window.pendingBusinessName = result.name;
                        const content = '<div class="info-window">' +
                            '<h3>' + result.name + '</h3>' +
                            '<p>' + (result.address || '') + '</p>' +
                            '<p style="margin-top:10px;">' +
                            '<button id="selectBizBtn" ' +
                            'style="background:#1a73e8;color:white;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;font-weight:bold;">' +
                            'Use This Business Name</button></p>' +
                            '</div>';
                        infoWindow.setContent(content);
                        infoWindow.setPosition(location);
                        infoWindow.open(map);

                        google.maps.event.addListenerOnce(infoWindow, 'domready', () => {{
                            const btn = document.getElementById('selectBizBtn');
                            if (btn) {{
                                btn.addEventListener('click', () => {{
                                    selectBusiness(window.pendingBusinessName);
                                }});
                            }}
                        }});
                    }}
                }} else if (attempts > 50) {{
                    // Timeout after 10 seconds
                    clearInterval(checkResult);
                    infoWindow.setContent('<div style="padding:10px;">Timeout waiting for response</div>');
                    setTimeout(() => {{ showAddressAtLocation(location); }}, 2000);
                }}
            }}, 200);
        }}

        function showAddressAtLocation(location, debugInfo) {{
            // Show address for non-POI clicks
            const lat = location.lat();
            const lng = location.lng();
            const debug = debugInfo ? '<p style="color:#999;font-size:11px;">' + debugInfo + '</p>' : '';
            const url = `https://maps.googleapis.com/maps/api/geocode/json?latlng=${{lat}},${{lng}}&key=${{API_KEY}}`;
            
            fetch(url)
                .then(response => response.json())
                .then(data => {{
                    if (data.status === 'OK' && data.results && data.results.length > 0) {{
                        const address = data.results[0].formatted_address;
                        const parts = address.split(',');
                        const shortAddr = parts[0].trim();
                        
                        window.pendingBusinessName = shortAddr;
                        const content = '<div class="info-window">' +
                            '<h3>' + shortAddr + '</h3>' +
                            '<p>' + address + '</p>' +
                            debug +
                            '<p style="margin-top:10px;">' +
                            '<button id="selectBizBtn" ' +
                            'style="background:#1a73e8;color:white;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;font-weight:bold;">' +
                            'Use This Address</button></p>' +
                            '</div>';
                        infoWindow.setContent(content);
                        infoWindow.setPosition(location);
                        infoWindow.open(map);

                        google.maps.event.addListenerOnce(infoWindow, 'domready', () => {{
                            const btn = document.getElementById('selectBizBtn');
                            if (btn) {{
                                btn.addEventListener('click', () => {{
                                    selectBusiness(window.pendingBusinessName);
                                }});
                            }}
                        }});
                    }}
                }})
                .catch(err => {{
                    console.error('Geocode error:', err);
                }});
        }}

        function selectBusiness(businessName) {{
            // Send the selected business name back to Python
            if (window.pyCallback) {{
                window.pyCallback(businessName);
            }}
            // Also store it globally for polling
            window.selectedBusinessName = businessName;
            infoWindow.close();
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

        // DirectionsService for route display
        let directionsService;
        let directionsRenderers = [];

        function initDirections() {{
            if (!directionsService) {{
                directionsService = new google.maps.DirectionsService();
            }}
        }}

        function clearRoutes() {{
            directionsRenderers.forEach(r => r.setMap(null));
            directionsRenderers = [];
            if (window.searchMarker) {{
                window.searchMarker.setMap(null);
                window.searchMarker = null;
            }}
            if (window.searchInfoWindow) {{
                window.searchInfoWindow.close();
                window.searchInfoWindow = null;
            }}
        }}

        async function showRoute(startAddress, endAddress, category, tripInfo) {{
            console.log('showRoute called:', startAddress, '->', endAddress);
            if (!map) {{
                console.error('Map not initialized');
                return 'Map not initialized';
            }}

            clearMap();
            clearRoutes();

            const color = getCategoryColor(category || 'PERSONAL');

            try {{
                // Routes API request per Google documentation
                const requestBody = {{
                    origin: {{
                        address: startAddress
                    }},
                    destination: {{
                        address: endAddress
                    }},
                    travelMode: 'DRIVE',
                    routingPreference: 'TRAFFIC_AWARE',
                    computeAlternativeRoutes: false,
                    routeModifiers: {{
                        avoidTolls: false,
                        avoidHighways: false,
                        avoidFerries: false
                    }},
                    languageCode: 'en-US',
                    units: 'IMPERIAL'
                }};

                console.log('Routes API request:', JSON.stringify(requestBody, null, 2));

                const response = await fetch('https://routes.googleapis.com/directions/v2:computeRoutes', {{
                    method: 'POST',
                    headers: {{
                        'Content-Type': 'application/json',
                        'X-Goog-Api-Key': API_KEY,
                        'X-Goog-FieldMask': 'routes.duration,routes.distanceMeters,routes.polyline.encodedPolyline,routes.legs.startLocation,routes.legs.endLocation'
                    }},
                    body: JSON.stringify(requestBody)
                }});

                console.log('Response status:', response.status);
                const data = await response.json();
                console.log('Routes API response:', JSON.stringify(data, null, 2));

                if (data.error) {{
                    throw new Error(data.error.message || data.error.status || 'Unknown error');
                }}

                if (!data.routes || data.routes.length === 0) {{
                    throw new Error('No route found');
                }}

                const route = data.routes[0];

                // Decode the polyline
                const decodedPath = google.maps.geometry.encoding.decodePath(route.polyline.encodedPolyline);
                console.log('Decoded path has', decodedPath.length, 'points');

                // Check if route is suspiciously simple (less than 5 points for any real road route)
                const isSimplifiedRoute = decodedPath.length < 5;

                // Draw the route
                const routeLine = new google.maps.Polyline({{
                    path: decodedPath,
                    geodesic: true,
                    strokeColor: color,
                    strokeWeight: 5,
                    strokeOpacity: isSimplifiedRoute ? 0 : 0.8,
                    // Show dashed line if route is oversimplified
                    icons: isSimplifiedRoute ? [{{
                        icon: {{ path: 'M 0,-1 0,1', strokeOpacity: 0.6, scale: 4 }},
                        offset: '0',
                        repeat: '20px'
                    }}] : []
                }});
                routeLine.setMap(map);
                polylines.push(routeLine);

                // Get start and end from the path
                const startLoc = decodedPath[0];
                const endLoc = decodedPath[decodedPath.length - 1];

                // Add start marker with "S" label
                const startMarker = new google.maps.Marker({{
                    position: startLoc,
                    map: map,
                    label: {{
                        text: 'S',
                        color: 'white',
                        fontWeight: 'bold',
                        fontSize: '12px'
                    }},
                    icon: {{
                        path: google.maps.SymbolPath.CIRCLE,
                        scale: 14,
                        fillColor: '#4CAF50',
                        fillOpacity: 1,
                        strokeColor: 'white',
                        strokeWeight: 2
                    }},
                    title: 'START: ' + startAddress
                }});
                markers.push(startMarker);

                // Add end marker with "E" label
                const endMarker = new google.maps.Marker({{
                    position: endLoc,
                    map: map,
                    label: {{
                        text: 'E',
                        color: 'white',
                        fontWeight: 'bold',
                        fontSize: '12px'
                    }},
                    icon: {{
                        path: google.maps.SymbolPath.CIRCLE,
                        scale: 14,
                        fillColor: '#D32F2F',
                        fillOpacity: 1,
                        strokeColor: 'white',
                        strokeWeight: 2
                    }},
                    title: 'END: ' + endAddress
                }});
                markers.push(endMarker);

                // Fit map to show entire route with padding
                const bounds = new google.maps.LatLngBounds();
                decodedPath.forEach(point => bounds.extend(point));
                map.fitBounds(bounds, {{ padding: 50 }});

                // Show info window
                if (tripInfo) {{
                    const distanceMiles = (route.distanceMeters / 1609.34).toFixed(1);
                    const durationSecs = parseInt(route.duration.replace('s', ''));
                    const durationMins = Math.round(durationSecs / 60);
                    const categoryClass = (category || 'personal').toLowerCase();

                    const infoContent = '<div class="info-window">' +
                        '<h3>' + (tripInfo.businessName || tripInfo.endAddress) + '</h3>' +
                        '<p><span class="category ' + categoryClass + '">' + (category || 'PERSONAL') + '</span></p>' +
                        '<p><strong>Date:</strong> ' + (tripInfo.date || 'N/A') + '</p>' +
                        '<p><strong>Recorded:</strong> ' + (tripInfo.distance || 0) + ' miles</p>' +
                        '<p><strong>Route:</strong> ' + distanceMiles + ' mi (' + durationMins + ' min)</p>' +
                        '<p><strong>From:</strong> ' + tripInfo.startAddress + '</p>' +
                        '<p><strong>To:</strong> ' + tripInfo.endAddress + '</p>' +
                        '</div>';

                    const infoWin = new google.maps.InfoWindow({{ content: infoContent }});
                    infoWin.open(map, endMarker);
                }}

                console.log('Route displayed successfully');
                return 'Route displayed';

            }} catch (e) {{
                console.error('Routes API error:', e);
                // Show error on map
                const errorDiv = document.createElement('div');
                errorDiv.style.cssText = 'position:absolute;top:10px;left:50%;transform:translateX(-50%);background:#ff5252;color:white;padding:10px 20px;border-radius:4px;font-family:Arial;z-index:1000;max-width:80%;text-align:center;';
                errorDiv.textContent = 'Routes API error: ' + e.message;
                document.body.appendChild(errorDiv);
                setTimeout(() => errorDiv.remove(), 8000);
                // Fall back to showing destination
                showAddress(endAddress);
                return 'Error: ' + e.message;
            }}
        }}

        async function showDailyJourney(trips) {{
            console.log('showDailyJourney called with', trips.length, 'trips');
            if (!map || trips.length === 0) return 'No trips to display';

            clearMap();
            clearRoutes();

            const bounds = new google.maps.LatLngBounds();
            const allPaths = [];

            // Process each trip using Routes API
            for (let i = 0; i < trips.length; i++) {{
                const trip = trips[i];
                const color = getCategoryColor(trip.category);

                try {{
                    const requestBody = {{
                        origin: {{ address: trip.startAddress }},
                        destination: {{ address: trip.endAddress }},
                        travelMode: 'DRIVE',
                        routingPreference: 'TRAFFIC_AWARE',
                        computeAlternativeRoutes: false,
                        languageCode: 'en-US',
                        units: 'IMPERIAL'
                    }};

                    const response = await fetch('https://routes.googleapis.com/directions/v2:computeRoutes', {{
                        method: 'POST',
                        headers: {{
                            'Content-Type': 'application/json',
                            'X-Goog-Api-Key': API_KEY,
                            'X-Goog-FieldMask': 'routes.polyline.encodedPolyline,routes.legs.startLocation,routes.legs.endLocation'
                        }},
                        body: JSON.stringify(requestBody)
                    }});

                    const data = await response.json();
                    if (data.error || !data.routes || data.routes.length === 0) {{
                        console.error('Route failed for trip', i, data.error || 'No routes returned');
                        // Fall back to geocoding the addresses and drawing a dashed line
                        try {{
                            const {{ Place }} = await google.maps.importLibrary("places");
                            const startPlaces = await Place.searchByText({{ textQuery: trip.startAddress, fields: ['location'], maxResultCount: 1 }});
                            const endPlaces = await Place.searchByText({{ textQuery: trip.endAddress, fields: ['location'], maxResultCount: 1 }});

                            if (startPlaces?.places?.length && endPlaces?.places?.length) {{
                                const startLoc = startPlaces.places[0].location;
                                const endLoc = endPlaces.places[0].location;
                                const fallbackPath = [startLoc, endLoc];

                                // Draw dashed line to indicate it's not a real route
                                const fallbackLine = new google.maps.Polyline({{
                                    path: fallbackPath,
                                    geodesic: true,
                                    strokeColor: color,
                                    strokeWeight: 3,
                                    strokeOpacity: 0,
                                    icons: [{{
                                        icon: {{ path: 'M 0,-1 0,1', strokeOpacity: 0.6, scale: 3 }},
                                        offset: '0',
                                        repeat: '15px'
                                    }}]
                                }});
                                fallbackLine.setMap(map);
                                polylines.push(fallbackLine);
                                bounds.extend(startLoc);
                                bounds.extend(endLoc);

                                // Add marker for this trip
                                const marker = new google.maps.Marker({{
                                    position: endLoc,
                                    map: map,
                                    label: {{ text: String(i + 1), color: 'white', fontWeight: 'bold' }},
                                    icon: {{
                                        path: google.maps.SymbolPath.CIRCLE,
                                        scale: 14,
                                        fillColor: color,
                                        fillOpacity: 0.6,
                                        strokeColor: 'white',
                                        strokeWeight: 2
                                    }},
                                    title: 'Trip ' + (i + 1) + ' (route unavailable)'
                                }});
                                markers.push(marker);
                            }}
                        }} catch (fallbackErr) {{
                            console.error('Fallback geocoding also failed:', fallbackErr);
                        }}
                        continue;
                    }}

                    const route = data.routes[0];
                    const decodedPath = google.maps.geometry.encoding.decodePath(route.polyline.encodedPolyline);
                    console.log('Route', i, 'decoded with', decodedPath.length, 'points');
                    allPaths.push({{ path: decodedPath, color: color, trip: trip, index: i }});

                    // Draw the route segment
                    const routeLine = new google.maps.Polyline({{
                        path: decodedPath,
                        geodesic: true,
                        strokeColor: color,
                        strokeWeight: 4,
                        strokeOpacity: 0.8
                    }});
                    routeLine.setMap(map);
                    polylines.push(routeLine);

                    // Extend bounds
                    decodedPath.forEach(point => bounds.extend(point));

                    // Add numbered marker at start
                    const startLoc = decodedPath[0];
                    const startMarker = new google.maps.Marker({{
                        position: startLoc,
                        map: map,
                        label: {{ text: String(i + 1), color: 'white', fontWeight: 'bold' }},
                        icon: {{
                            path: google.maps.SymbolPath.CIRCLE,
                            scale: 14,
                            fillColor: color,
                            fillOpacity: 1,
                            strokeColor: 'white',
                            strokeWeight: 2
                        }},
                        title: 'Stop ' + (i + 1) + ': ' + trip.startAddress
                    }});

                    startMarker.addListener('click', ((idx, t) => () => {{
                        const categoryClass = (t.category || 'personal').toLowerCase();
                        infoWindow.setContent(
                            '<div class="info-window">' +
                            '<h3>Trip ' + (idx + 1) + ': ' + (t.businessName || t.endAddress) + '</h3>' +
                            '<p><span class="category ' + categoryClass + '">' + (t.category || 'PERSONAL') + '</span></p>' +
                            '<p><strong>Time:</strong> ' + (t.time || 'N/A') + '</p>' +
                            '<p><strong>Distance:</strong> ' + (t.distance || 0) + ' miles</p>' +
                            '<p><strong>From:</strong> ' + t.startAddress + '</p>' +
                            '<p><strong>To:</strong> ' + t.endAddress + '</p>' +
                            '</div>'
                        );
                        infoWindow.open(map, startMarker);
                    }})(i, trip));

                    markers.push(startMarker);

                }} catch (e) {{
                    console.error('Error processing trip', i, e);
                }}
            }}

            // Add final destination marker
            if (allPaths.length > 0) {{
                const lastPath = allPaths[allPaths.length - 1];
                const lastTrip = lastPath.trip;
                const endLoc = lastPath.path[lastPath.path.length - 1];
                const endMarker = new google.maps.Marker({{
                    position: endLoc,
                    map: map,
                    label: {{ text: 'END', color: 'white', fontWeight: 'bold', fontSize: '10px' }},
                    icon: {{
                        path: google.maps.SymbolPath.CIRCLE,
                        scale: 16,
                        fillColor: '#d32f2f',
                        fillOpacity: 1,
                        strokeColor: 'white',
                        strokeWeight: 3
                    }},
                    title: 'Final: ' + lastTrip.endAddress
                }});
                markers.push(endMarker);

                map.fitBounds(bounds, {{ padding: 50 }});
            }}

            return 'Daily journey displayed with ' + allPaths.length + ' routes';
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

    def show_route(self, trip):
        """Show route for a single trip with directions"""
        self.selected_trip = trip  # Store for business name updates

        start_addr = trip.get('start_address', '')
        end_addr = trip.get('end_address', '')
        category = trip.get('computed_category', 'PERSONAL')

        trip_info = {
            'date': trip['started'].strftime('%Y-%m-%d %H:%M') if hasattr(trip.get('started'), 'strftime') else str(trip.get('started', '')),
            'distance': round(trip.get('distance', 0), 1),
            'startAddress': start_addr,
            'endAddress': end_addr,
            'businessName': trip.get('business_name', '')
        }

        # Use JSON encoding for safe JavaScript string passing
        js = f"showRoute({json.dumps(start_addr)}, {json.dumps(end_addr)}, {json.dumps(category)}, {json.dumps(trip_info)});"
        self.page().runJavaScript(js)

    def show_daily_journey(self, trips):
        """Show all trips for a day connected together"""
        if not trips:
            return

        from datetime import datetime
        sorted_trips = sorted(trips, key=lambda t: t.get('started', datetime.min))

        trips_data = []
        for trip in sorted_trips:
            trip_data = {
                'startAddress': trip.get('start_address', '').replace("'", "\\'"),
                'endAddress': trip.get('end_address', '').replace("'", "\\'"),
                'category': trip.get('computed_category', 'PERSONAL'),
                'time': trip['started'].strftime('%H:%M') if hasattr(trip.get('started'), 'strftime') else '',
                'distance': round(trip.get('distance', 0), 1),
                'businessName': trip.get('business_name', '')
            }
            trips_data.append(trip_data)

        trips_json = json.dumps(trips_data)
        js = f"showDailyJourney({trips_json});"
        self.page().runJavaScript(js)


class SummaryWidget(QWidget):
    """Widget displaying summary statistics with visual cards"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _create_stat_card(self, title: str, color: str, bg_color: str) -> dict:
        """Create a styled stat card widget"""
        card = QFrame()
        card.setFrameStyle(QFrame.Shape.StyledPanel)
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {bg_color};
                border: 2px solid {color};
                border-radius: 8px;
                padding: 10px;
            }}
            QLabel {{
                border: none;
                background: transparent;
                padding: 0px;
            }}
        """)

        layout = QVBoxLayout(card)
        layout.setSpacing(5)

        # Title
        title_label = QLabel(title)
        title_label.setStyleSheet(f"color: {color}; font-weight: bold; font-size: 12px; border: none;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        # Miles (big number)
        miles_label = QLabel("0.0")
        miles_label.setStyleSheet(f"color: {color}; font-size: 28px; font-weight: bold; border: none;")
        miles_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(miles_label)

        # Miles text
        miles_text = QLabel("miles")
        miles_text.setStyleSheet(f"color: {color}; font-size: 11px; border: none;")
        miles_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(miles_text)

        # Percentage
        pct_label = QLabel("")
        pct_label.setStyleSheet(f"color: {color}; font-size: 14px; font-weight: bold; border: none;")
        pct_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(pct_label)

        # Trip count
        trips_label = QLabel("0 trips")
        trips_label.setStyleSheet(f"color: {color}; font-size: 11px; border: none;")
        trips_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(trips_label)

        return {
            'card': card,
            'miles': miles_label,
            'pct': pct_label,
            'trips': trips_label
        }

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Title
        title = QLabel("Mileage Summary")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #333; padding: 10px; border: none;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Date range label
        self.date_range_label = QLabel("")
        self.date_range_label.setStyleSheet("font-size: 12px; color: #666; padding-bottom: 10px; border: none;")
        self.date_range_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.date_range_label)

        # Total miles card (larger, at top)
        total_frame = QFrame()
        total_frame.setStyleSheet("""
            QFrame {
                background-color: #f5f5f5;
                border: 2px solid #424242;
                border-radius: 10px;
                padding: 15px;
            }
            QLabel {
                border: none;
                background: transparent;
            }
        """)
        total_layout = QVBoxLayout(total_frame)

        total_title = QLabel("TOTAL MILEAGE")
        total_title.setStyleSheet("color: #424242; font-weight: bold; font-size: 12px; border: none;")
        total_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        total_layout.addWidget(total_title)

        self.total_miles_label = QLabel("0.0")
        self.total_miles_label.setStyleSheet("color: #212121; font-size: 42px; font-weight: bold; border: none;")
        self.total_miles_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        total_layout.addWidget(self.total_miles_label)

        total_text = QLabel("miles")
        total_text.setStyleSheet("color: #616161; font-size: 14px; border: none;")
        total_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        total_layout.addWidget(total_text)

        self.total_trips_label = QLabel("0 trips")
        self.total_trips_label.setStyleSheet("color: #616161; font-size: 12px; border: none;")
        self.total_trips_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        total_layout.addWidget(self.total_trips_label)

        layout.addWidget(total_frame)

        # Category cards in a row
        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(10)

        # Business card (green)
        self.business_card = self._create_stat_card("BUSINESS", "#2e7d32", "#e8f5e9")
        cards_layout.addWidget(self.business_card['card'])

        # Personal card (orange)
        self.personal_card = self._create_stat_card("PERSONAL", "#e65100", "#fff3e0")
        cards_layout.addWidget(self.personal_card['card'])

        # Commute card (blue)
        self.commute_card = self._create_stat_card("COMMUTE", "#1565c0", "#e3f2fd")
        cards_layout.addWidget(self.commute_card['card'])

        layout.addLayout(cards_layout)

        # Additional stats section
        extras_frame = QFrame()
        extras_frame.setStyleSheet("""
            QFrame {
                background-color: #fafafa;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 10px;
            }
            QLabel {
                border: none;
                background: transparent;
            }
        """)
        extras_layout = QVBoxLayout(extras_frame)

        extras_title = QLabel("Additional Statistics")
        extras_title.setStyleSheet("font-weight: bold; color: #666; border: none;")
        extras_layout.addWidget(extras_title)

        # Grid for extra stats
        extras_grid = QHBoxLayout()

        # Avg per trip
        avg_frame = QVBoxLayout()
        self.avg_miles_label = QLabel("0.0")
        self.avg_miles_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #333; border: none;")
        self.avg_miles_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        avg_frame.addWidget(self.avg_miles_label)
        avg_text = QLabel("avg miles/trip")
        avg_text.setStyleSheet("font-size: 10px; color: #666; border: none;")
        avg_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        avg_frame.addWidget(avg_text)
        extras_grid.addLayout(avg_frame)

        # Business per week
        biz_week_frame = QVBoxLayout()
        self.biz_per_week_label = QLabel("0.0")
        self.biz_per_week_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #2e7d32; border: none;")
        self.biz_per_week_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        biz_week_frame.addWidget(self.biz_per_week_label)
        biz_week_text = QLabel("business mi/week")
        biz_week_text.setStyleSheet("font-size: 10px; color: #666; border: none;")
        biz_week_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        biz_week_frame.addWidget(biz_week_text)
        extras_grid.addLayout(biz_week_frame)

        # Days with trips
        days_frame = QVBoxLayout()
        self.days_with_trips_label = QLabel("0")
        self.days_with_trips_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #333; border: none;")
        self.days_with_trips_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        days_frame.addWidget(self.days_with_trips_label)
        days_text = QLabel("days with trips")
        days_text.setStyleSheet("font-size: 10px; color: #666; border: none;")
        days_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        days_frame.addWidget(days_text)
        extras_grid.addLayout(days_frame)

        extras_layout.addLayout(extras_grid)
        layout.addWidget(extras_frame)

        layout.addStretch()

    def update_stats(self, data: dict):
        """Update the summary display with new data"""
        totals = data.get('totals', {})
        trips = data.get('trips', [])
        weekly_stats = data.get('weekly_stats', {})

        # Total miles
        total_miles = totals.get('total_miles', 0)
        self.total_miles_label.setText(f"{total_miles:,.1f}")
        self.total_trips_label.setText(f"{len(trips)} trips")

        # Business
        business_miles = totals.get('business_miles', 0)
        business_pct = totals.get('business_pct', 0)
        business_count = sum(1 for t in trips if t.get('computed_category') == 'BUSINESS')
        self.business_card['miles'].setText(f"{business_miles:,.1f}")
        self.business_card['pct'].setText(f"{business_pct:.1f}%")
        self.business_card['trips'].setText(f"{business_count} trips")

        # Personal
        personal_miles = totals.get('personal_miles', 0)
        personal_pct = totals.get('personal_pct', 0)
        personal_count = sum(1 for t in trips if t.get('computed_category') == 'PERSONAL')
        self.personal_card['miles'].setText(f"{personal_miles:,.1f}")
        self.personal_card['pct'].setText(f"{personal_pct:.1f}%")
        self.personal_card['trips'].setText(f"{personal_count} trips")

        # Commute
        commute_miles = totals.get('commute_miles', 0)
        commute_pct = totals.get('commute_pct', 0)
        commute_count = sum(1 for t in trips if t.get('computed_category') == 'COMMUTE')
        self.commute_card['miles'].setText(f"{commute_miles:,.1f}")
        self.commute_card['pct'].setText(f"{commute_pct:.1f}%")
        self.commute_card['trips'].setText(f"{commute_count} trips")

        # Date range
        if trips:
            dates = [t['started'] for t in trips if t.get('started')]
            if dates:
                min_date = min(dates).strftime('%b %d, %Y')
                max_date = max(dates).strftime('%b %d, %Y')
                self.date_range_label.setText(f"{min_date} — {max_date}")

        # Additional stats
        if trips:
            avg_miles = total_miles / len(trips)
            self.avg_miles_label.setText(f"{avg_miles:.1f}")

            # Unique days with trips
            unique_days = len(set(t['started'].date() for t in trips if t.get('started')))
            self.days_with_trips_label.setText(str(unique_days))

            # Business miles per week
            if weekly_stats:
                num_weeks = len(weekly_stats)
                if num_weeks > 0:
                    biz_per_week = business_miles / num_weeks
                    self.biz_per_week_label.setText(f"{biz_per_week:.1f}")
        else:
            self.avg_miles_label.setText("0.0")
            self.days_with_trips_label.setText("0")
            self.biz_per_week_label.setText("0.0")


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

        # Header with count and sort options
        header_layout = QHBoxLayout()
        self.count_label = QLabel("No unresolved addresses")
        self.count_label.setStyleSheet("font-weight: bold; padding: 5px;")
        header_layout.addWidget(self.count_label)

        header_layout.addWidget(QLabel("Sort:"))
        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["By Visits", "By Miles", "By Address", "By Street"])
        self.sort_combo.currentTextChanged.connect(self._sort_list)
        header_layout.addWidget(self.sort_combo)

        refresh_btn = QPushButton("Refresh")
        refresh_btn.setMaximumWidth(80)
        refresh_btn.clicked.connect(self._refresh_list)
        header_layout.addWidget(refresh_btn)

        layout.addLayout(header_layout)

        # Tip for multi-select
        tip_label = QLabel("Tip: Hold Ctrl to select multiple nearby addresses, then label them together")
        tip_label.setStyleSheet("color: #666; font-style: italic; padding: 2px 5px;")
        layout.addWidget(tip_label)

        # Address list - now with multi-select
        self.address_list = QTableWidget()
        self.address_list.setColumnCount(4)
        self.address_list.setHorizontalHeaderLabels(["Address", "Street", "Visits", "Miles"])
        self.address_list.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.address_list.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)  # Multi-select!
        self.address_list.setAlternatingRowColors(True)
        self.address_list.setSortingEnabled(True)

        header = self.address_list.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        self.address_list.setColumnWidth(1, 120)
        self.address_list.setColumnWidth(2, 50)
        self.address_list.setColumnWidth(3, 60)

        self.address_list.itemSelectionChanged.connect(self._on_address_selected)
        self.address_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.address_list.customContextMenuRequested.connect(self._show_address_context_menu)
        layout.addWidget(self.address_list)

        # Business name input area
        input_group = QGroupBox("Assign Location/Business Name")
        input_layout = QVBoxLayout(input_group)

        self.selected_address_label = QLabel("Select address(es) above")
        self.selected_address_label.setWordWrap(True)
        self.selected_address_label.setStyleSheet("color: #666; padding: 5px; background: #f5f5f5;")
        self.selected_address_label.setMaximumHeight(60)
        input_layout.addWidget(self.selected_address_label)

        # Quick pick from existing locations
        quick_pick_layout = QHBoxLayout()
        quick_pick_layout.addWidget(QLabel("Quick Pick:"))

        self.location_combo = QComboBox()
        self.location_combo.setMinimumWidth(200)
        self.location_combo.addItem("-- Select existing location --")
        self.location_combo.currentIndexChanged.connect(self._on_location_selected)
        quick_pick_layout.addWidget(self.location_combo)

        self.apply_location_btn = QPushButton("Apply")
        self.apply_location_btn.clicked.connect(self._apply_selected_location)
        self.apply_location_btn.setEnabled(False)
        quick_pick_layout.addWidget(self.apply_location_btn)

        quick_pick_layout.addStretch()
        input_layout.addLayout(quick_pick_layout)

        # Or enter new name
        input_layout.addWidget(QLabel("Or enter new name:"))

        name_layout = QHBoxLayout()
        self.business_name_input = QLineEdit()
        self.business_name_input.setPlaceholderText("Enter business name for selected address(es)...")
        self.business_name_input.returnPressed.connect(self._save_mapping)
        name_layout.addWidget(self.business_name_input)

        # Category dropdown
        name_layout.addWidget(QLabel("Category:"))
        self.category_combo = QComboBox()
        self.category_combo.addItems(["Business", "Personal", "Commute"])
        self.category_combo.setCurrentText("Business")
        self.category_combo.setFixedWidth(100)
        name_layout.addWidget(self.category_combo)

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

        self.select_nearby_btn = QPushButton("Select Nearby")
        self.select_nearby_btn.setToolTip("Select other addresses on the same street")
        self.select_nearby_btn.clicked.connect(self._select_nearby)
        self.select_nearby_btn.setEnabled(False)
        action_layout.addWidget(self.select_nearby_btn)

        action_layout.addStretch()
        input_layout.addLayout(action_layout)

        layout.addWidget(input_group)

        self.current_address = None
        self.selected_addresses = []  # List of selected addresses for batch operations

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

        # Sort by visit count (most visited first) and extract street names
        self.addresses_data = []
        for addr, stats in sorted(address_stats.items(), key=lambda x: -x[1]['visits']):
            street = self._extract_street(addr)
            self.addresses_data.append({
                'address': addr,
                'street': street,
                'visits': stats['visits'],
                'miles': stats['miles'],
                'lat': stats.get('lat'),
                'lng': stats.get('lng')
            })

        self._populate_list()
        self._populate_location_dropdown()

    def _populate_location_dropdown(self):
        """Populate the quick pick dropdown with existing locations"""
        self.location_combo.blockSignals(True)
        self.location_combo.clear()
        self.location_combo.addItem("-- Select existing location --")

        # Add standard locations
        self.location_combo.addItem("Home", "Home")
        self.location_combo.addItem("Office", "Office")
        self.location_combo.addItem("[PERSONAL]", "[PERSONAL]")

        # Load existing business names from mapping file
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        existing_names = set()
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
                    for name in mappings.values():
                        if name and name not in ['Home', 'Office', '[PERSONAL]', 'Unknown']:
                            existing_names.add(name)
            except:
                pass

        # Load from address cache too
        cache_file = os.path.join(get_app_dir(), 'address_cache.json')
        if os.path.exists(cache_file):
            try:
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cache = json.load(f)
                    for entry in cache.values():
                        if isinstance(entry, dict):
                            name = entry.get('name', '')
                        else:
                            name = str(entry)
                        if name and name not in ['Home', 'Office', '[PERSONAL]', 'Unknown', '']:
                            existing_names.add(name)
            except:
                pass

        # Add separator and existing names (sorted)
        if existing_names:
            self.location_combo.insertSeparator(self.location_combo.count())
            for name in sorted(existing_names):
                self.location_combo.addItem(name, name)

        self.location_combo.blockSignals(False)

    def _on_location_selected(self, index: int):
        """Handle location dropdown selection"""
        if index > 0:  # Not the placeholder
            self.apply_location_btn.setEnabled(bool(self.selected_addresses))
        else:
            self.apply_location_btn.setEnabled(False)

    def _apply_selected_location(self):
        """Apply the selected location from dropdown to selected addresses"""
        if not self.selected_addresses:
            return

        location = self.location_combo.currentData()
        if not location:
            return

        self._save_multiple_to_mapping_file(self.selected_addresses, location)
        self.location_combo.setCurrentIndex(0)  # Reset dropdown
        self._remove_selected_from_list()
        self.mapping_saved.emit()

    def _extract_street(self, address: str) -> str:
        """Extract street name from address for grouping nearby addresses"""
        import re
        # Try to extract the street name (everything after the house number, before city)
        # Examples: "15827 61st Ln NE, Kenmore" -> "61st Ln NE"
        #           "80th St SW, Everett" -> "80th St SW"
        parts = address.split(',')
        if parts:
            street_part = parts[0].strip()
            # Remove leading numbers (house number)
            match = re.match(r'^\d+\s+(.+)$', street_part)
            if match:
                return match.group(1)
            return street_part
        return address

    def _populate_list(self):
        """Populate the address list table"""
        self.address_list.setSortingEnabled(False)
        self.address_list.setRowCount(len(self.addresses_data))

        for row, data in enumerate(self.addresses_data):
            addr_item = QTableWidgetItem(data['address'])
            addr_item.setData(Qt.ItemDataRole.UserRole, row)  # Store original index
            self.address_list.setItem(row, 0, addr_item)

            self.address_list.setItem(row, 1, QTableWidgetItem(data.get('street', '')))

            # Use numeric sort for visits and miles
            visits_item = QTableWidgetItem()
            visits_item.setData(Qt.ItemDataRole.DisplayRole, data['visits'])
            self.address_list.setItem(row, 2, visits_item)

            miles_item = QTableWidgetItem()
            miles_item.setData(Qt.ItemDataRole.DisplayRole, round(data['miles'], 1))
            self.address_list.setItem(row, 3, miles_item)

        self.address_list.setSortingEnabled(True)
        count = len(self.addresses_data)
        self.count_label.setText(f"{count} unresolved address{'es' if count != 1 else ''}")

    def _sort_list(self, sort_by: str):
        """Sort the address list"""
        if sort_by == "By Visits":
            self.address_list.sortByColumn(2, Qt.SortOrder.DescendingOrder)
        elif sort_by == "By Miles":
            self.address_list.sortByColumn(3, Qt.SortOrder.DescendingOrder)
        elif sort_by == "By Address":
            self.address_list.sortByColumn(0, Qt.SortOrder.AscendingOrder)
        elif sort_by == "By Street":
            self.address_list.sortByColumn(1, Qt.SortOrder.AscendingOrder)

    def _on_address_selected(self):
        """Handle address selection (supports multi-select)"""
        selected_rows = self.address_list.selectionModel().selectedRows()

        # Get selected addresses from the visible table (may be sorted)
        self.selected_addresses = []
        for row_index in selected_rows:
            row = row_index.row()
            addr_item = self.address_list.item(row, 0)
            if addr_item:
                self.selected_addresses.append(addr_item.text())

        if self.selected_addresses:
            # Set current_address to first selected (for backward compatibility)
            self.current_address = self.selected_addresses[0]

            # Update label based on selection count
            if len(self.selected_addresses) == 1:
                self.selected_address_label.setText(self.current_address)
                self.selected_address_label.setStyleSheet("color: #000; padding: 5px; background: #e3f2fd;")
            else:
                # Show count and first few addresses
                preview = ", ".join(self.selected_addresses[:3])
                if len(self.selected_addresses) > 3:
                    preview += f"... (+{len(self.selected_addresses) - 3} more)"
                self.selected_address_label.setText(f"{len(self.selected_addresses)} addresses selected: {preview}")
                self.selected_address_label.setStyleSheet("color: #000; padding: 5px; background: #fff3e0;")

            # Enable buttons
            self.save_btn.setEnabled(True)
            self.personal_btn.setEnabled(True)
            self.view_map_btn.setEnabled(True)
            self.open_gmaps_btn.setEnabled(True)
            self.select_nearby_btn.setEnabled(True)
            # Enable apply button if a location is selected in dropdown
            self.apply_location_btn.setEnabled(self.location_combo.currentIndex() > 0)

            # Emit signal to show first address on map
            for data in self.addresses_data:
                if data['address'] == self.current_address:
                    lat = data.get('lat') or 0.0
                    lng = data.get('lng') or 0.0
                    self.address_selected.emit(self.current_address, lat, lng)
                    break
        else:
            self.current_address = None
            self.selected_addresses = []
            self.selected_address_label.setText("Select address(es) above")
            self.selected_address_label.setStyleSheet("color: #666; padding: 5px; background: #f5f5f5;")
            self.save_btn.setEnabled(False)
            self.personal_btn.setEnabled(False)
            self.view_map_btn.setEnabled(False)
            self.open_gmaps_btn.setEnabled(False)
            self.select_nearby_btn.setEnabled(False)
            self.apply_location_btn.setEnabled(False)

    def _save_mapping(self):
        """Save the business name mapping for all selected addresses"""
        if not self.selected_addresses:
            return

        name = self.business_name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "No Name", "Please enter a business name.")
            return

        # Get category from dropdown
        category = self.category_combo.currentText().upper()

        # Save all selected addresses with the same name and category
        self._save_multiple_to_mapping_file(self.selected_addresses, name, category)
        self.business_name_input.clear()
        self._remove_selected_from_list()
        self.mapping_saved.emit()

    def _mark_personal(self):
        """Mark all selected addresses as personal"""
        if not self.selected_addresses:
            return

        self._save_multiple_to_mapping_file(self.selected_addresses, "[PERSONAL]", "PERSONAL")
        self._remove_selected_from_list()
        self.mapping_saved.emit()

    def _save_multiple_to_mapping_file(self, addresses: List[str], name: str, category: str = None):
        """Save multiple address mappings to the business_mapping.json file"""
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')

        # Load existing mapping
        mappings = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except:
                pass

        # Add all mappings with new format including category and source
        for address in addresses:
            entry = {"name": name, "source": "manual"}
            if category:
                entry["category"] = category
            mappings[address] = entry

        # Save file
        try:
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, indent=2, ensure_ascii=False)
            if len(addresses) == 1:
                QMessageBox.information(self, "Saved", f"Saved: {addresses[0][:40]}... = {name}")
            else:
                QMessageBox.information(self, "Saved", f"Saved {len(addresses)} addresses as: {name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save: {e}")

    def _save_to_mapping_file(self, address: str, name: str):
        """Save a single mapping to the business_mapping.json file (legacy)"""
        self._save_multiple_to_mapping_file([address], name)

    def _remove_selected_from_list(self):
        """Remove all selected addresses from the list"""
        if self.selected_addresses:
            addresses_to_remove = set(self.selected_addresses)
            self.addresses_data = [d for d in self.addresses_data if d['address'] not in addresses_to_remove]
            self._populate_list()
            self.current_address = None
            self.selected_addresses = []
            self._on_address_selected()

    def _remove_current_from_list(self):
        """Remove the current address from the list (legacy - uses selected)"""
        self._remove_selected_from_list()

    def _select_nearby(self):
        """Select all addresses near the current selection (by distance or street)"""
        if not self.current_address:
            return

        # Find the current address data
        current_data = None
        for data in self.addresses_data:
            if data['address'] == self.current_address:
                current_data = data
                break

        if not current_data:
            return

        current_street = current_data.get('street', '')
        current_lat = current_data.get('lat')
        current_lng = current_data.get('lng')

        # Load business mapping to find already-resolved nearby addresses
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        business_mapping = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    business_mapping = json.load(f)
            except:
                pass

        # Find nearby addresses in the UNRESOLVED list
        nearby_unresolved = []
        for data in self.addresses_data:
            if data['address'] == self.current_address:
                continue  # Skip self

            match_reason = self._check_nearby_match(
                current_street, current_lat, current_lng,
                data.get('street', ''), data.get('lat'), data.get('lng')
            )

            if match_reason:
                nearby_unresolved.append({
                    'address': data['address'],
                    'reason': match_reason,
                    'visits': data.get('visits', 0),
                    'miles': data.get('miles', 0),
                    'name': data.get('business_name', ''),
                    'resolved': False
                })

        # Also find nearby addresses that are ALREADY RESOLVED in business_mapping
        nearby_resolved = []
        for mapped_addr, mapped_name in business_mapping.items():
            if mapped_addr == self.current_address:
                continue

            # Extract street from mapped address
            mapped_street = self._extract_street(mapped_addr)

            match_reason = self._check_nearby_match(
                current_street, current_lat, current_lng,
                mapped_street, None, None  # No lat/lng for mapped addresses
            )

            if match_reason:
                nearby_resolved.append({
                    'address': mapped_addr,
                    'reason': match_reason,
                    'name': mapped_name,
                    'resolved': True
                })

        # Build the dialog message
        suggested_name = None
        name_counts = {}

        # Count business names from resolved addresses to suggest most common
        for n in nearby_resolved:
            name = n.get('name', '')
            if name and name not in ['[PERSONAL]', 'Unknown', 'Home', 'Office']:
                name_counts[name] = name_counts.get(name, 0) + 1

        if name_counts:
            suggested_name = max(name_counts, key=name_counts.get)

        if not nearby_unresolved and not nearby_resolved:
            QMessageBox.information(
                self, "No Nearby Addresses",
                f"No addresses found near:\n{self.current_address}\n\n"
                "Try selecting addresses manually with Ctrl+Click."
            )
            return

        # Build message
        msg = f"Selected: {self.current_address[:55]}...\n\n"

        # Show suggestion if we found resolved nearby addresses
        if suggested_name:
            msg += f"SUGGESTED NAME: {suggested_name}\n"
            msg += f"  ({name_counts[suggested_name]} nearby address(es) already use this name)\n\n"

        # Show already-resolved nearby addresses
        if nearby_resolved:
            msg += f"Already resolved nearby ({len(nearby_resolved)}):\n"
            for n in nearby_resolved[:5]:
                name_display = n.get('name', 'Unknown')
                msg += f"  ✓ {name_display}: {n['address'][:35]}...\n"
            if len(nearby_resolved) > 5:
                msg += f"    ... and {len(nearby_resolved) - 5} more\n"
            msg += "\n"

        # Show unresolved nearby addresses
        if nearby_unresolved:
            msg += f"Unresolved nearby ({len(nearby_unresolved)}):\n"
            for n in nearby_unresolved[:5]:
                msg += f"  • {n['address'][:45]}...\n"
            if len(nearby_unresolved) > 5:
                msg += f"    ... and {len(nearby_unresolved) - 5} more\n"
            msg += "\n"

        if nearby_unresolved:
            msg += "Select the unresolved addresses?"

            reply = QMessageBox.question(
                self, "Select Nearby Addresses",
                msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                # Select all matching unresolved rows (keep current selection too)
                nearby_addresses = {n['address'] for n in nearby_unresolved}
                nearby_addresses.add(self.current_address)

                self.address_list.clearSelection()
                for row in range(self.address_list.rowCount()):
                    addr_item = self.address_list.item(row, 0)
                    if addr_item and addr_item.text() in nearby_addresses:
                        self.address_list.selectRow(row)

                # Pre-fill the business name input with the suggested name
                if suggested_name:
                    self.business_name_input.setText(suggested_name)
        else:
            # No unresolved addresses, just show info
            QMessageBox.information(
                self, "Nearby Addresses",
                msg + "All nearby addresses are already resolved!"
            )

    def _check_nearby_match(self, current_street: str, current_lat, current_lng,
                            other_street: str, other_lat, other_lng) -> str:
        """Check if an address matches nearby criteria. Returns match reason or None."""
        import re

        # Check 1: Same street name (exact match)
        if other_street == current_street and current_street:
            return f"Same street: {current_street}"

        # Check 2: Similar street name (fuzzy match - same street number pattern)
        if current_street and other_street:
            current_num = re.match(r'^(\d+)', current_street)
            other_num = re.match(r'^(\d+)', other_street)
            if current_num and other_num and current_num.group(1) == other_num.group(1):
                return f"Similar street: {other_street}"

        # Check 3: Geographic proximity (within 0.25 miles / ~400 meters)
        if current_lat and current_lng and other_lat and other_lng:
            distance = self._calculate_distance(current_lat, current_lng, other_lat, other_lng)
            if distance <= 0.25:
                return f"Within {distance:.2f} mi"

        return None

    def _calculate_distance(self, lat1: float, lng1: float, lat2: float, lng2: float) -> float:
        """Calculate distance between two points in miles (Haversine formula)"""
        import math
        R = 3959  # Earth's radius in miles

        lat1_rad = math.radians(lat1)
        lat2_rad = math.radians(lat2)
        delta_lat = math.radians(lat2 - lat1)
        delta_lng = math.radians(lng2 - lng1)

        a = math.sin(delta_lat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(delta_lng/2)**2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))

        return R * c

    def _show_address_context_menu(self, pos):
        """Show context menu on right-click in address list"""
        row = self.address_list.rowAt(pos.y())
        if row < 0:
            return

        # Select the row if not already selected
        if not self.address_list.item(row, 0).isSelected():
            self.address_list.selectRow(row)

        menu = QMenu(self)

        # Assign name submenu with quick picks
        assign_menu = menu.addMenu("Assign Name")
        home_action = assign_menu.addAction("Home")
        office_action = assign_menu.addAction("Office")
        personal_action = assign_menu.addAction("[PERSONAL]")
        assign_menu.addSeparator()

        # Add existing business names from mapping
        existing_names = self._get_existing_business_names()
        name_actions = {}
        if existing_names:
            recent_menu = assign_menu.addMenu("Existing Names")
            for name in sorted(existing_names)[:15]:  # Limit to 15 most common
                action = recent_menu.addAction(name)
                name_actions[action] = name
            assign_menu.addSeparator()

        custom_action = assign_menu.addAction("Custom Name...")

        menu.addSeparator()

        # Selection actions
        select_nearby_action = menu.addAction("Select Nearby...")

        menu.addSeparator()

        # Map actions
        view_map_action = menu.addAction("View on Map")
        open_gmaps_action = menu.addAction("Open in Google Maps")

        action = menu.exec(self.address_list.viewport().mapToGlobal(pos))

        if action == home_action:
            self._apply_name_to_selected("Home")
        elif action == office_action:
            self._apply_name_to_selected("Office")
        elif action == personal_action:
            self._apply_name_to_selected("[PERSONAL]")
        elif action == custom_action:
            self._prompt_custom_name()
        elif action in name_actions:
            self._apply_name_to_selected(name_actions[action])
        elif action == select_nearby_action:
            self._select_nearby()
        elif action == view_map_action:
            self._view_on_map()
        elif action == open_gmaps_action:
            self._open_in_google_maps()

    def _get_existing_business_names(self) -> set:
        """Get existing business names from mapping and cache files"""
        names = set()
        skip_names = {'Home', 'Office', '[PERSONAL]', 'Unknown', '', 'NO_BUSINESS_FOUND'}

        # From business mapping
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
                    for name in mappings.values():
                        if name and name not in skip_names:
                            names.add(name)
            except:
                pass

        # From address cache
        cache_file = os.path.join(get_app_dir(), 'address_cache.json')
        if os.path.exists(cache_file):
            try:
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cache = json.load(f)
                    for entry in cache.values():
                        if isinstance(entry, dict):
                            name = entry.get('name', '')
                        else:
                            name = str(entry)
                        if name and name not in skip_names:
                            names.add(name)
            except:
                pass

        return names

    def _apply_name_to_selected(self, name: str):
        """Apply a name to all selected addresses"""
        if not self.selected_addresses:
            return
        self._save_multiple_to_mapping_file(self.selected_addresses, name)
        self._remove_selected_from_list()
        self.mapping_saved.emit()

    def _prompt_custom_name(self):
        """Prompt for a custom name and apply to selected addresses"""
        if not self.selected_addresses:
            return

        from PyQt6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(
            self,
            "Custom Business Name",
            f"Enter name for {len(self.selected_addresses)} address(es):",
            text=self.business_name_input.text()
        )

        if ok and name.strip():
            self._apply_name_to_selected(name.strip())

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


class SettingsDialog(QDialog):
    """Dialog for editing application settings including addresses and API keys"""

    settings_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setMinimumSize(550, 500)
        self.resize(600, 550)

        self.config_file = "config.json"
        self.config = self._load_config()

        self._setup_ui()

    def _load_config(self) -> dict:
        """Load current configuration"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {}

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Title
        title = QLabel("Application Settings")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #1976d2;")
        layout.addWidget(title)

        # Scroll area for settings
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(20)

        # === ADDRESSES SECTION ===
        addr_group = QGroupBox("Home & Office Addresses")
        addr_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }
        """)
        addr_layout = QVBoxLayout(addr_group)
        addr_layout.setSpacing(12)

        # Home address
        home_label = QLabel("Home Address:")
        home_label.setToolTip("Your home address - trips starting/ending here help identify commutes")
        addr_layout.addWidget(home_label)

        self.home_address = QLineEdit()
        self.home_address.setPlaceholderText("e.g., 123 Main St, Seattle, WA 98101")
        self.home_address.setText(self.config.get('home_address', ''))
        self.home_address.setToolTip(
            "Enter your home address.\n\n"
            "This is used to identify commute trips and distinguish\n"
            "between personal and business travel."
        )
        addr_layout.addWidget(self.home_address)

        # Work address
        work_label = QLabel("Work/Office Address:")
        work_label.setToolTip("Your primary workplace - trips to/from here are categorized as commute")
        addr_layout.addWidget(work_label)

        self.work_address = QLineEdit()
        self.work_address.setPlaceholderText("e.g., 456 Corporate Blvd, Bellevue, WA 98004")
        self.work_address.setText(self.config.get('work_address', ''))
        self.work_address.setToolTip(
            "Enter your primary workplace address.\n\n"
            "Trips between home and this address are categorized\n"
            "as commute (not tax-deductible business miles)."
        )
        addr_layout.addWidget(self.work_address)

        scroll_layout.addWidget(addr_group)

        # === API KEYS SECTION ===
        api_group = QGroupBox("API Configuration")
        api_group.setStyleSheet(addr_group.styleSheet())
        api_layout = QVBoxLayout(api_group)
        api_layout.setSpacing(12)

        # Google Places API key
        api_label = QLabel("Google Places API Key:")
        api_layout.addWidget(api_label)

        api_key_layout = QHBoxLayout()
        self.api_key = QLineEdit()
        self.api_key.setPlaceholderText("Enter your Google Places API key")
        self.api_key.setText(self.config.get('google_places_api_key', ''))
        self.api_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_key.setToolTip(
            "Google Places API key for business name lookups.\n\n"
            "Get a key from: console.cloud.google.com\n"
            "Enable: Places API, Geocoding API, Routes API\n\n"
            "Leave blank to use free OpenStreetMap lookups instead."
        )
        api_key_layout.addWidget(self.api_key)

        # Show/hide toggle
        self.show_key_btn = QPushButton("Show")
        self.show_key_btn.setFixedWidth(60)
        self.show_key_btn.setCheckable(True)
        self.show_key_btn.clicked.connect(self._toggle_api_key_visibility)
        api_key_layout.addWidget(self.show_key_btn)

        api_layout.addLayout(api_key_layout)

        # API key help text
        api_help = QLabel(
            "The Google Places API provides accurate business name lookups.\n"
            "Without it, the app uses OpenStreetMap (free but less accurate)."
        )
        api_help.setStyleSheet("color: #666; font-size: 11px;")
        api_help.setWordWrap(True)
        api_layout.addWidget(api_help)

        # Test API button
        test_layout = QHBoxLayout()
        test_layout.addStretch()
        self.test_api_btn = QPushButton("Test API Key")
        self.test_api_btn.setToolTip("Verify that your API key is valid and working")
        self.test_api_btn.clicked.connect(self._test_api_key)
        test_layout.addWidget(self.test_api_btn)
        api_layout.addLayout(test_layout)

        self.api_status = QLabel("")
        self.api_status.setWordWrap(True)
        api_layout.addWidget(self.api_status)

        scroll_layout.addWidget(api_group)

        # === CATEGORIZATION THRESHOLDS SECTION ===
        thresh_group = QGroupBox("Categorization Thresholds")
        thresh_group.setStyleSheet(addr_group.styleSheet())
        thresh_layout = QVBoxLayout(thresh_group)
        thresh_layout.setSpacing(12)

        # Business distance threshold
        biz_dist_layout = QHBoxLayout()
        biz_dist_label = QLabel("Business trip minimum distance:")
        biz_dist_label.setToolTip("Trips longer than this on weekdays are considered business")
        biz_dist_layout.addWidget(biz_dist_label)
        biz_dist_layout.addStretch()

        self.business_distance = QDoubleSpinBox()
        self.business_distance.setRange(1.0, 50.0)
        self.business_distance.setValue(self.config.get('business_distance_threshold', 8.0))
        self.business_distance.setSuffix(" miles")
        self.business_distance.setToolTip(
            "Trips exceeding this distance on weekdays\n"
            "are automatically categorized as business travel."
        )
        biz_dist_layout.addWidget(self.business_distance)
        thresh_layout.addLayout(biz_dist_layout)

        # Merge gap threshold
        merge_gap_layout = QHBoxLayout()
        merge_gap_label = QLabel("Merge stops shorter than:")
        merge_gap_label.setToolTip("Stops shorter than this are merged (red lights, etc.)")
        merge_gap_layout.addWidget(merge_gap_label)
        merge_gap_layout.addStretch()

        self.merge_gap = QDoubleSpinBox()
        self.merge_gap.setRange(1.0, 10.0)
        self.merge_gap.setValue(self.config.get('merge_gap_minutes', 3.0))
        self.merge_gap.setSuffix(" minutes")
        self.merge_gap.setToolTip(
            "Short stops (traffic lights, brief pauses) are merged\n"
            "into a single trip if shorter than this duration."
        )
        merge_gap_layout.addWidget(self.merge_gap)
        thresh_layout.addLayout(merge_gap_layout)

        # Micro-trip threshold
        micro_layout = QHBoxLayout()
        micro_label = QLabel("Micro-trip threshold:")
        micro_label.setToolTip("Trips shorter than this are flagged as potential GPS drift")
        micro_layout.addWidget(micro_label)
        micro_layout.addStretch()

        self.micro_threshold = QDoubleSpinBox()
        self.micro_threshold.setRange(0.05, 1.0)
        self.micro_threshold.setValue(self.config.get('micro_trip_threshold', 0.15))
        self.micro_threshold.setSuffix(" miles")
        self.micro_threshold.setSingleStep(0.05)
        self.micro_threshold.setToolTip(
            "Very short trips are flagged as possible GPS drift\n"
            "or parking lot movements for your review."
        )
        micro_layout.addWidget(self.micro_threshold)
        thresh_layout.addLayout(micro_layout)

        scroll_layout.addWidget(thresh_group)

        scroll_layout.addStretch()
        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)

        # === BUTTONS ===
        button_layout = QHBoxLayout()

        # Reset to defaults button
        reset_btn = QPushButton("Reset to Defaults")
        reset_btn.setToolTip("Reset all settings to their default values")
        reset_btn.clicked.connect(self._reset_to_defaults)
        reset_btn.setStyleSheet("background-color: #757575;")
        button_layout.addWidget(reset_btn)

        button_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        cancel_btn.setStyleSheet("background-color: #757575;")
        button_layout.addWidget(cancel_btn)

        save_btn = QPushButton("Save Settings")
        save_btn.setDefault(True)
        save_btn.clicked.connect(self._save_settings)
        button_layout.addWidget(save_btn)

        layout.addLayout(button_layout)

    def _toggle_api_key_visibility(self):
        """Toggle API key visibility"""
        if self.show_key_btn.isChecked():
            self.api_key.setEchoMode(QLineEdit.EchoMode.Normal)
            self.show_key_btn.setText("Hide")
        else:
            self.api_key.setEchoMode(QLineEdit.EchoMode.Password)
            self.show_key_btn.setText("Show")

    def _test_api_key(self):
        """Test the API key by making a simple request"""
        api_key = self.api_key.text().strip()
        if not api_key:
            self.api_status.setText("No API key entered.")
            self.api_status.setStyleSheet("color: #ff9800;")
            return

        self.api_status.setText("Testing API key...")
        self.api_status.setStyleSheet("color: #666;")
        self.test_api_btn.setEnabled(False)
        QApplication.processEvents()

        try:
            import googlemaps
            client = googlemaps.Client(key=api_key)
            # Try a simple geocode request
            result = client.geocode("Seattle, WA")
            if result:
                self.api_status.setText("API key is valid and working!")
                self.api_status.setStyleSheet("color: #388e3c; font-weight: bold;")
            else:
                self.api_status.setText("API key accepted but returned no results.")
                self.api_status.setStyleSheet("color: #ff9800;")
        except googlemaps.exceptions.ApiError as e:
            self.api_status.setText(f"API Error: {str(e)}")
            self.api_status.setStyleSheet("color: #d32f2f;")
        except Exception as e:
            self.api_status.setText(f"Error: {str(e)}")
            self.api_status.setStyleSheet("color: #d32f2f;")
        finally:
            self.test_api_btn.setEnabled(True)

    def _reset_to_defaults(self):
        """Reset all settings to default values"""
        reply = QMessageBox.question(
            self, "Reset Settings",
            "Are you sure you want to reset all settings to their default values?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.home_address.clear()
            self.work_address.clear()
            self.api_key.clear()
            self.business_distance.setValue(8.0)
            self.merge_gap.setValue(3.0)
            self.micro_threshold.setValue(0.15)
            self.api_status.clear()

    def _save_settings(self):
        """Save settings to config file"""
        # Build config dict
        new_config = {
            'home_address': self.home_address.text().strip(),
            'work_address': self.work_address.text().strip(),
            'google_places_api_key': self.api_key.text().strip(),
            'business_distance_threshold': self.business_distance.value(),
            'merge_gap_minutes': self.merge_gap.value(),
            'micro_trip_threshold': self.micro_threshold.value()
        }

        # Preserve any other existing config values
        for key, value in self.config.items():
            if key not in new_config:
                new_config[key] = value

        try:
            # Create backup before saving
            if os.path.exists(self.config_file):
                import shutil
                shutil.copy2(self.config_file, self.config_file + '.bak')

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(new_config, f, indent=2)

            self.settings_changed.emit()
            QMessageBox.information(
                self, "Settings Saved",
                "Settings have been saved successfully.\n\n"
                "Some changes may require re-running the analysis to take effect."
            )
            self.accept()
        except Exception as e:
            QMessageBox.critical(
                self, "Error Saving Settings",
                f"Could not save settings:\n{str(e)}"
            )


class NoteDialog(QDialog):
    """Dialog for editing trip notes with a multi-line text area"""

    def __init__(self, trip_info: str, current_note: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Trip Note")
        self.setMinimumSize(500, 300)
        self.resize(600, 350)

        layout = QVBoxLayout(self)

        # Trip info label
        info_label = QLabel(f"Note for trip:\n{trip_info}")
        info_label.setStyleSheet("color: #666; padding: 5px; background: #f5f5f5; border-radius: 3px;")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        # Multi-line text edit
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(current_note)
        self.text_edit.setPlaceholderText("Enter your notes here...")
        self.text_edit.setMinimumHeight(150)
        layout.addWidget(self.text_edit)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        save_btn = QPushButton("Save")
        save_btn.setDefault(True)
        save_btn.clicked.connect(self.accept)
        button_layout.addWidget(save_btn)

        layout.addLayout(button_layout)

    def get_note(self) -> str:
        return self.text_edit.toPlainText()


class JsonEditorDialog(QMainWindow):
    """Dialog for viewing and editing business mappings with Category support"""

    data_saved = pyqtSignal()

    def __init__(self, file_path: str, title: str, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.data = {}
        self.setWindowTitle(title)
        self.setMinimumSize(900, 600)
        self._setup_ui()
        self._load_data()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Toolbar
        toolbar = QHBoxLayout()

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search...")
        self.search_box.textChanged.connect(self._filter_table)
        toolbar.addWidget(QLabel("Search:"))
        toolbar.addWidget(self.search_box)

        toolbar.addStretch()

        # Filter by source
        toolbar.addWidget(QLabel("Show:"))
        self.source_filter = QComboBox()
        self.source_filter.addItems(["All", "Manual Only", "API Only", "Unresolved"])
        self.source_filter.currentTextChanged.connect(self._filter_table)
        toolbar.addWidget(self.source_filter)

        toolbar.addStretch()

        add_btn = QPushButton("Add Entry")
        add_btn.clicked.connect(self._add_entry)
        toolbar.addWidget(add_btn)

        delete_btn = QPushButton("Delete Selected")
        delete_btn.clicked.connect(self._delete_selected)
        toolbar.addWidget(delete_btn)

        # Re-lookup button for API entries
        relookup_btn = QPushButton("Re-lookup Selected")
        relookup_btn.setToolTip("Clear selected API entries to allow re-lookup")
        relookup_btn.clicked.connect(self._relookup_selected)
        toolbar.addWidget(relookup_btn)

        layout.addLayout(toolbar)

        # Table with 4 columns: Address, Business Name, Category, Source
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Address", "Business Name", "Category", "Source"])
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 300)
        self.table.setColumnWidth(1, 200)
        self.table.setColumnWidth(2, 100)
        self.table.setColumnWidth(3, 100)

        self.table.cellChanged.connect(self._on_cell_changed)
        layout.addWidget(self.table)

        # Stats label
        self.stats_label = QLabel()
        layout.addWidget(self.stats_label)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        save_btn = QPushButton("Save Changes")
        save_btn.clicked.connect(self._save_data)
        btn_layout.addWidget(save_btn)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _load_data(self):
        """Load JSON data from file"""
        self.data = {}
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
            except Exception as e:
                QMessageBox.warning(self, "Load Error", f"Failed to load file:\n{e}")

        self._populate_table()

    def _populate_table(self):
        """Populate table with data"""
        self.table.blockSignals(True)

        # Filter out comment entries (keys starting with _)
        filtered_data = {k: v for k, v in self.data.items() if not k.startswith('_')}
        self.table.setRowCount(len(filtered_data))

        for row, (key, value) in enumerate(sorted(filtered_data.items())):
            # Address
            key_item = QTableWidgetItem(str(key))
            self.table.setItem(row, 0, key_item)

            # Handle different value types
            if isinstance(value, dict):
                name = value.get('name', '')
                category = value.get('category', '')
                source = value.get('source', 'manual')
            else:
                # Old format: just a string
                name = str(value)
                category = ''
                source = 'manual'

            # Business Name
            name_item = QTableWidgetItem(name)
            if name == "NO_BUSINESS_FOUND":
                name_item.setForeground(QColor('#999999'))
            self.table.setItem(row, 1, name_item)

            # Category (use combo box)
            category_combo = QComboBox()
            category_combo.addItems(["", "Business", "Personal", "Commute"])
            if category:
                idx = category_combo.findText(category, Qt.MatchFlag.MatchFixedString)
                if idx >= 0:
                    category_combo.setCurrentIndex(idx)
            self.table.setCellWidget(row, 2, category_combo)

            # Source (read-only display)
            source_item = QTableWidgetItem(source)
            source_item.setFlags(source_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            if source in ['google_api', 'osm_api']:
                source_item.setForeground(QColor('#2196F3'))
            else:
                source_item.setForeground(QColor('#4CAF50'))
            self.table.setItem(row, 3, source_item)

        self.table.blockSignals(False)
        self._update_stats()

    def _update_stats(self):
        """Update the stats label"""
        visible = sum(1 for row in range(self.table.rowCount()) if not self.table.isRowHidden(row))
        total = self.table.rowCount()
        self.stats_label.setText(f"Showing {visible} of {total} entries")

    def _filter_table(self, text: str = None):
        """Filter table rows based on search text and source filter"""
        search_text = self.search_box.text().lower()
        source_filter = self.source_filter.currentText()

        for row in range(self.table.rowCount()):
            key_item = self.table.item(row, 0)
            name_item = self.table.item(row, 1)
            source_item = self.table.item(row, 3)

            key_text = key_item.text().lower() if key_item else ""
            name_text = name_item.text().lower() if name_item else ""
            source_text = source_item.text() if source_item else ""

            # Text search filter
            text_match = not search_text or (search_text in key_text or search_text in name_text)

            # Source filter
            if source_filter == "Manual Only":
                source_match = source_text == "manual"
            elif source_filter == "API Only":
                source_match = source_text in ["google_api", "osm_api"]
            elif source_filter == "Unresolved":
                source_match = name_text == "no_business_found"
            else:
                source_match = True

            self.table.setRowHidden(row, not (text_match and source_match))

        self._update_stats()

    def _on_cell_changed(self, row: int, col: int):
        """Handle cell edit"""
        # Mark as modified (could add visual indicator)
        pass

    def _add_entry(self):
        """Add a new entry"""
        self.table.blockSignals(True)
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem(""))
        self.table.setItem(row, 1, QTableWidgetItem(""))

        # Add category combo
        category_combo = QComboBox()
        category_combo.addItems(["", "Business", "Personal", "Commute"])
        self.table.setCellWidget(row, 2, category_combo)

        # Source is "manual" for new entries
        source_item = QTableWidgetItem("manual")
        source_item.setFlags(source_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        source_item.setForeground(QColor('#4CAF50'))
        self.table.setItem(row, 3, source_item)

        self.table.blockSignals(False)
        self.table.scrollToItem(self.table.item(row, 0))
        self.table.editItem(self.table.item(row, 0))
        self._update_stats()

    def _delete_selected(self):
        """Delete selected rows"""
        rows = set(index.row() for index in self.table.selectedIndexes())
        if not rows:
            return

        reply = QMessageBox.question(
            self, "Confirm Delete",
            f"Delete {len(rows)} selected entries?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            for row in sorted(rows, reverse=True):
                self.table.removeRow(row)
            self._update_stats()

    def _relookup_selected(self):
        """Mark selected API entries for re-lookup by deleting them"""
        rows = set(index.row() for index in self.table.selectedIndexes())
        if not rows:
            QMessageBox.information(self, "No Selection", "Please select entries to re-lookup.")
            return

        # Filter to only API entries
        api_rows = []
        for row in rows:
            source_item = self.table.item(row, 3)
            if source_item and source_item.text() in ['google_api', 'osm_api']:
                api_rows.append(row)

        if not api_rows:
            QMessageBox.information(self, "No API Entries",
                "None of the selected entries are API lookups.\n"
                "Only API lookup entries can be marked for re-lookup.")
            return

        reply = QMessageBox.question(
            self, "Confirm Re-lookup",
            f"Remove {len(api_rows)} API entries to allow re-lookup?\n\n"
            "Manual entries will not be affected.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            for row in sorted(api_rows, reverse=True):
                self.table.removeRow(row)
            self._update_stats()
            QMessageBox.information(self, "Done",
                f"Removed {len(api_rows)} entries.\n"
                "Run analysis with lookup to re-fetch business names.")

    def _save_data(self):
        """Save data back to JSON file"""
        new_data = {}

        for row in range(self.table.rowCount()):
            key_item = self.table.item(row, 0)
            name_item = self.table.item(row, 1)
            category_combo = self.table.cellWidget(row, 2)
            source_item = self.table.item(row, 3)

            if key_item and key_item.text().strip():
                key = key_item.text().strip()
                name = name_item.text().strip() if name_item else ""
                category = category_combo.currentText() if category_combo else ""
                source = source_item.text() if source_item else "manual"

                # Build entry in new format
                entry = {'name': name, 'source': source}
                if category:
                    entry['category'] = category

                new_data[key] = entry

        try:
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump(new_data, f, indent=2, ensure_ascii=False)

            self.data = new_data
            QMessageBox.information(self, "Saved", f"Saved {len(new_data)} entries to:\n{os.path.basename(self.file_path)}")
            self.data_saved.emit()
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save:\n{e}")


class UnifiedTripView(QWidget):
    """Unified view for trips - grouped by destination with filtering"""

    trip_selected = pyqtSignal(dict)  # Emits when a trip/destination is selected
    trip_updated = pyqtSignal(dict, str, str)  # trip, field, new_value
    mapping_changed = pyqtSignal()  # Emits when business mapping should be saved
    show_daily_journey = pyqtSignal(list)  # list of trips for a day

    def __init__(self, parent=None):
        super().__init__(parent)
        self.trips_data = []  # All trips
        self.grouped_data = []  # Trips grouped by destination
        self.day_grouped_data = []  # Trips grouped by date
        self.view_mode = "grouped"  # "grouped", "individual", or "by_day"
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        # Filter bar - three rows for better spacing
        filter_frame = QFrame()
        filter_frame.setFrameStyle(QFrame.Shape.StyledPanel)
        filter_vlayout = QVBoxLayout(filter_frame)
        filter_vlayout.setContentsMargins(8, 4, 8, 4)
        filter_vlayout.setSpacing(4)

        # Row 1: View mode and Category
        row1 = QHBoxLayout()
        row1.setSpacing(8)

        row1.addWidget(QLabel("View:"))
        self.view_mode_combo = QComboBox()
        self.view_mode_combo.addItems(["By Destination", "All Trips", "By Day", "By Week"])
        self.view_mode_combo.setMinimumWidth(120)
        self.view_mode_combo.setToolTip(
            "Change how trips are displayed:\n\n"
            "• By Destination - Group trips by end address (best for categorizing)\n"
            "• All Trips - Show each trip individually with full details\n"
            "• By Day - Expandable tree grouped by date\n"
            "• By Week - Expandable tree grouped by week with daily breakdown"
        )
        self.view_mode_combo.currentTextChanged.connect(self._on_view_mode_changed)
        row1.addWidget(self.view_mode_combo)

        row1.addWidget(QLabel("Category:"))
        self.category_filter = QComboBox()
        self.category_filter.addItems(["All", "Business", "Personal", "Commute"])
        self.category_filter.setMinimumWidth(90)
        self.category_filter.setToolTip(
            "Filter trips by category:\n\n"
            "• All - Show all trips regardless of category\n"
            "• Business - Only show business-related trips (tax deductible)\n"
            "• Personal - Only show personal trips\n"
            "• Commute - Only show home-to-office commute trips"
        )
        self.category_filter.currentTextChanged.connect(self._apply_filters)
        row1.addWidget(self.category_filter)

        row1.addWidget(QLabel("Status:"))
        self.status_filter = QComboBox()
        self.status_filter.addItems(["All", "Unresolved", "Resolved", "Unconfirmed", "Duplicates"])
        self.status_filter.setMinimumWidth(100)
        self.status_filter.setToolTip(
            "Filter by resolution status:\n\n"
            "• All - Show all destinations\n"
            "• Unresolved - Destinations that need a business name assigned\n"
            "• Resolved - Destinations with confirmed business names\n"
            "• Unconfirmed - Business trips without confirmed names"
        )
        self.status_filter.currentTextChanged.connect(self._apply_filters)
        row1.addWidget(self.status_filter)

        row1.addStretch()

        # Stats label on right of row 1
        self.stats_label = QLabel("0 items")
        self.stats_label.setStyleSheet("font-weight: bold; color: #555;")
        row1.addWidget(self.stats_label)

        filter_vlayout.addLayout(row1)

        # Row 2: Business filter, Search, and Micro-trips checkbox
        row2 = QHBoxLayout()
        row2.setSpacing(8)

        row2.addWidget(QLabel("Business:"))
        self.business_filter = QComboBox()
        self.business_filter.addItem("All")
        self.business_filter.setMinimumWidth(150)
        self.business_filter.setToolTip(
            "Filter trips by business name.\n\n"
            "Shows only trips to/from the selected business.\n"
            "The list is populated from your saved business mappings."
        )
        self.business_filter.currentTextChanged.connect(self._apply_filters)
        row2.addWidget(self.business_filter, 1)

        row2.addWidget(QLabel("Search:"))
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search...")
        self.search_box.setMinimumWidth(120)
        self.search_box.setToolTip(
            "Search trips by text.\n\n"
            "Searches across addresses, business names, and other fields.\n"
            "Results update as you type."
        )
        self.search_box.textChanged.connect(self._apply_filters)
        row2.addWidget(self.search_box, 1)

        # Micro-trip filter checkbox (moved to row 2)
        self.hide_micro_trips = QCheckBox("Hide Micro-trips")
        self.hide_micro_trips.setToolTip(
            "Hide very short trips that are likely GPS drift or parking adjustments.\n\n"
            "Micro-trips are trips under 0.15 miles that start and end on the same\n"
            "street. These are often caused by GPS inaccuracy or vehicle repositioning.\n\n"
            "Uncheck to see all trips including micro-trips (marked with ⚠)."
        )
        self.hide_micro_trips.stateChanged.connect(self._apply_filters)
        row2.addWidget(self.hide_micro_trips)

        filter_vlayout.addLayout(row2)

        layout.addWidget(filter_frame)

        # Main table
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.table.setSortingEnabled(True)
        self.table.setToolTip(
            "Trip data table.\n\n"
            "• Click a column header to sort\n"
            "• Double-click a row to view on map\n"
            "• Right-click for options (categorize, rename, etc.)\n"
            "• Hold Ctrl/Shift to select multiple rows\n"
            "• Drag column borders to resize"
        )

        # Context menu
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)

        layout.addWidget(self.table)

        # Tree widget for By Day view (expandable)
        self.tree = QTreeWidget()
        self.tree.setAlternatingRowColors(True)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.tree.setToolTip(
            "Expandable trip tree view.\n\n"
            "• Click the arrow or double-click to expand/collapse\n"
            "• Click a trip to view it on the map\n"
            "• Right-click for options (categorize, etc.)\n"
            "• Green = Business miles, Orange = Personal miles"
        )
        self.tree.setHeaderLabels(["Date / Trip", "Trips / Category", "Business", "Personal", "Total / Destination"])
        self.tree.setColumnWidth(0, 180)
        self.tree.setColumnWidth(1, 120)
        self.tree.setColumnWidth(2, 100)
        self.tree.setColumnWidth(3, 100)
        self.tree.header().setStretchLastSection(True)
        self.tree.itemClicked.connect(self._on_tree_item_clicked)
        self.tree.itemDoubleClicked.connect(self._on_tree_item_double_clicked)
        self.tree.itemExpanded.connect(self._on_tree_item_expanded)
        self.tree.itemSelectionChanged.connect(self._on_tree_selection_changed)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._show_tree_context_menu)
        self.tree.hide()  # Hidden by default
        layout.addWidget(self.tree)

        # Set up for grouped view initially
        self._setup_grouped_columns()

    def _setup_grouped_columns(self):
        """Set up columns for grouped (by destination) view"""
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Business Name", "Category", "Trips", "Miles", "Status", "Destination Address"
        ])
        header = self.table.horizontalHeader()
        for i in range(6):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 200)  # Business Name
        self.table.setColumnWidth(1, 85)   # Category
        self.table.setColumnWidth(2, 50)   # Trips
        self.table.setColumnWidth(3, 70)   # Miles
        self.table.setColumnWidth(4, 110)  # Status
        self.table.setColumnWidth(5, 280)  # Destination Address

    def _setup_individual_columns(self):
        """Set up columns for individual trips view"""
        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels([
            "Date", "Day", "Start", "End", "Category", "Reason", "Distance", "From", "To", "Business Name", "Notes"
        ])
        header = self.table.horizontalHeader()
        for i in range(11):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 95)   # Date
        self.table.setColumnWidth(1, 50)   # Day
        self.table.setColumnWidth(2, 55)   # Start time
        self.table.setColumnWidth(3, 55)   # End time
        self.table.setColumnWidth(4, 85)   # Category
        self.table.setColumnWidth(5, 140)  # Reason
        self.table.setColumnWidth(6, 70)   # Distance
        self.table.setColumnWidth(7, 180)  # From
        self.table.setColumnWidth(8, 180)  # To
        self.table.setColumnWidth(9, 150)  # Business Name
        self.table.setColumnWidth(10, 180) # Notes

    def _setup_by_day_columns(self):
        """Set up columns for by-day view"""
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Date", "Day", "Trips", "Miles", "Business Miles", "Actions"
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 100)
        self.table.setColumnWidth(1, 50)
        self.table.setColumnWidth(2, 50)
        self.table.setColumnWidth(3, 80)
        self.table.setColumnWidth(4, 100)

    def _setup_weekly_columns(self):
        """Set up columns for weekly view"""
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Week Starting", "Trips", "Business", "Personal", "Commute", "Total", "Bus %"
        ])
        header = self.table.horizontalHeader()
        for i in range(7):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 100)  # Week
        self.table.setColumnWidth(1, 50)   # Trips
        self.table.setColumnWidth(2, 80)   # Business
        self.table.setColumnWidth(3, 80)   # Personal
        self.table.setColumnWidth(4, 80)   # Commute
        self.table.setColumnWidth(5, 80)   # Total
        self.table.setColumnWidth(6, 60)   # %

    def _on_view_mode_changed(self, mode: str):
        """Handle view mode change"""
        if mode == "By Destination":
            self.view_mode = "grouped"
            self.table.show()
            self.tree.hide()
            self._setup_grouped_columns()
        elif mode == "By Day":
            self.view_mode = "by_day"
            self.table.hide()
            self.tree.show()
            self._setup_by_day_tree_columns()
        elif mode == "By Week":
            self.view_mode = "weekly"
            self.table.hide()
            self.tree.show()
            self._setup_weekly_tree_columns()
        else:
            self.view_mode = "individual"
            self.table.show()
            self.tree.hide()
            self._setup_individual_columns()
        self._refresh_table()

    def _setup_by_day_tree_columns(self):
        """Set up tree columns for By Day view"""
        self.tree.setHeaderLabels(["Date / Trip", "Trips / Category", "Business", "Personal", "Total / Destination"])
        self.tree.setColumnWidth(0, 180)
        self.tree.setColumnWidth(1, 120)
        self.tree.setColumnWidth(2, 100)
        self.tree.setColumnWidth(3, 100)
        self.tree.header().setStretchLastSection(True)

    def _setup_weekly_tree_columns(self):
        """Set up tree columns for By Week view"""
        self.tree.setHeaderLabels(["Date / Trip", "Trips / Category", "Business", "Personal", "Total / Destination"])
        self.tree.setColumnWidth(0, 180)
        self.tree.setColumnWidth(1, 120)
        self.tree.setColumnWidth(2, 100)
        self.tree.setColumnWidth(3, 100)
        self.tree.header().setStretchLastSection(True)

    def load_trips(self, trips: List[Dict]):
        """Load trip data and group by destination"""
        # Clear existing data first to free memory
        self.trips_data = []
        self.grouped_data = []
        self.day_grouped_data = []
        self.table.setRowCount(0)
        self.tree.clear()

        # Now load new data
        self.trips_data = trips

        # Group trips by destination address
        dest_groups = {}
        for trip in trips:
            dest = trip.get('end_address', '').strip()
            if not dest:
                continue
            if dest not in dest_groups:
                dest_groups[dest] = {
                    'address': dest,
                    'trips': [],
                    'total_miles': 0,
                    'business_name': trip.get('business_name', ''),
                    'categories': set(),
                    'lat': trip.get('end_lat'),
                    'lng': trip.get('end_lng')
                }
            dest_groups[dest]['trips'].append(trip)
            dest_groups[dest]['total_miles'] += trip.get('distance', 0)
            dest_groups[dest]['categories'].add(trip.get('computed_category', 'PERSONAL'))
            # Use the most recent business name
            if trip.get('business_name'):
                dest_groups[dest]['business_name'] = trip.get('business_name')

        # Convert to list and add computed fields
        self.grouped_data = []
        for addr, data in dest_groups.items():
            # Determine primary category (most common)
            cat_counts = {}
            for t in data['trips']:
                cat = t.get('computed_category', 'PERSONAL')
                cat_counts[cat] = cat_counts.get(cat, 0) + 1
            primary_cat = max(cat_counts, key=cat_counts.get) if cat_counts else 'PERSONAL'

            # Determine status
            business_name = data['business_name']
            if business_name and business_name not in ['', 'Unknown', 'NO_BUSINESS_FOUND']:
                status = 'Has Name'
            elif primary_cat == 'BUSINESS':
                status = 'Unconfirmed Business'
            else:
                status = 'Needs Name'

            # Extract street name
            street = self._extract_street(addr)

            self.grouped_data.append({
                'address': addr,
                'business_name': business_name if business_name not in ['NO_BUSINESS_FOUND'] else '',
                'primary_category': primary_cat,
                'trip_count': len(data['trips']),
                'total_miles': data['total_miles'],
                'status': status,
                'street': street,
                'lat': data['lat'],
                'lng': data['lng'],
                'trips': data['trips']
            })

        # Group trips by date
        day_groups = {}
        for trip in trips:
            trip_date = trip.get('started')
            if not trip_date or not hasattr(trip_date, 'date'):
                continue
            date_key = trip_date.date()
            if date_key not in day_groups:
                day_groups[date_key] = {
                    'date': date_key,
                    'trips': [],
                    'total_miles': 0,
                    'business_miles': 0,
                    'personal_miles': 0,
                    'commute_miles': 0
                }
            day_groups[date_key]['trips'].append(trip)
            distance = trip.get('distance', 0)
            day_groups[date_key]['total_miles'] += distance
            cat = trip.get('computed_category', 'PERSONAL')
            if cat == 'BUSINESS':
                day_groups[date_key]['business_miles'] += distance
            elif cat == 'PERSONAL':
                day_groups[date_key]['personal_miles'] += distance
            elif cat == 'COMMUTE':
                day_groups[date_key]['commute_miles'] += distance

        # Convert to sorted list (most recent first)
        self.day_grouped_data = []
        for date_key in sorted(day_groups.keys(), reverse=True):
            data = day_groups[date_key]
            # Sort trips within each day by time
            data['trips'] = sorted(data['trips'], key=lambda t: t.get('started'))
            self.day_grouped_data.append(data)

        # Group trips by week
        week_groups = {}
        for trip in trips:
            trip_date = trip.get('started')
            if not trip_date or not hasattr(trip_date, 'date'):
                continue
            # Get week start (Monday)
            week_start = trip_date - timedelta(days=trip_date.weekday())
            week_key = week_start.strftime('%Y-%m-%d')
            if week_key not in week_groups:
                week_groups[week_key] = {
                    'week_start': week_start,
                    'trips': [],
                    'total_miles': 0,
                    'business_miles': 0,
                    'personal_miles': 0,
                    'commute_miles': 0
                }
            week_groups[week_key]['trips'].append(trip)
            distance = trip.get('distance', 0)
            week_groups[week_key]['total_miles'] += distance
            cat = trip.get('computed_category', 'PERSONAL')
            if cat == 'BUSINESS':
                week_groups[week_key]['business_miles'] += distance
            elif cat == 'PERSONAL':
                week_groups[week_key]['personal_miles'] += distance
            elif cat == 'COMMUTE':
                week_groups[week_key]['commute_miles'] += distance

        # Convert to sorted list (most recent first)
        self.week_grouped_data = []
        for week_key in sorted(week_groups.keys(), reverse=True):
            data = week_groups[week_key]
            data['week_key'] = week_key
            self.week_grouped_data.append(data)

        # Update business filter dropdown
        self._update_business_filter()
        self._refresh_table()

    def _extract_street(self, address: str) -> str:
        """Extract street name from address"""
        import re
        parts = address.split(',')
        if parts:
            street_part = parts[0].strip()
            match = re.match(r'^\d+\s+(.+)$', street_part)
            if match:
                return match.group(1)
            return street_part
        return address

    def _update_business_filter(self):
        """Update the business name filter dropdown"""
        self.business_filter.blockSignals(True)
        current = self.business_filter.currentText()
        self.business_filter.clear()
        self.business_filter.addItem("All")
        self.business_filter.addItem("(No Name)")

        names = set()
        for g in self.grouped_data:
            name = g.get('business_name', '')
            if name and name not in ['', 'Unknown', 'NO_BUSINESS_FOUND']:
                names.add(name)

        for name in sorted(names):
            self.business_filter.addItem(name)

        # Restore selection if possible
        idx = self.business_filter.findText(current)
        if idx >= 0:
            self.business_filter.setCurrentIndex(idx)
        self.business_filter.blockSignals(False)

    def _refresh_table(self):
        """Refresh the table based on current view mode and filters"""
        if self.view_mode == "by_day":
            self._populate_by_day_tree()
            self._apply_filters()
            return

        if self.view_mode == "weekly":
            self._populate_weekly_tree()
            self._apply_filters()
            return

        self.table.setSortingEnabled(False)

        if self.view_mode == "grouped":
            self._populate_grouped_view()
        else:
            self._populate_individual_view()

        self.table.setSortingEnabled(True)
        self._apply_filters()

    def _populate_grouped_view(self):
        """Populate table with grouped destination data"""
        self.table.setRowCount(len(self.grouped_data))

        for row, data in enumerate(self.grouped_data):
            # Column 0: Business name (most important - what is this place?)
            name = data.get('business_name', '')
            name_item = QTableWidgetItem(name)
            name_item.setData(Qt.ItemDataRole.UserRole, row)  # Store index for selection
            if data['status'] == 'Unconfirmed Business':
                name_item.setText('[Needs Confirmation]')
                name_item.setForeground(QColor('#ff8f00'))
                name_item.setFont(QFont('', -1, -1, True))
            elif data['status'] == 'Needs Name':
                name_item.setText('[Unknown]')
                name_item.setForeground(QColor('#999999'))
                name_item.setFont(QFont('', -1, -1, True))
            self.table.setItem(row, 0, name_item)

            # Column 1: Category
            cat = data['primary_category']
            cat_item = QTableWidgetItem(cat)
            if cat == 'BUSINESS':
                cat_item.setBackground(QColor('#e8f5e9'))
                cat_item.setForeground(QColor('#2e7d32'))
            elif cat == 'PERSONAL':
                cat_item.setBackground(QColor('#fff3e0'))
                cat_item.setForeground(QColor('#e65100'))
            elif cat == 'COMMUTE':
                cat_item.setBackground(QColor('#e3f2fd'))
                cat_item.setForeground(QColor('#1565c0'))
            self.table.setItem(row, 1, cat_item)

            # Column 2: Trip count
            trip_count = data['trip_count']
            count_item = NumericTableWidgetItem(str(trip_count), trip_count)
            self.table.setItem(row, 2, count_item)

            # Column 3: Total miles
            total_miles = data['total_miles']
            miles_item = NumericTableWidgetItem(f"{total_miles:.1f}", total_miles)
            self.table.setItem(row, 3, miles_item)

            # Column 4: Status
            status = data['status']
            if status == 'Needs Name':
                status_text = 'Needs Name'
            elif status == 'Unconfirmed Business':
                status_text = 'Unconfirmed'
            else:
                status_text = 'Confirmed'
            status_item = QTableWidgetItem(status_text)
            if status == 'Needs Name':
                status_item.setForeground(QColor('#d32f2f'))
            elif status == 'Unconfirmed Business':
                status_item.setForeground(QColor('#ff8f00'))
            else:
                status_item.setForeground(QColor('#388e3c'))
            self.table.setItem(row, 4, status_item)

            # Column 5: Destination address
            addr_item = QTableWidgetItem(data['address'])
            self.table.setItem(row, 5, addr_item)

    def _populate_individual_view(self):
        """Populate table with individual trip data"""
        self.table.setRowCount(len(self.trips_data))

        for row, trip in enumerate(self.trips_data):
            # Date - with merge/micro/duplicate indicator if applicable
            date_str = trip['started'].strftime('%Y-%m-%d')
            tooltip = None
            if trip.get('is_duplicate'):
                date_str = f"⚡ {date_str}"
                tooltip = "POTENTIAL DUPLICATE: This trip has the same start time and destination as another trip"
            elif trip.get('is_micro_trip'):
                date_str = f"⚠ {date_str}"
                tooltip = f"MICRO-TRIP: {trip.get('micro_reason', 'Very short distance')}\nRight-click for options"
            elif trip.get('is_merged'):
                merge_count = trip.get('merge_count', 2)
                date_str = f"⟨{merge_count}⟩ {date_str}"
                tooltip = f"This trip was merged from {trip.get('merge_count', 2)} short segments (red lights/traffic stops)"
            date_item = QTableWidgetItem(date_str)
            date_item.setData(Qt.ItemDataRole.UserRole, row)
            if tooltip:
                date_item.setToolTip(tooltip)
            if trip.get('is_duplicate'):
                date_item.setForeground(QColor('#d32f2f'))  # Red for duplicates
            elif trip.get('is_micro_trip'):
                date_item.setForeground(QColor('#ff9800'))  # Orange for micro-trips
            self.table.setItem(row, 0, date_item)

            # Day
            self.table.setItem(row, 1, QTableWidgetItem(trip['started'].strftime('%a')))

            # Start time
            start_time = trip['started'].strftime('%H:%M')
            self.table.setItem(row, 2, QTableWidgetItem(start_time))

            # End time - parse from 'stopped' field
            stopped = trip.get('stopped', '')
            end_time = ''
            if stopped:
                try:
                    # Handle various date formats
                    from datetime import datetime
                    for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%m/%d/%Y %H:%M', '%m/%d/%Y %I:%M %p']:
                        try:
                            end_dt = datetime.strptime(stopped, fmt)
                            end_time = end_dt.strftime('%H:%M')
                            break
                        except ValueError:
                            continue
                except Exception:
                    end_time = ''
            self.table.setItem(row, 3, QTableWidgetItem(end_time))

            # Category
            cat = trip.get('computed_category', 'PERSONAL')
            cat_item = QTableWidgetItem(cat)
            if cat == 'BUSINESS':
                cat_item.setBackground(QColor('#e8f5e9'))
                cat_item.setForeground(QColor('#2e7d32'))
            elif cat == 'PERSONAL':
                cat_item.setBackground(QColor('#fff3e0'))
                cat_item.setForeground(QColor('#e65100'))
            elif cat == 'COMMUTE':
                cat_item.setBackground(QColor('#e3f2fd'))
                cat_item.setForeground(QColor('#1565c0'))
            self.table.setItem(row, 4, cat_item)

            # Category Reason
            reason = trip.get('category_reason', '')
            reason_item = QTableWidgetItem(reason)
            reason_item.setForeground(QColor('#757575'))  # Gray text
            self.table.setItem(row, 5, reason_item)

            # Distance (use NumericTableWidgetItem for proper sorting)
            dist_val = trip.get('distance', 0)
            dist_item = NumericTableWidgetItem(f"{dist_val:.1f} mi", dist_val)
            self.table.setItem(row, 6, dist_item)

            # From/To
            self.table.setItem(row, 7, QTableWidgetItem(trip.get('start_address', '')))
            self.table.setItem(row, 8, QTableWidgetItem(trip.get('end_address', '')))

            # Business name
            name = trip.get('business_name', '')
            name_item = QTableWidgetItem(name)
            if cat == 'BUSINESS' and not name:
                name_item.setText('[Unconfirmed]')
                name_item.setForeground(QColor('#ff8f00'))
                name_item.setFont(QFont('', -1, -1, True))
            self.table.setItem(row, 9, name_item)

            # Notes - load from trip_notes.json
            notes = load_trip_notes()
            trip_key = get_trip_key(trip)
            note_text = notes.get(trip_key, '')
            note_item = QTableWidgetItem(note_text)
            if note_text:
                note_item.setForeground(QColor('#666666'))
            self.table.setItem(row, 10, note_item)

    def _populate_by_day_view(self):
        """Populate table with trips grouped by day"""
        self.table.setRowCount(len(self.day_grouped_data))

        for row, data in enumerate(self.day_grouped_data):
            date = data['date']
            
            # Date
            date_str = date.strftime('%Y-%m-%d')
            date_item = QTableWidgetItem(date_str)
            date_item.setData(Qt.ItemDataRole.UserRole, row)
            date_item.setFont(QFont('', -1, QFont.Weight.Bold.value))
            self.table.setItem(row, 0, date_item)

            # Day of week
            day_str = date.strftime('%a')
            day_item = QTableWidgetItem(day_str)
            # Color weekends differently
            if date.weekday() >= 5:  # Saturday or Sunday
                day_item.setForeground(QColor('#9c27b0'))
                day_item.setFont(QFont('', -1, QFont.Weight.Bold.value))
            self.table.setItem(row, 1, day_item)

            # Trip count
            count_item = QTableWidgetItem()
            count_item.setData(Qt.ItemDataRole.DisplayRole, len(data['trips']))
            self.table.setItem(row, 2, count_item)

            # Total miles (use NumericTableWidgetItem for proper sorting)
            total_miles = data['total_miles']
            miles_item = NumericTableWidgetItem(f"{total_miles:.1f}", total_miles)
            self.table.setItem(row, 3, miles_item)

            # Business miles (use NumericTableWidgetItem for proper sorting)
            biz_miles = data['business_miles']
            biz_item = NumericTableWidgetItem(f"{biz_miles:.1f}", biz_miles)
            if biz_miles > 0:
                biz_item.setForeground(QColor('#2e7d32'))
            self.table.setItem(row, 4, biz_item)

            # Actions hint
            hint_item = QTableWidgetItem("Double-click to view day's trips")
            hint_item.setForeground(QColor('#666666'))
            hint_item.setFont(QFont('', -1, -1, True))
            self.table.setItem(row, 5, hint_item)

    def _populate_weekly_tree(self):
        """Populate tree widget with expandable weeks containing days and trips"""
        self.tree.clear()

        if not hasattr(self, 'week_grouped_data'):
            self.week_grouped_data = []

        for week_data in self.week_grouped_data:
            week_key = week_data['week_key']
            trips = week_data['trips']
            total_miles = week_data['total_miles']
            biz_miles = week_data['business_miles']
            biz_pct = (biz_miles / total_miles * 100) if total_miles > 0 else 0

            # Create week item (top level)
            week_item = QTreeWidgetItem([
                f"Week of {week_key}",
                f"{len(trips)} trips",
                f"{biz_miles:.1f} mi",
                f"{week_data['personal_miles']:.1f} mi",
                f"{total_miles:.1f} mi ({biz_pct:.0f}% business)"
            ])
            week_item.setData(0, Qt.ItemDataRole.UserRole, {'type': 'week', 'data': week_data})
            week_item.setFont(0, QFont('', -1, QFont.Weight.Bold.value))
            week_item.setForeground(2, QColor('#2e7d32'))
            week_item.setForeground(3, QColor('#e65100'))

            # Group trips by day within this week
            day_groups = {}
            for trip in trips:
                trip_date = trip.get('started')
                if not trip_date:
                    continue
                date_key = trip_date.date() if hasattr(trip_date, 'date') else trip_date
                if date_key not in day_groups:
                    day_groups[date_key] = []
                day_groups[date_key].append(trip)

            # Add day items under week
            for date_key in sorted(day_groups.keys()):
                day_trips = day_groups[date_key]
                day_miles = sum(t.get('distance', 0) for t in day_trips)
                day_biz = sum(t.get('distance', 0) for t in day_trips if t.get('computed_category') == 'BUSINESS')
                day_personal = sum(t.get('distance', 0) for t in day_trips if t.get('computed_category') == 'PERSONAL')

                date_str = date_key.strftime('%Y-%m-%d (%a)') if hasattr(date_key, 'strftime') else str(date_key)
                day_item = QTreeWidgetItem([
                    f"  {date_str}",
                    f"{len(day_trips)} trips",
                    f"{day_biz:.1f} mi",
                    f"{day_personal:.1f} mi",
                    f"{day_miles:.1f} mi"
                ])
                day_item.setData(0, Qt.ItemDataRole.UserRole, {'type': 'day', 'trips': day_trips, 'date': date_key})
                if day_biz > 0:
                    day_item.setForeground(2, QColor('#2e7d32'))
                if day_personal > 0:
                    day_item.setForeground(3, QColor('#e65100'))

                # Add individual trips under day
                for trip in sorted(day_trips, key=lambda t: t.get('started')):
                    time_str = trip['started'].strftime('%H:%M') if hasattr(trip['started'], 'strftime') else ''
                    cat = trip.get('computed_category', 'PERSONAL')
                    dist = trip.get('distance', 0)
                    dest = trip.get('end_address', '')[:40]
                    biz_name = trip.get('business_name', '')

                    trip_item = QTreeWidgetItem([
                        f"    {time_str}",
                        cat,
                        f"{dist:.1f} mi",
                        biz_name,
                        dest
                    ])
                    trip_item.setData(0, Qt.ItemDataRole.UserRole, {'type': 'trip', 'trip': trip})

                    # Color by category
                    if cat == 'BUSINESS':
                        trip_item.setForeground(1, QColor('#2e7d32'))
                    elif cat == 'PERSONAL':
                        trip_item.setForeground(1, QColor('#e65100'))
                    elif cat == 'COMMUTE':
                        trip_item.setForeground(1, QColor('#1565c0'))

                    day_item.addChild(trip_item)

                week_item.addChild(day_item)

            self.tree.addTopLevelItem(week_item)

    def _populate_by_day_tree(self):
        """Populate tree widget with expandable days and trips"""
        self.tree.clear()

        for day_data in self.day_grouped_data:
            date = day_data['date']
            trips = day_data['trips']

            # Create day item (parent)
            date_str = date.strftime('%Y-%m-%d (%a)')
            trip_count = len(trips)
            total_miles = day_data['total_miles']
            biz_miles = day_data['business_miles']

            day_item = QTreeWidgetItem([
                f"{date_str} - {trip_count} trips",
                "",
                "",
                f"{total_miles:.1f} mi",
                f"Business: {biz_miles:.1f} mi"
            ])
            day_item.setData(0, Qt.ItemDataRole.UserRole, {'type': 'day', 'data': day_data})

            # Style the day row
            font = QFont()
            font.setBold(True)
            day_item.setFont(0, font)
            if date.weekday() >= 5:  # Weekend
                day_item.setForeground(0, QColor('#9c27b0'))

            # Add trip items as children
            for trip in trips:
                time_str = trip['started'].strftime('%H:%M') if hasattr(trip.get('started'), 'strftime') else ''
                cat = trip.get('computed_category', 'PERSONAL')
                distance = trip.get('distance', 0)
                dest = trip.get('end_address', '')[:50]
                biz_name = trip.get('business_name', '')
                if biz_name:
                    dest = f"{biz_name} ({dest})"

                trip_item = QTreeWidgetItem([
                    f"  → {trip.get('start_address', '')[:40]}...",
                    time_str,
                    cat,
                    f"{distance:.1f} mi",
                    dest
                ])
                trip_item.setData(0, Qt.ItemDataRole.UserRole, {'type': 'trip', 'data': trip})

                # Color by category
                if cat == 'BUSINESS':
                    trip_item.setForeground(2, QColor('#2e7d32'))
                    trip_item.setBackground(2, QColor('#e8f5e9'))
                elif cat == 'PERSONAL':
                    trip_item.setForeground(2, QColor('#e65100'))
                    trip_item.setBackground(2, QColor('#fff3e0'))
                elif cat == 'COMMUTE':
                    trip_item.setForeground(2, QColor('#1565c0'))
                    trip_item.setBackground(2, QColor('#e3f2fd'))

                day_item.addChild(trip_item)

            self.tree.addTopLevelItem(day_item)

    def _on_tree_item_clicked(self, item: QTreeWidgetItem, column: int):
        """Handle single click on tree item"""
        user_data = item.data(0, Qt.ItemDataRole.UserRole)
        if not user_data:
            return

        if user_data['type'] == 'trip':
            # Single trip selected - show it on map (handle both 'trip' and 'data' keys)
            trip = user_data.get('trip') or user_data.get('data')
            if trip:
                self.trip_selected.emit(trip)
        elif user_data['type'] == 'day':
            # Day selected - show all trips for that day (handle both formats)
            trips = user_data.get('trips', [])
            if not trips:
                day_data = user_data.get('data', {})
                trips = day_data.get('trips', [])
            if trips:
                self.show_daily_journey.emit(trips)
        elif user_data['type'] == 'week':
            # Week selected - show all trips for that week
            week_data = user_data.get('data', {})
            trips = week_data.get('trips', [])
            if trips:
                self.show_daily_journey.emit(trips)

    def _on_tree_item_double_clicked(self, item: QTreeWidgetItem, column: int):
        """Handle double click on tree item - expand/collapse or show route"""
        user_data = item.data(0, Qt.ItemDataRole.UserRole)
        if not user_data:
            return

        if user_data['type'] == 'day':
            # Toggle expansion
            item.setExpanded(not item.isExpanded())
        elif user_data['type'] == 'trip':
            # Show trip on map (handle both 'trip' and 'data' keys)
            trip = user_data.get('trip') or user_data.get('data')
            if trip:
                self.trip_selected.emit(trip)

    def _on_tree_item_expanded(self, item: QTreeWidgetItem):
        """Handle tree item expansion"""
        pass  # Could load additional data here if needed

    def _on_tree_selection_changed(self):
        """Handle tree selection change (for arrow key navigation)"""
        selected = self.tree.selectedItems()
        if not selected:
            return

        item = selected[0]
        user_data = item.data(0, Qt.ItemDataRole.UserRole)
        if not user_data:
            return

        if user_data['type'] == 'trip':
            # Show single trip on map - handle both 'trip' and 'data' keys for compatibility
            trip = user_data.get('trip') or user_data.get('data')
            if trip:
                self.trip_selected.emit(trip)
        elif user_data['type'] == 'day':
            # Show all trips for the day - handle both formats
            # Weekly view: {'type': 'day', 'trips': [...]}
            # By-day view: {'type': 'day', 'data': {'trips': [...]}}
            trips = user_data.get('trips')
            if not trips:
                day_data = user_data.get('data', {})
                trips = day_data.get('trips', [])
            if trips:
                self.show_daily_journey.emit(trips)

    def _show_tree_context_menu(self, pos):
        """Show context menu for tree widget items"""
        item = self.tree.itemAt(pos)
        if not item:
            return

        user_data = item.data(0, Qt.ItemDataRole.UserRole)
        if not user_data:
            return

        menu = QMenu(self)
        item_type = user_data.get('type', '')

        # Get the trip(s) for this item
        trips = []
        if item_type == 'trip':
            trip = user_data.get('trip') or user_data.get('data')
            trips = [trip] if trip else []
        elif item_type == 'day':
            # Handle both formats: {'trips': [...]} or {'data': {'trips': [...]}}
            trips = user_data.get('trips', [])
            if not trips:
                day_data = user_data.get('data', {})
                trips = day_data.get('trips', [])
        elif item_type == 'week':
            week_data = user_data.get('data', {})
            trips = week_data.get('trips', [])

        if not trips:
            return

        # Assign name submenu
        assign_menu = menu.addMenu("Assign Name")
        home_action = assign_menu.addAction("Home")
        office_action = assign_menu.addAction("Office")
        personal_action = assign_menu.addAction("[PERSONAL]")
        assign_menu.addSeparator()

        # Existing names - alphabetical groups
        existing_names = self._get_existing_business_names()
        name_actions = {}
        if existing_names:
            existing_menu = assign_menu.addMenu("Existing Names")
            sorted_names = sorted(existing_names, key=str.lower)
            alpha_groups = [
                ('A-C', 'ABC'), ('D-F', 'DEF'), ('G-I', 'GHI'), ('J-L', 'JKL'),
                ('M-O', 'MNO'), ('P-R', 'PQR'), ('S-U', 'STU'), ('V-Z', 'VWXYZ'),
                ('#', '0123456789')
            ]
            for group_label, letters in alpha_groups:
                group_names = [n for n in sorted_names if n and n[0].upper() in letters]
                if group_names:
                    group_menu = existing_menu.addMenu(f"{group_label} ({len(group_names)})")
                    for name in group_names:
                        action = group_menu.addAction(name)
                        name_actions[action] = name
            assign_menu.addSeparator()

        custom_action = assign_menu.addAction("Custom Name...")
        menu.addSeparator()

        # Category submenu
        cat_menu = menu.addMenu("Set Category")
        business_action = cat_menu.addAction("Business")
        personal_cat_action = cat_menu.addAction("Personal")
        commute_action = cat_menu.addAction("Commute")
        menu.addSeparator()

        # Add/Edit Note (only for single trip)
        edit_note_action = None
        if item_type == 'trip' and len(trips) == 1:
            edit_note_action = menu.addAction("Add/Edit Note...")
            menu.addSeparator()

        # Map actions
        view_map_action = menu.addAction("View on Map")
        show_day_action = menu.addAction("Show Day's Journey")

        action = menu.exec(self.tree.viewport().mapToGlobal(pos))

        if action == home_action:
            self._apply_name_to_trips(trips, "Home")
        elif action == office_action:
            self._apply_name_to_trips(trips, "Office")
        elif action == personal_action:
            self._apply_name_to_trips(trips, "[PERSONAL]")
        elif action == custom_action:
            self._prompt_custom_name_for_trips(trips)
        elif action in name_actions:
            self._apply_name_to_trips(trips, name_actions[action])
        elif action == business_action:
            self._apply_category_to_trips(trips, "BUSINESS")
        elif action == personal_cat_action:
            self._apply_category_to_trips(trips, "PERSONAL")
        elif action == commute_action:
            self._apply_category_to_trips(trips, "COMMUTE")
        elif action == edit_note_action and trips:
            self._edit_note_for_trip(trips[0])
        elif action == view_map_action and trips:
            self.trip_selected.emit(trips[0])
        elif action == show_day_action and trips:
            self.show_daily_journey.emit(trips)

    def _apply_name_to_trips(self, trips: list, name: str):
        """Apply business name directly to a list of trips"""
        for trip in trips:
            trip['business_name'] = name
            addr = trip.get('end_address', '')
            if addr:
                cat = trip.get('computed_category')
                self._save_business_mapping(addr, name, cat)
        self.mapping_changed.emit()
        self._refresh_table()

    def _prompt_custom_name_for_trips(self, trips: list):
        """Prompt for custom name for trips"""
        from PyQt6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(
            self, "Custom Business Name",
            f"Enter name for {len(trips)} trip(s):"
        )
        if ok and name.strip():
            self._apply_name_to_trips(trips, name.strip())

    def _apply_category_to_trips(self, trips: list, category: str):
        """Apply category directly to a list of trips"""
        for trip in trips:
            trip['computed_category'] = category
            trip['auto_category'] = category.lower()
        self.trip_updated.emit({}, 'category', category)
        self._refresh_table()

    def _edit_note_for_trip(self, trip: dict):
        """Edit note for a single trip"""
        trip_key = get_trip_key(trip)
        notes = load_trip_notes()
        current_note = notes.get(trip_key, '')

        trip_info = f"{trip['started'].strftime('%Y-%m-%d %H:%M')} - {trip.get('end_address', '')[:60]}"
        dialog = NoteDialog(trip_info, current_note, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            success, msg = save_trip_note(trip, dialog.get_note())
            if success:
                self._refresh_table()
            else:
                QMessageBox.warning(self, "Save Error", msg)

    def _apply_filters(self):
        """Apply all filters to the table or tree"""
        category = self.category_filter.currentText()
        status = self.status_filter.currentText()
        business = self.business_filter.currentText()
        search = self.search_box.text().lower()
        hide_micro = self.hide_micro_trips.isChecked()

        # Auto-switch to All Trips view when filtering by Unresolved or Duplicates
        # because grouped view hides individual unresolved trips
        if status in ["Unresolved", "Duplicates"] and self.view_mode == "grouped":
            self.view_mode_combo.setCurrentText("All Trips")
            return  # setCurrentText triggers _on_view_mode_changed which calls _apply_filters

        visible_count = 0

        # Handle tree filtering for by_day mode
        if self.view_mode == "by_day":
            for i in range(self.tree.topLevelItemCount()):
                day_item = self.tree.topLevelItem(i)
                user_data = day_item.data(0, Qt.ItemDataRole.UserRole)
                if not user_data:
                    continue
                day_data = user_data.get('data', {})
                trips = day_data.get('trips', [])

                show_day = True
                visible_trips = 0

                # Check each trip in this day
                for j in range(day_item.childCount()):
                    trip_item = day_item.child(j)
                    trip_data = trip_item.data(0, Qt.ItemDataRole.UserRole)
                    if not trip_data:
                        continue
                    trip = trip_data.get('data', {})

                    show_trip = True

                    # Micro-trip filter
                    if hide_micro and trip.get('is_micro_trip'):
                        show_trip = False

                    # Category filter
                    if category != "All" and trip.get('computed_category', '') != category.upper():
                        show_trip = False

                    # Search filter
                    if search and show_trip:
                        from_addr = trip.get('start_address', '').lower()
                        to_addr = trip.get('end_address', '').lower()
                        name = trip.get('business_name', '').lower()
                        if search not in from_addr and search not in to_addr and search not in name:
                            show_trip = False

                    trip_item.setHidden(not show_trip)
                    if show_trip:
                        visible_trips += 1

                # Hide day if no visible trips
                if visible_trips == 0:
                    show_day = False

                day_item.setHidden(not show_day)
                if show_day:
                    visible_count += 1

            total = len(self.day_grouped_data)
            self.stats_label.setText(f"{visible_count} of {total} days")
            return

        for row in range(self.table.rowCount()):
            show = True
            data_index = self._get_data_index(row)

            if self.view_mode == "grouped":
                data = self.grouped_data[data_index] if data_index < len(self.grouped_data) else None
                if data:
                    # Category filter
                    if category != "All" and data['primary_category'] != category.upper():
                        show = False
                    # Status filter
                    if status != "All":
                        if status == "Unresolved" and data['status'] not in ['Needs Name', 'Unconfirmed Business']:
                            show = False
                        elif status == "Resolved" and data['status'] != 'Has Name':
                            show = False
                        elif status == "Unconfirmed" and data['status'] != 'Unconfirmed Business':
                            show = False
                    # Business name filter
                    if business != "All":
                        if business == "(No Name)" and data.get('business_name', ''):
                            show = False
                        elif business != "(No Name)" and data.get('business_name', '') != business:
                            show = False
                    # Search
                    if search:
                        addr = data['address'].lower()
                        name = data.get('business_name', '').lower()
                        if search not in addr and search not in name:
                            show = False
            else:
                trip = self.trips_data[data_index] if data_index < len(self.trips_data) else None
                if trip:
                    # Micro-trip filter
                    if hide_micro and trip.get('is_micro_trip'):
                        show = False
                    # Category filter
                    if category != "All" and trip.get('computed_category', '') != category.upper():
                        show = False
                    # Status filter
                    if status != "All":
                        has_name = bool(trip.get('business_name', ''))
                        is_business = trip.get('computed_category') == 'BUSINESS'
                        is_duplicate = trip.get('is_duplicate', False)
                        if status == "Unresolved" and has_name:
                            show = False
                        elif status == "Resolved" and not has_name:
                            show = False
                        elif status == "Unconfirmed" and not (is_business and not has_name):
                            show = False
                        elif status == "Duplicates" and not is_duplicate:
                            show = False
                    # Business name filter
                    if business != "All":
                        trip_name = trip.get('business_name', '')
                        if business == "(No Name)" and trip_name:
                            show = False
                        elif business != "(No Name)" and trip_name != business:
                            show = False
                    # Search
                    if search:
                        from_addr = trip.get('start_address', '').lower()
                        to_addr = trip.get('end_address', '').lower()
                        name = trip.get('business_name', '').lower()
                        if search not in from_addr and search not in to_addr and search not in name:
                            show = False

            self.table.setRowHidden(row, not show)
            if show:
                visible_count += 1

        if self.view_mode == "grouped":
            total = len(self.grouped_data)
            self.stats_label.setText(f"{visible_count} of {total} destinations")
        elif self.view_mode == "by_day":
            total = len(self.day_grouped_data)
            self.stats_label.setText(f"{visible_count} of {total} days")
        else:
            total = len(self.trips_data)
            self.stats_label.setText(f"{visible_count} of {total} trips")

    def _on_selection_changed(self):
        """Handle selection change"""
        rows = self.table.selectionModel().selectedRows()
        if rows:
            visual_row = rows[0].row()
            data_index = self._get_data_index(visual_row)
            
            if self.view_mode == "grouped" and data_index < len(self.grouped_data):
                data = self.grouped_data[data_index]
                # Emit first trip for map display
                if data['trips']:
                    self.trip_selected.emit(data['trips'][0])
            elif self.view_mode == "by_day" and data_index < len(self.day_grouped_data):
                # For day view, emit the daily journey
                data = self.day_grouped_data[data_index]
                if data['trips']:
                    self.show_daily_journey.emit(data['trips'])
            elif self.view_mode == "individual" and data_index < len(self.trips_data):
                self.trip_selected.emit(self.trips_data[data_index])

    def _get_data_index(self, visual_row: int) -> int:
        """Get the original data index for a visual row (handles sorting)"""
        first_item = self.table.item(visual_row, 0)
        if first_item is None:
            return visual_row
        data_index = first_item.data(Qt.ItemDataRole.UserRole)
        return data_index if data_index is not None else visual_row

    def _on_cell_double_clicked(self, visual_row: int, col: int):
        """Handle double-click to edit"""
        # Get the original data index from UserRole (stored in first column)
        data_index = self._get_data_index(visual_row)
        
        if self.view_mode == "grouped":
            if data_index < len(self.grouped_data):
                data = self.grouped_data[data_index]
                if col == 1:  # Business name column
                    self._edit_business_name_grouped(data_index, data)
                elif col == 2:  # Category column
                    self._edit_category_grouped(data_index, data)
        elif self.view_mode == "by_day":
            if data_index < len(self.day_grouped_data):
                data = self.day_grouped_data[data_index]
                # Double-click shows the day's journey on map
                if data['trips']:
                    self.show_daily_journey.emit(data['trips'])
        else:
            if data_index < len(self.trips_data):
                trip = self.trips_data[data_index]
                if col == 6:  # Business name column
                    self._edit_business_name_individual(data_index, trip)
                elif col == 2:  # Category column
                    self._edit_category_individual(data_index, trip)

    def _edit_business_name_grouped(self, row: int, data: dict):
        """Edit business name for a grouped destination"""
        from PyQt6.QtWidgets import QInputDialog
        current_name = data.get('business_name', '')
        if current_name in ['[Unconfirmed]', 'NO_BUSINESS_FOUND']:
            current_name = ''

        name, ok = QInputDialog.getText(
            self, "Edit Business Name",
            f"Business name for:\n{data['address'][:60]}",
            text=current_name
        )

        if ok:
            # Update all trips to this destination
            for trip in data['trips']:
                trip['business_name'] = name

            data['business_name'] = name
            data['status'] = 'Has Name' if name else ('Unconfirmed Business' if data['primary_category'] == 'BUSINESS' else 'Needs Name')

            # Update table cell (column 0 = Business Name in grouped view)
            name_item = self.table.item(row, 0)
            if name_item:
                if name:
                    name_item.setText(name)
                    name_item.setForeground(QColor('#000000'))
                    name_item.setFont(QFont())
                elif data['status'] == 'Unconfirmed Business':
                    name_item.setText('[Needs Confirmation]')
                    name_item.setForeground(QColor('#ff8f00'))
                    name_item.setFont(QFont('', -1, -1, True))
                else:
                    name_item.setText('[Unknown]')
                    name_item.setForeground(QColor('#999999'))
                    name_item.setFont(QFont('', -1, -1, True))

            # Update status cell (column 4 = Status in grouped view)
            status_item = self.table.item(row, 4)
            if status_item:
                if data['status'] == 'Needs Name':
                    status_item.setText('Needs Name')
                    status_item.setForeground(QColor('#d32f2f'))
                elif data['status'] == 'Unconfirmed Business':
                    status_item.setText('Unconfirmed')
                    status_item.setForeground(QColor('#ff8f00'))
                else:
                    status_item.setText('Confirmed')
                    status_item.setForeground(QColor('#388e3c'))

            # Save to mapping file
            if name:
                self._save_business_mapping(data['address'], name)

    def _edit_category_grouped(self, row: int, data: dict):
        """Edit category for all trips to a destination"""
        from PyQt6.QtWidgets import QInputDialog
        categories = ["BUSINESS", "PERSONAL", "COMMUTE"]
        current = data['primary_category']
        current_idx = categories.index(current) if current in categories else 1

        category, ok = QInputDialog.getItem(
            self, "Change Category",
            f"Category for all trips to:\n{data['address'][:60]}",
            categories, current_idx, False
        )

        if ok and category != current:
            # Update all trips
            for trip in data['trips']:
                trip['computed_category'] = category
                trip['auto_category'] = category.lower()

            data['primary_category'] = category
            data['status'] = 'Has Name' if data.get('business_name') else ('Unconfirmed Business' if category == 'BUSINESS' else 'Needs Name')

            # Update table (column 1 = Category in grouped view)
            cat_item = self.table.item(row, 1)
            if cat_item:
                cat_item.setText(category)
                if category == 'BUSINESS':
                    cat_item.setBackground(QColor('#e8f5e9'))
                    cat_item.setForeground(QColor('#2e7d32'))
                elif category == 'PERSONAL':
                    cat_item.setBackground(QColor('#fff3e0'))
                    cat_item.setForeground(QColor('#e65100'))
                elif category == 'COMMUTE':
                    cat_item.setBackground(QColor('#e3f2fd'))
                    cat_item.setForeground(QColor('#1565c0'))

            self.trip_updated.emit(data['trips'][0] if data['trips'] else {}, 'category', category)

    def _edit_business_name_individual(self, row: int, trip: dict):
        """Edit business name for a single trip"""
        from PyQt6.QtWidgets import QInputDialog
        current_name = trip.get('business_name', '')

        name, ok = QInputDialog.getText(
            self, "Edit Business Name",
            f"Business name for:\n{trip.get('end_address', '')[:60]}",
            text=current_name
        )

        if ok:
            trip['business_name'] = name
            # Column 8 = Business Name in individual view
            name_item = self.table.item(row, 8)
            if name_item:
                name_item.setText(name if name else '[Unconfirmed]' if trip.get('computed_category') == 'BUSINESS' else '')

            if name:
                self._save_business_mapping(trip.get('end_address', ''), name)

    def _edit_category_individual(self, row: int, trip: dict):
        """Edit category for a single trip"""
        from PyQt6.QtWidgets import QInputDialog
        categories = ["BUSINESS", "PERSONAL", "COMMUTE"]
        current = trip.get('computed_category', 'PERSONAL')
        current_idx = categories.index(current) if current in categories else 1

        category, ok = QInputDialog.getItem(
            self, "Change Category",
            f"Category for trip to:\n{trip.get('end_address', '')[:60]}",
            categories, current_idx, False
        )

        if ok and category != current:
            trip['computed_category'] = category
            # Column 4 = Category in individual view
            cat_item = self.table.item(row, 4)
            if cat_item:
                cat_item.setText(category)
                if category == 'BUSINESS':
                    cat_item.setBackground(QColor('#e8f5e9'))
                    cat_item.setForeground(QColor('#2e7d32'))
                elif category == 'PERSONAL':
                    cat_item.setBackground(QColor('#fff3e0'))
                    cat_item.setForeground(QColor('#e65100'))
                elif category == 'COMMUTE':
                    cat_item.setBackground(QColor('#e3f2fd'))
                    cat_item.setForeground(QColor('#1565c0'))

            self.trip_updated.emit(trip, 'category', category)

    def _show_context_menu(self, pos):
        """Show context menu on right-click"""
        visual_row = self.table.rowAt(pos.y())
        if visual_row < 0:
            return

        # Select row if not already selected
        if not self.table.item(visual_row, 0).isSelected():
            self.table.selectRow(visual_row)

        # Convert clicked row to data index
        row = self._get_data_index(visual_row)

        # Get selected rows and convert to data indices
        visual_rows = list(set(idx.row() for idx in self.table.selectedIndexes()))
        selected_rows = [self._get_data_index(vr) for vr in visual_rows]

        menu = QMenu(self)

        # Assign name submenu
        assign_menu = menu.addMenu("Assign Name")
        home_action = assign_menu.addAction("Home")
        office_action = assign_menu.addAction("Office")
        personal_action = assign_menu.addAction("[PERSONAL]")
        assign_menu.addSeparator()

        # Existing names - organized alphabetically
        existing_names = self._get_existing_business_names()
        name_actions = {}
        if existing_names:
            existing_menu = assign_menu.addMenu("Existing Names")
            # Group names alphabetically (A-C, D-F, etc.)
            sorted_names = sorted(existing_names, key=str.lower)
            alpha_groups = [
                ('A-C', 'ABC'),
                ('D-F', 'DEF'),
                ('G-I', 'GHI'),
                ('J-L', 'JKL'),
                ('M-O', 'MNO'),
                ('P-R', 'PQR'),
                ('S-U', 'STU'),
                ('V-Z', 'VWXYZ'),
                ('#', '0123456789')
            ]
            for group_label, letters in alpha_groups:
                group_names = [n for n in sorted_names if n and n[0].upper() in letters]
                if group_names:
                    group_menu = existing_menu.addMenu(f"{group_label} ({len(group_names)})")
                    for name in group_names:
                        action = group_menu.addAction(name)
                        name_actions[action] = name
            assign_menu.addSeparator()

        custom_action = assign_menu.addAction("Custom Name...")

        menu.addSeparator()

        # Category submenu
        cat_menu = menu.addMenu("Set Category")
        business_action = cat_menu.addAction("Business")
        personal_cat_action = cat_menu.addAction("Personal")
        commute_action = cat_menu.addAction("Commute")

        menu.addSeparator()

        # Select nearby (only in grouped view)
        select_nearby_action = None
        if self.view_mode == "grouped":
            select_nearby_action = menu.addAction("Select Nearby...")
            menu.addSeparator()

        # Add/Edit Note (only in individual view for single selection)
        edit_note_action = None
        view_merged_action = None
        micro_merge_prev_action = None
        micro_merge_next_action = None
        micro_mark_valid_action = None
        micro_discard_action = None
        if self.view_mode == "individual" and len(selected_rows) == 1:
            edit_note_action = menu.addAction("Add/Edit Note...")
            # Check if this is a merged trip
            trip = self.trips_data[row] if row < len(self.trips_data) else None
            if trip and trip.get('is_merged'):
                view_merged_action = menu.addAction("View Merged Details...")
            # Check if this is a micro-trip
            if trip and trip.get('is_micro_trip'):
                menu.addSeparator()
                micro_menu = menu.addMenu("Micro-Trip Options")
                micro_merge_prev_action = micro_menu.addAction("Merge with Previous Trip")
                micro_merge_next_action = micro_menu.addAction("Merge with Next Trip")
                micro_menu.addSeparator()
                micro_mark_valid_action = micro_menu.addAction("Mark as Valid Trip")
                micro_discard_action = micro_menu.addAction("Discard Trip")
            menu.addSeparator()

        # Map actions
        view_map_action = menu.addAction("View on Map")
        show_day_action = menu.addAction("Show Day's Journey")
        open_gmaps_action = menu.addAction("Open in Google Maps")

        action = menu.exec(self.table.viewport().mapToGlobal(pos))

        if action == home_action:
            self._apply_name_to_selected(selected_rows, "Home")
        elif action == office_action:
            self._apply_name_to_selected(selected_rows, "Office")
        elif action == personal_action:
            self._apply_name_to_selected(selected_rows, "[PERSONAL]")
        elif action == custom_action:
            self._prompt_custom_name(selected_rows)
        elif action in name_actions:
            self._apply_name_to_selected(selected_rows, name_actions[action])
        elif action == business_action:
            self._apply_category_to_selected(selected_rows, "BUSINESS")
        elif action == personal_cat_action:
            self._apply_category_to_selected(selected_rows, "PERSONAL")
        elif action == commute_action:
            self._apply_category_to_selected(selected_rows, "COMMUTE")
        elif action == select_nearby_action:
            self._select_nearby(row)
        elif action == edit_note_action:
            self._edit_trip_note(row)
        elif action == view_merged_action:
            self._view_merged_details(row)
        elif action == micro_merge_prev_action:
            self._merge_micro_trip(row, merge_with='previous')
        elif action == micro_merge_next_action:
            self._merge_micro_trip(row, merge_with='next')
        elif action == micro_mark_valid_action:
            self._mark_micro_trip_valid(row)
        elif action == micro_discard_action:
            self._discard_micro_trip(row)
        elif action == view_map_action:
            self._view_on_map(row)
        elif action == show_day_action:
            self._show_day_journey(row)
        elif action == open_gmaps_action:
            self._open_in_google_maps(row)

    def _apply_name_to_selected(self, data_indices: List[int], name: str, category: str = None):
        """Apply business name to selected rows (data_indices are original data indices)

        If category is provided, it will be saved with the mapping for future use.
        """
        for data_index in data_indices:
            if self.view_mode == "grouped" and data_index < len(self.grouped_data):
                data = self.grouped_data[data_index]
                trip_category = category or data.get('primary_category')
                for trip in data['trips']:
                    trip['business_name'] = name
                data['business_name'] = name
                data['status'] = 'Has Name'
                self._save_business_mapping(data['address'], name, trip_category)

            elif self.view_mode == "individual" and data_index < len(self.trips_data):
                trip = self.trips_data[data_index]
                trip_category = category or trip.get('computed_category')
                trip['business_name'] = name
                self._save_business_mapping(trip.get('end_address', ''), name, trip_category)

            elif self.view_mode == "by_day" and data_index < len(self.day_grouped_data):
                # For by_day view, apply to all trips on that day
                day_data = self.day_grouped_data[data_index]
                for trip in day_data.get('trips', []):
                    trip_category = category or trip.get('computed_category')
                    trip['business_name'] = name
                    self._save_business_mapping(trip.get('end_address', ''), name, trip_category)

        self.mapping_changed.emit()
        self._update_business_filter()
        self._refresh_table()

    def _prompt_custom_name(self, rows: List[int]):
        """Prompt for custom name"""
        from PyQt6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(
            self, "Custom Business Name",
            f"Enter name for {len(rows)} selected item(s):"
        )
        if ok and name.strip():
            self._apply_name_to_selected(rows, name.strip())

    def _edit_trip_note(self, data_index: int):
        """Edit note for a trip (individual view only)"""
        if self.view_mode != "individual" or data_index >= len(self.trips_data):
            return

        trip = self.trips_data[data_index]
        trip_key = get_trip_key(trip)
        notes = load_trip_notes()
        current_note = notes.get(trip_key, '')

        trip_info = f"{trip['started'].strftime('%Y-%m-%d %H:%M')} - {trip.get('end_address', '')[:60]}"
        dialog = NoteDialog(trip_info, current_note, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            success, msg = save_trip_note(trip, dialog.get_note())
            if success:
                self._refresh_table()
            else:
                QMessageBox.warning(self, "Save Error", msg)

    def _view_merged_details(self, data_index: int):
        """Show details of a merged trip"""
        if self.view_mode != "individual" or data_index >= len(self.trips_data):
            return

        trip = self.trips_data[data_index]
        if not trip.get('is_merged') or 'merged_from' not in trip:
            return

        merged_from = trip['merged_from']

        # Build details text
        details = "This trip was automatically merged from the following short segments:\n"
        details += "(Likely caused by stopping at red lights or in traffic)\n\n"

        total_distance = 0
        for i, segment in enumerate(merged_from, 1):
            details += f"Segment {i}:\n"
            details += f"  Started: {segment.get('started', 'N/A')}\n"
            details += f"  Stopped: {segment.get('stopped', 'N/A')}\n"
            details += f"  Distance: {segment.get('distance', 0):.2f} mi\n"
            details += f"  From: {segment.get('start_address', 'N/A')}\n"
            details += f"  To: {segment.get('end_address', 'N/A')}\n\n"
            total_distance += segment.get('distance', 0)

        details += f"Combined Distance: {total_distance:.2f} mi"

        msg = QMessageBox(self)
        msg.setWindowTitle("Merged Trip Details")
        msg.setText(f"Trip merged from {len(merged_from)} segments")
        msg.setDetailedText(details)
        msg.setIcon(QMessageBox.Icon.Information)
        msg.exec()

    def _merge_micro_trip(self, data_index: int, merge_with: str):
        """Merge a micro-trip with the previous or next trip"""
        if self.view_mode != "individual" or data_index >= len(self.trips_data):
            return

        trip = self.trips_data[data_index]

        # Find the trip to merge with
        if merge_with == 'previous' and data_index > 0:
            target_index = data_index - 1
            target_trip = self.trips_data[target_index]
            # Add micro-trip's distance to target
            target_trip['distance'] = target_trip.get('distance', 0) + trip.get('distance', 0)
            # Update end address/time to micro-trip's end
            target_trip['end_address'] = trip.get('end_address', target_trip.get('end_address'))
            target_trip['stopped'] = trip.get('stopped', target_trip.get('stopped'))
            target_trip['end_odometer'] = trip.get('end_odometer', target_trip.get('end_odometer'))
            # Mark as merged
            target_trip['is_merged'] = True
            target_trip['merge_count'] = target_trip.get('merge_count', 1) + 1
            # Remove the micro-trip
            self.trips_data.pop(data_index)
            self.trip_updated.emit({}, 'merge', 'previous')

        elif merge_with == 'next' and data_index < len(self.trips_data) - 1:
            target_index = data_index + 1
            target_trip = self.trips_data[target_index]
            # Add micro-trip's distance to target
            target_trip['distance'] = target_trip.get('distance', 0) + trip.get('distance', 0)
            # Update start address/time to micro-trip's start
            target_trip['start_address'] = trip.get('start_address', target_trip.get('start_address'))
            target_trip['started'] = trip.get('started', target_trip.get('started'))
            target_trip['start_odometer'] = trip.get('start_odometer', target_trip.get('start_odometer'))
            # Mark as merged
            target_trip['is_merged'] = True
            target_trip['merge_count'] = target_trip.get('merge_count', 1) + 1
            # Remove the micro-trip
            self.trips_data.pop(data_index)
            self.trip_updated.emit({}, 'merge', 'next')

        else:
            QMessageBox.warning(self, "Cannot Merge",
                f"No {'previous' if merge_with == 'previous' else 'next'} trip to merge with.")
            return

        self._refresh_table()

    def _mark_micro_trip_valid(self, data_index: int):
        """Mark a micro-trip as a valid trip (remove the flag)"""
        if self.view_mode != "individual" or data_index >= len(self.trips_data):
            return

        trip = self.trips_data[data_index]
        trip['is_micro_trip'] = False
        trip.pop('micro_reason', None)
        self._refresh_table()

    def _discard_micro_trip(self, data_index: int):
        """Discard (remove) a micro-trip from the analysis"""
        if self.view_mode != "individual" or data_index >= len(self.trips_data):
            return

        trip = self.trips_data[data_index]
        reply = QMessageBox.question(
            self, "Discard Trip",
            f"Are you sure you want to discard this trip?\n\n"
            f"Date: {trip['started'].strftime('%Y-%m-%d %H:%M')}\n"
            f"Distance: {trip.get('distance', 0):.2f} mi\n"
            f"From: {trip.get('start_address', 'N/A')}\n"
            f"To: {trip.get('end_address', 'N/A')}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.trips_data.pop(data_index)
            self.trip_updated.emit({}, 'discard', None)
            self._refresh_table()

    def _apply_category_to_selected(self, data_indices: List[int], category: str):
        """Apply category to selected rows (data_indices are original data indices)"""
        for data_index in data_indices:
            if self.view_mode == "grouped" and data_index < len(self.grouped_data):
                data = self.grouped_data[data_index]
                for trip in data['trips']:
                    trip['computed_category'] = category
                    trip['auto_category'] = category.lower()
                data['primary_category'] = category

            elif self.view_mode == "individual" and data_index < len(self.trips_data):
                trip = self.trips_data[data_index]
                trip['computed_category'] = category
                trip['auto_category'] = category.lower()

            elif self.view_mode == "by_day" and data_index < len(self.day_grouped_data):
                # For by_day view, apply to all trips on that day
                day_data = self.day_grouped_data[data_index]
                for trip in day_data.get('trips', []):
                    trip['computed_category'] = category
                    trip['auto_category'] = category.lower()

        self.trip_updated.emit({}, 'category', category)
        self._refresh_table()

    def _select_nearby(self, row: int):
        """Select nearby destinations"""
        if row >= len(self.grouped_data):
            return

        current = self.grouped_data[row]
        current_street = current.get('street', '')
        current_lat = current.get('lat')
        current_lng = current.get('lng')

        # Load business mapping for suggestions
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        business_mapping = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    business_mapping = json.load(f)
            except:
                pass

        # Find nearby
        nearby_rows = []
        nearby_resolved = []

        for i, data in enumerate(self.grouped_data):
            if i == row:
                continue

            match_reason = self._check_nearby_match(
                current_street, current_lat, current_lng,
                data.get('street', ''), data.get('lat'), data.get('lng')
            )

            if match_reason:
                nearby_rows.append((i, data, match_reason))

        # Check business mapping for resolved nearby
        for addr, name in business_mapping.items():
            if addr == current['address']:
                continue
            mapped_street = self._extract_street(addr)
            match_reason = self._check_nearby_match(
                current_street, current_lat, current_lng,
                mapped_street, None, None
            )
            if match_reason:
                nearby_resolved.append({'address': addr, 'name': name, 'reason': match_reason})

        # Find suggested name
        suggested_name = None
        name_counts = {}
        for n in nearby_resolved:
            name = n.get('name', '')
            if name and name not in ['[PERSONAL]', 'Unknown', 'Home', 'Office']:
                name_counts[name] = name_counts.get(name, 0) + 1
        if name_counts:
            suggested_name = max(name_counts, key=name_counts.get)

        if not nearby_rows and not nearby_resolved:
            QMessageBox.information(self, "No Nearby", f"No nearby destinations found for:\n{current['address']}")
            return

        # Build message
        msg = f"Selected: {current['address'][:55]}...\n\n"

        if suggested_name:
            msg += f"SUGGESTED NAME: {suggested_name}\n"
            msg += f"  ({name_counts[suggested_name]} nearby already use this name)\n\n"

        if nearby_resolved:
            msg += f"Already resolved nearby ({len(nearby_resolved)}):\n"
            for n in nearby_resolved[:5]:
                msg += f"  ✓ {n['name']}: {n['address'][:35]}...\n"
            if len(nearby_resolved) > 5:
                msg += f"    ... and {len(nearby_resolved) - 5} more\n"
            msg += "\n"

        if nearby_rows:
            msg += f"Unresolved nearby ({len(nearby_rows)}):\n"
            for i, data, reason in nearby_rows[:5]:
                msg += f"  • {data['address'][:45]}...\n"
            if len(nearby_rows) > 5:
                msg += f"    ... and {len(nearby_rows) - 5} more\n"
            msg += "\nSelect these destinations?"

            reply = QMessageBox.question(
                self, "Select Nearby",
                msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.table.clearSelection()
                self.table.selectRow(row)
                for i, data, reason in nearby_rows:
                    self.table.selectRow(i)

                # TODO: Could pre-fill suggested name somewhere
        else:
            QMessageBox.information(self, "Nearby Addresses", msg + "All nearby are already resolved!")

    def _check_nearby_match(self, current_street, current_lat, current_lng, other_street, other_lat, other_lng):
        """Check if nearby match"""
        import re

        if other_street == current_street and current_street:
            return f"Same street: {current_street}"

        if current_street and other_street:
            current_num = re.match(r'^(\d+)', current_street)
            other_num = re.match(r'^(\d+)', other_street)
            if current_num and other_num and current_num.group(1) == other_num.group(1):
                return f"Similar street"

        if current_lat and current_lng and other_lat and other_lng:
            distance = self._calculate_distance(current_lat, current_lng, other_lat, other_lng)
            if distance <= 0.25:
                return f"Within {distance:.2f} mi"

        return None

    def _calculate_distance(self, lat1, lng1, lat2, lng2):
        """Calculate distance in miles"""
        import math
        R = 3959
        lat1_rad = math.radians(lat1)
        lat2_rad = math.radians(lat2)
        delta_lat = math.radians(lat2 - lat1)
        delta_lng = math.radians(lng2 - lng1)
        a = math.sin(delta_lat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(delta_lng/2)**2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
        return R * c

    def _view_on_map(self, row: int):
        """Emit signal to view on map"""
        if self.view_mode == "grouped" and row < len(self.grouped_data):
            data = self.grouped_data[row]
            if data['trips']:
                self.trip_selected.emit(data['trips'][0])
        elif self.view_mode == "individual" and row < len(self.trips_data):
            self.trip_selected.emit(self.trips_data[row])
        elif self.view_mode == "by_day" and row < len(self.day_grouped_data):
            day_data = self.day_grouped_data[row]
            if day_data.get('trips'):
                self.trip_selected.emit(day_data['trips'][0])

    def _open_in_google_maps(self, row: int):
        """Open in external Google Maps"""
        import webbrowser
        addr = ""
        if self.view_mode == "grouped" and row < len(self.grouped_data):
            addr = self.grouped_data[row]['address']
        elif self.view_mode == "individual" and row < len(self.trips_data):
            addr = self.trips_data[row].get('end_address', '')
        elif self.view_mode == "by_day" and row < len(self.day_grouped_data):
            day_data = self.day_grouped_data[row]
            if day_data.get('trips'):
                addr = day_data['trips'][0].get('end_address', '')

        if addr:
            url = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(addr)}"
            webbrowser.open(url)

    def _show_day_journey(self, row: int):
        """Show all trips for the same day as the selected trip"""
        trip_date = None

        if self.view_mode == "grouped" and row < len(self.grouped_data):
            data = self.grouped_data[row]
            if data['trips']:
                trip_date = data['trips'][0].get('started')
        elif self.view_mode == "individual" and row < len(self.trips_data):
            trip_date = self.trips_data[row].get('started')
        elif self.view_mode == "by_day" and row < len(self.day_grouped_data):
            day_data = self.day_grouped_data[row]
            trip_date = day_data.get('date')

        if trip_date:
            # Find all trips on the same date
            target_date = trip_date.date() if hasattr(trip_date, 'date') else trip_date
            day_trips = [
                t for t in self.trips_data 
                if hasattr(t.get('started'), 'date') and t['started'].date() == target_date
            ]
            
            if day_trips:
                self.show_daily_journey.emit(day_trips)

    def _save_business_mapping(self, address: str, name: str, category: str = None):
        """Save business mapping to unified file with optional category

        Format: {address: {"name": name, "category": category, "source": "manual"}}
        """
        if not address or not name:
            return

        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')

        # Load existing mapping
        mappings = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except:
                pass

        # Build entry with source=manual
        entry = {"name": name, "source": "manual"}
        if category:
            entry["category"] = category
        else:
            # Check if existing entry has category to preserve it
            existing = mappings.get(address)
            if isinstance(existing, dict) and existing.get('category'):
                entry["category"] = existing['category']

        mappings[address] = entry

        try:
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, indent=2, ensure_ascii=False)
        except:
            pass

    def _get_existing_business_names(self) -> set:
        """Get existing business names from mappings"""
        names = set()
        skip = {'Home', 'Office', '[PERSONAL]', 'Unknown', '', 'NO_BUSINESS_FOUND'}

        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    for value in json.load(f).values():
                        # Handle both old format (string) and new format (dict)
                        if isinstance(value, dict):
                            name = value.get('name', '')
                        else:
                            name = value
                        if name and name not in skip:
                            names.add(name)
            except:
                pass

        return names


class TripTableWidget(QTableWidget):
    """Table widget for displaying trip data - LEGACY, kept for compatibility"""

    trip_selected = pyqtSignal(dict)
    trip_updated = pyqtSignal(dict, str, str)  # trip, field, new_value - emitted when trip is modified
    mapping_changed = pyqtSignal()  # emitted when business mapping should be saved
    show_daily_journey = pyqtSignal(list)  # list of trips for a day

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()
        self.trips_data = []
        self._loading = False  # Flag to prevent signals during load

    def _setup_ui(self):
        self.setColumnCount(7)
        self.setHorizontalHeaderLabels([
            "Date", "Day", "Category", "Distance",
            "From", "To", "Business Name"
        ])

        # Set column widths
        header = self.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Interactive)

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

        # Cell double-click for editing
        self.cellDoubleClicked.connect(self._on_cell_double_clicked)

    def _on_cell_double_clicked(self, row: int, col: int):
        """Handle double-click to edit business name or category"""
        if row < 0 or row >= len(self.trips_data):
            return

        trip = self.trips_data[row]

        if col == 6:  # Business Name column
            self._edit_business_name(row, trip)
        elif col == 2:  # Category column
            self._show_category_picker(row, trip)

    def _edit_business_name(self, row: int, trip: dict):
        """Show dialog to edit business name"""
        current_name = trip.get('business_name', '')
        if current_name == '[Unconfirmed Business]':
            current_name = ''

        from PyQt6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(
            self,
            "Edit Business Name",
            f"Business name for:\n{trip.get('end_address', '')[:60]}",
            text=current_name
        )

        if ok:
            # Update the trip data
            trip['business_name'] = name
            # Update the table cell
            name_item = self.item(row, 6)
            if name_item:
                name_item.setText(name if name else '[Unconfirmed Business]' if trip.get('computed_category') == 'BUSINESS' else '')
                if not name and trip.get('computed_category') == 'BUSINESS':
                    name_item.setForeground(QColor('#ff8f00'))
                    name_item.setFont(QFont('', -1, -1, True))
                else:
                    name_item.setForeground(QColor('#000000'))
                    name_item.setFont(QFont())

            # Save to mapping file if name provided
            if name:
                self._save_business_mapping(trip.get('end_address', ''), name)

    def _show_category_picker(self, row: int, trip: dict):
        """Show dialog to pick category"""
        from PyQt6.QtWidgets import QInputDialog
        categories = ["BUSINESS", "PERSONAL", "COMMUTE"]
        current = trip.get('computed_category', 'PERSONAL')
        current_idx = categories.index(current) if current in categories else 1

        category, ok = QInputDialog.getItem(
            self,
            "Change Category",
            f"Category for trip to:\n{trip.get('end_address', '')[:60]}",
            categories,
            current_idx,
            False
        )

        if ok and category != current:
            self._update_trip_category(row, trip, category)

    def _show_context_menu(self, pos):
        """Show context menu on right-click"""
        row = self.rowAt(pos.y())
        if row < 0 or row >= len(self.trips_data):
            return

        trip = self.trips_data[row]
        menu = QMenu(self)

        # Assign name submenu with quick picks
        assign_menu = menu.addMenu("Assign Name")
        home_action = assign_menu.addAction("Home")
        office_action = assign_menu.addAction("Office")
        assign_menu.addSeparator()

        # Add existing business names from mapping - organized alphabetically
        existing_names = self._get_existing_business_names()
        name_actions = {}
        if existing_names:
            existing_menu = assign_menu.addMenu("Existing Names")
            sorted_names = sorted(existing_names, key=str.lower)
            alpha_groups = [
                ('A-C', 'ABC'), ('D-F', 'DEF'), ('G-I', 'GHI'), ('J-L', 'JKL'),
                ('M-O', 'MNO'), ('P-R', 'PQR'), ('S-U', 'STU'), ('V-Z', 'VWXYZ'),
                ('#', '0123456789')
            ]
            for group_label, letters in alpha_groups:
                group_names = [n for n in sorted_names if n and n[0].upper() in letters]
                if group_names:
                    group_menu = existing_menu.addMenu(f"{group_label} ({len(group_names)})")
                    for name in group_names:
                        action = group_menu.addAction(name)
                        name_actions[action] = name
            assign_menu.addSeparator()

        custom_action = assign_menu.addAction("Custom Name...")

        menu.addSeparator()

        # Category submenu
        change_category = menu.addMenu("Set Category")
        business_action = change_category.addAction("Business")
        personal_action = change_category.addAction("Personal")
        commute_action = change_category.addAction("Commute")

        # Mark current category
        current_cat = trip.get('computed_category', 'PERSONAL')
        if current_cat == 'BUSINESS':
            business_action.setCheckable(True)
            business_action.setChecked(True)
        elif current_cat == 'PERSONAL':
            personal_action.setCheckable(True)
            personal_action.setChecked(True)
        elif current_cat == 'COMMUTE':
            commute_action.setCheckable(True)
            commute_action.setChecked(True)

        menu.addSeparator()

        # Map actions
        view_map_action = menu.addAction("View on Map")
        open_gmaps_action = menu.addAction("Open in Google Maps")

        action = menu.exec(self.viewport().mapToGlobal(pos))

        if action == home_action:
            self._set_business_name(row, trip, "Home")
        elif action == office_action:
            self._set_business_name(row, trip, "Office")
        elif action == custom_action:
            self._edit_business_name(row, trip)
        elif action in name_actions:
            self._set_business_name(row, trip, name_actions[action])
        elif action == business_action:
            self._update_trip_category(row, trip, 'BUSINESS')
        elif action == personal_action:
            self._update_trip_category(row, trip, 'PERSONAL')
        elif action == commute_action:
            self._update_trip_category(row, trip, 'COMMUTE')
        elif action == view_map_action:
            self.trip_selected.emit(trip)
        elif action == open_gmaps_action:
            import webbrowser
            addr = trip.get('end_address', '')
            url = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(addr)}"
            webbrowser.open(url)

    def _set_business_name(self, row: int, trip: dict, name: str):
        """Set business name for a trip"""
        trip['business_name'] = name
        name_item = self.item(row, 6)
        if name_item:
            name_item.setText(name)
            name_item.setForeground(QColor('#000000'))
            name_item.setFont(QFont())

        # Save to mapping file
        self._save_business_mapping(trip.get('end_address', ''), name)

    def _get_existing_business_names(self) -> set:
        """Get existing business names from mapping and cache files"""
        names = set()
        skip_names = {'Home', 'Office', '[PERSONAL]', 'Unknown', '', 'NO_BUSINESS_FOUND'}

        # From business mapping
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
                    for name in mappings.values():
                        if name and name not in skip_names:
                            names.add(name)
            except:
                pass

        # From address cache
        cache_file = os.path.join(get_app_dir(), 'address_cache.json')
        if os.path.exists(cache_file):
            try:
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cache = json.load(f)
                    for entry in cache.values():
                        if isinstance(entry, dict):
                            name = entry.get('name', '')
                        else:
                            name = str(entry)
                        if name and name not in skip_names:
                            names.add(name)
            except:
                pass

        return names

    def _update_trip_category(self, row: int, trip: dict, new_category: str):
        """Update a trip's category"""
        old_category = trip.get('computed_category', 'PERSONAL')
        if new_category == old_category:
            return

        # Update trip data
        trip['computed_category'] = new_category
        trip['auto_category'] = new_category.lower()

        # Update table cell
        cat_item = self.item(row, 2)
        if cat_item:
            cat_item.setText(new_category)
            if new_category == 'BUSINESS':
                cat_item.setBackground(QColor('#e8f5e9'))
                cat_item.setForeground(QColor('#2e7d32'))
            elif new_category == 'PERSONAL':
                cat_item.setBackground(QColor('#fff3e0'))
                cat_item.setForeground(QColor('#e65100'))
            elif new_category == 'COMMUTE':
                cat_item.setBackground(QColor('#e3f2fd'))
                cat_item.setForeground(QColor('#1565c0'))

        # Update business name display
        name_item = self.item(row, 6)
        if name_item:
            business_name = trip.get('business_name', '')
            if new_category == 'BUSINESS' and not business_name:
                name_item.setText('[Unconfirmed Business]')
                name_item.setForeground(QColor('#ff8f00'))
                name_item.setFont(QFont('', -1, -1, True))
            elif not business_name:
                name_item.setText('')
                name_item.setForeground(QColor('#000000'))
                name_item.setFont(QFont())

        # Emit signal for parent to update stats
        self.trip_updated.emit(trip, 'category', new_category)

    def _save_business_mapping(self, address: str, name: str, category: str = None):
        """Save a business name mapping to the unified mapping file"""
        if not address or not name:
            return

        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')

        # Load existing mapping
        mappings = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except:
                pass

        # Build entry with source=manual
        entry = {"name": name, "source": "manual"}
        if category:
            entry["category"] = category

        mappings[address] = entry

        try:
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, indent=2, ensure_ascii=False)
            self.mapping_changed.emit()
        except:
            pass

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

            # Business name - show "Unconfirmed Business" for business trips without a name
            business_name = trip.get('business_name', '')
            name_item = QTableWidgetItem(business_name)
            if category == 'BUSINESS' and not business_name:
                name_item.setText('[Unconfirmed Business]')
                name_item.setForeground(QColor('#ff8f00'))  # Amber/orange for attention
                name_item.setFont(QFont('', -1, -1, True))  # Italic
            self.setItem(row, 6, name_item)

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
        self._undo_stack = []  # Stack of (address, old_mapping) tuples
        self._redo_stack = []  # Stack of (address, old_mapping) tuples
        self._setup_ui()
        self._setup_menu()
        self._restore_window_state()

    def _setup_ui(self):
        self.setWindowTitle("Mileage Analyzer")
        self.setMinimumSize(1200, 800)
        self.resize(1400, 900)  # Default size slightly larger than minimum

        # Apply modern stylesheet
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QPushButton {
                background-color: #1976d2;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #1565c0;
            }
            QPushButton:pressed {
                background-color: #0d47a1;
            }
            QPushButton:disabled {
                background-color: #bdbdbd;
                color: #757575;
            }
            QPushButton#exportBtn {
                background-color: #388e3c;
            }
            QPushButton#exportBtn:hover {
                background-color: #2e7d32;
            }
            QDateEdit {
                padding: 5px 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background: white;
            }
            QDateEdit:focus {
                border-color: #1976d2;
            }
            QCheckBox {
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QTabWidget::pane {
                border: 1px solid #ddd;
                border-radius: 4px;
                background: white;
            }
            QTabBar::tab {
                background: #e0e0e0;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom: 2px solid #1976d2;
            }
            QTabBar::tab:hover:!selected {
                background: #eeeeee;
            }
            QProgressBar {
                border: none;
                border-radius: 4px;
                background: #e0e0e0;
                text-align: center;
            }
            QProgressBar::chunk {
                background: #1976d2;
                border-radius: 4px;
            }
            QStatusBar {
                background: #fafafa;
                border-top: 1px solid #e0e0e0;
            }
            QComboBox {
                padding: 5px 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background: white;
                min-width: 120px;
            }
            QComboBox:focus {
                border-color: #1976d2;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox QAbstractItemView {
                background: white;
                border: 1px solid #ccc;
                selection-background-color: #e3f2fd;
                selection-color: #1976d2;
                padding: 4px;
            }
            QComboBox QAbstractItemView::item {
                padding: 6px 8px;
                min-height: 20px;
            }
            QComboBox QAbstractItemView::item:hover {
                background: #f5f5f5;
            }
            QLineEdit {
                padding: 5px 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background: white;
            }
            QLineEdit:focus {
                border-color: #1976d2;
            }
            QSplitter::handle {
                background: #e0e0e0;
            }
            QSplitter::handle:horizontal {
                width: 3px;
            }
            QSplitter::handle:hover {
                background: #1976d2;
            }
            QToolTip {
                background-color: #424242;
                color: white;
                border: none;
                padding: 5px 8px;
                border-radius: 4px;
                font-size: 12px;
            }
            /* Modern flat table styling */
            QTableWidget {
                background-color: white;
                alternate-background-color: #fafafa;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                gridline-color: #f0f0f0;
                selection-background-color: #e3f2fd;
                selection-color: #1976d2;
            }
            QTableWidget::item {
                padding: 8px;
                border: none;
            }
            QTableWidget::item:selected {
                background-color: #1976d2;
                color: white;
            }
            QTableWidget::item:hover:!selected {
                background-color: #f5f5f5;
            }
            QHeaderView::section {
                background-color: #fafafa;
                color: #424242;
                padding: 10px 8px;
                border: none;
                border-bottom: 2px solid #e0e0e0;
                border-right: 1px solid #f0f0f0;
                font-weight: bold;
            }
            QHeaderView::section:hover {
                background-color: #f0f0f0;
            }
            QHeaderView::section:pressed {
                background-color: #e0e0e0;
            }
            /* Modern flat tree widget styling */
            QTreeWidget {
                background-color: white;
                alternate-background-color: #fafafa;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                selection-background-color: #e3f2fd;
                selection-color: #1976d2;
            }
            QTreeWidget::item {
                padding: 6px 4px;
                border: none;
            }
            QTreeWidget::item:selected {
                background-color: #1976d2;
                color: white;
            }
            QTreeWidget::item:hover:!selected {
                background-color: #f5f5f5;
            }
            QTreeWidget::branch {
                background: transparent;
            }
            QTreeWidget::branch:has-children:!has-siblings:closed,
            QTreeWidget::branch:closed:has-children:has-siblings {
                border-image: none;
                image: none;
            }
            QTreeWidget::branch:open:has-children:!has-siblings,
            QTreeWidget::branch:open:has-children:has-siblings {
                border-image: none;
                image: none;
            }
            /* Context menu styling */
            QMenu {
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                padding: 4px 0px;
            }
            QMenu::item {
                padding: 8px 24px;
                border: none;
            }
            QMenu::item:selected {
                background-color: #e3f2fd;
                color: #1976d2;
            }
            QMenu::separator {
                height: 1px;
                background: #e0e0e0;
                margin: 4px 8px;
            }
            /* Scrollbar styling */
            QScrollBar:vertical {
                background: #fafafa;
                width: 12px;
                border: none;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #bdbdbd;
                border-radius: 6px;
                min-height: 30px;
            }
            QScrollBar::handle:vertical:hover {
                background: #9e9e9e;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar:horizontal {
                background: #fafafa;
                height: 12px;
                border: none;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal {
                background: #bdbdbd;
                border-radius: 6px;
                min-width: 30px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #9e9e9e;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
            }
            /* Group box styling */
            QGroupBox {
                font-weight: bold;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                margin-top: 12px;
                padding-top: 8px;
                background: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #424242;
            }
        """)

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

        # Left side: Unified trip view (no tabs needed now)
        self.unified_view = UnifiedTripView()
        self.unified_view.trip_selected.connect(self._on_trip_selected)
        self.unified_view.trip_updated.connect(self._on_trip_updated)
        self.unified_view.mapping_changed.connect(self._on_mapping_saved)
        self.unified_view.show_daily_journey.connect(self._on_show_daily_journey)

        splitter.addWidget(self.unified_view)

        # Right side: Map and Summary tabs
        self.right_tabs = QTabWidget()

        # Map tab
        self.map_view = MapView()
        self.map_view.business_selected.connect(self._on_business_selected_from_map)
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

        # Store splitter reference
        self.main_splitter = splitter

        # Set minimum widths to prevent collapse
        self.unified_view.setMinimumWidth(300)
        self.right_tabs.setMinimumWidth(400)

        main_layout.addWidget(splitter)

        # Set initial sizes (50/50 split)
        splitter.setSizes([600, 600])

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # Progress bar (hidden by default)
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(200)
        self.progress_bar.hide()
        self.status_bar.addPermanentWidget(self.progress_bar)

        # Cancel button for long operations (hidden by default)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setMaximumWidth(70)
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #d32f2f;
                padding: 4px 8px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #b71c1c;
            }
        """)
        self.cancel_btn.clicked.connect(self._cancel_operation)
        self.cancel_btn.hide()
        self.status_bar.addPermanentWidget(self.cancel_btn)

        self.status_bar.showMessage("Welcome! Click 'Open Trip Log' or press Ctrl+O to load a Volvo trip log file.")

    def _create_toolbar(self):
        """Create the toolbar with main actions"""
        toolbar = QFrame()
        toolbar.setFrameStyle(QFrame.Shape.StyledPanel)
        toolbar.setFixedHeight(50)
        toolbar.setStyleSheet("""
            QFrame {
                background: white;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
            }
            QLabel {
                color: #555;
                font-weight: 500;
            }
        """)
        layout = QHBoxLayout(toolbar)
        layout.setContentsMargins(12, 6, 12, 6)
        layout.setSpacing(10)

        # Open file button
        open_btn = QPushButton("Open Trip Log")
        open_btn.setToolTip(
            "Open a Volvo trip log file for analysis.\n\n"
            "Supported formats:\n"
            "• CSV - Comma-separated values exported from Volvo On Call\n"
            "• XLSX - Excel spreadsheet format\n\n"
            "Shortcut: Ctrl+O"
        )
        open_btn.setShortcut("Ctrl+O")
        open_btn.clicked.connect(self._open_file)
        layout.addWidget(open_btn)

        # Separator
        sep1 = QFrame()
        sep1.setFrameShape(QFrame.Shape.VLine)
        sep1.setStyleSheet("background: #e0e0e0;")
        layout.addWidget(sep1)

        # Date range
        layout.addWidget(QLabel("From:"))
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate().addDays(-30))
        self.start_date.setToolTip(
            "Start date for filtering trips.\n\n"
            "Only trips on or after this date will be included.\n"
            "Click the calendar icon to select a date."
        )
        layout.addWidget(self.start_date)

        layout.addWidget(QLabel("To:"))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())
        self.end_date.setToolTip(
            "End date for filtering trips.\n\n"
            "Only trips on or before this date will be included.\n"
            "Click the calendar icon to select a date."
        )
        layout.addWidget(self.end_date)

        # Date presets dropdown
        self.date_presets = QComboBox()
        self.date_presets.addItems([
            "Custom",
            "Last 30 Days",
            "This Month",
            "Last Month",
            "This Quarter",
            "Last Quarter",
            "Year to Date",
            "Last Year",
            "All Data"
        ])
        self.date_presets.setToolTip(
            "Quick date range presets.\n\n"
            "Select a preset to automatically set the date range.\n"
            "Choose 'Custom' to manually set dates."
        )
        self.date_presets.setMinimumWidth(110)
        self.date_presets.currentTextChanged.connect(self._on_date_preset_changed)
        layout.addWidget(self.date_presets)

        # Connect date changes to re-run analysis (if a file is loaded)
        self.start_date.dateChanged.connect(self._on_date_range_changed)
        self.end_date.dateChanged.connect(self._on_date_range_changed)

        # Separator
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.VLine)
        sep2.setStyleSheet("background: #e0e0e0;")
        layout.addWidget(sep2)

        # Business lookup checkbox - default to enabled for automatic lookup
        self.lookup_checkbox = QCheckBox("Business Lookup")
        self.lookup_checkbox.setChecked(True)  # Enable by default
        self.lookup_checkbox.setToolTip(
            "Enable automatic business name lookup.\n\n"
            "When enabled, uses Google Places API to identify businesses\n"
            "at each destination address. This helps categorize trips\n"
            "as Business vs Personal automatically.\n\n"
            "Results are cached locally for faster subsequent loads.\n"
            "Requires a Google API key configured in mileage_config.json."
        )
        layout.addWidget(self.lookup_checkbox)

        # Analyze button
        analyze_btn = QPushButton("Analyze")
        analyze_btn.setToolTip(
            "Run or re-run the mileage analysis.\n\n"
            "Processes the loaded trip data to:\n"
            "• Merge false stops (traffic lights, brief stops)\n"
            "• Flag micro-trips (GPS drift)\n"
            "• Categorize trips as Business/Personal/Commute\n"
            "• Calculate totals and statistics\n\n"
            "Shortcut: Ctrl+R"
        )
        analyze_btn.setShortcut("Ctrl+R")
        analyze_btn.clicked.connect(self._run_analysis)
        layout.addWidget(analyze_btn)

        layout.addStretch()

        # Export button
        export_btn = QPushButton("Export to Excel")
        export_btn.setObjectName("exportBtn")  # For custom styling
        export_btn.setToolTip(
            "Export analysis results to an Excel spreadsheet.\n\n"
            "Creates a multi-sheet workbook containing:\n"
            "• Summary - Overall mileage totals and statistics\n"
            "• By Category - Breakdown by Business/Personal/Commute\n"
            "• Detailed Trips - Every trip with full details\n"
            "• Weekly Summary - Week-by-week breakdown\n\n"
            "Shortcut: Ctrl+E"
        )
        export_btn.setShortcut("Ctrl+E")
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

        # Edit menu
        edit_menu = menubar.addMenu("&Edit")

        self.undo_action = QAction("&Undo Mapping Change", self)
        self.undo_action.setShortcut("Ctrl+Z")
        self.undo_action.triggered.connect(self._undo_mapping)
        self.undo_action.setEnabled(False)
        edit_menu.addAction(self.undo_action)

        self.redo_action = QAction("&Redo Mapping Change", self)
        self.redo_action.setShortcut("Ctrl+Y")
        self.redo_action.triggered.connect(self._redo_mapping)
        self.redo_action.setEnabled(False)
        edit_menu.addAction(self.redo_action)

        # View menu
        view_menu = menubar.addMenu("&View")

        refresh_action = QAction("&Refresh Analysis", self)
        refresh_action.setShortcut("F5")
        refresh_action.triggered.connect(self._run_analysis)
        view_menu.addAction(refresh_action)

        view_menu.addSeparator()

        self.dark_mode_action = QAction("&Dark Mode", self)
        self.dark_mode_action.setCheckable(True)
        self.dark_mode_action.setShortcut("Ctrl+D")
        self.dark_mode_action.triggered.connect(self._toggle_dark_mode)
        view_menu.addAction(self.dark_mode_action)

        # Tools menu
        tools_menu = menubar.addMenu("&Tools")

        settings_action = QAction("&Settings...", self)
        settings_action.setShortcut("Ctrl+,")
        settings_action.triggered.connect(self._show_settings)
        tools_menu.addAction(settings_action)

        tools_menu.addSeparator()

        mapping_action = QAction("Edit Business &Mappings...", self)
        mapping_action.triggered.connect(self._edit_mappings)
        tools_menu.addAction(mapping_action)

        tools_menu.addSeparator()

        clear_api_action = QAction("Clear &API Lookups...", self)
        clear_api_action.setToolTip("Remove API lookup entries to allow re-lookup")
        clear_api_action.triggered.connect(self._clear_api_lookups)
        tools_menu.addAction(clear_api_action)

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
            self._dates_initialized = False  # Reset so dates get set from new file
            self.status_bar.showMessage(f"Loading: {os.path.basename(file_path)}...")
            # First load trips quickly without lookup, then auto-start lookup
            self._run_analysis(enable_lookup=False, auto_continue=True)

    def _on_date_range_changed(self):
        """Re-run analysis when date range is changed (if a file is loaded)"""
        # When dates are manually changed, set preset to "Custom"
        if hasattr(self, 'date_presets') and not getattr(self, '_updating_dates', False):
            self.date_presets.blockSignals(True)
            self.date_presets.setCurrentText("Custom")
            self.date_presets.blockSignals(False)

        if self.current_file:
            # Re-filter with current lookup setting, no auto-continue
            self._run_analysis()

    def _on_date_preset_changed(self, preset: str):
        """Handle date preset selection"""
        if preset == "Custom":
            return  # User wants manual control

        today = QDate.currentDate()
        start_date = today
        end_date = today

        if preset == "Last 30 Days":
            start_date = today.addDays(-30)
            end_date = today

        elif preset == "This Month":
            start_date = QDate(today.year(), today.month(), 1)
            end_date = today

        elif preset == "Last Month":
            first_of_this_month = QDate(today.year(), today.month(), 1)
            end_date = first_of_this_month.addDays(-1)
            start_date = QDate(end_date.year(), end_date.month(), 1)

        elif preset == "This Quarter":
            quarter = (today.month() - 1) // 3
            start_month = quarter * 3 + 1
            start_date = QDate(today.year(), start_month, 1)
            end_date = today

        elif preset == "Last Quarter":
            quarter = (today.month() - 1) // 3
            if quarter == 0:
                # Q4 of last year
                start_date = QDate(today.year() - 1, 10, 1)
                end_date = QDate(today.year() - 1, 12, 31)
            else:
                start_month = (quarter - 1) * 3 + 1
                end_month = quarter * 3
                start_date = QDate(today.year(), start_month, 1)
                # Last day of the quarter
                if end_month == 3:
                    end_date = QDate(today.year(), 3, 31)
                elif end_month == 6:
                    end_date = QDate(today.year(), 6, 30)
                elif end_month == 9:
                    end_date = QDate(today.year(), 9, 30)

        elif preset == "Year to Date":
            start_date = QDate(today.year(), 1, 1)
            end_date = today

        elif preset == "Last Year":
            start_date = QDate(today.year() - 1, 1, 1)
            end_date = QDate(today.year() - 1, 12, 31)

        elif preset == "All Data":
            # Set to a very wide range; the analysis will clip to actual data
            start_date = QDate(2000, 1, 1)
            end_date = today

        # Update date pickers without triggering analysis twice
        self._updating_dates = True
        self.start_date.blockSignals(True)
        self.end_date.blockSignals(True)

        self.start_date.setDate(start_date)
        self.end_date.setDate(end_date)

        self.start_date.blockSignals(False)
        self.end_date.blockSignals(False)
        self._updating_dates = False

        # Now trigger the analysis once
        if self.current_file:
            self._run_analysis()

    def _run_analysis(self, enable_lookup=None, auto_continue=False):
        """Run the mileage analysis

        Args:
            enable_lookup: Override for business lookup. If None, uses checkbox state.
            auto_continue: If True, automatically run business lookup after initial load.
        """
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please open a trip log file first.")
            return

        # Stop any existing worker before starting a new one
        if hasattr(self, 'worker') and self.worker is not None:
            if self.worker.isRunning():
                self.worker.terminate()
                self.worker.wait(3000)  # Wait up to 3 seconds
            # Disconnect old signals to prevent duplicate connections
            try:
                self.worker.finished.disconnect()
                self.worker.progress.disconnect()
                self.worker.error.disconnect()
            except (TypeError, RuntimeError):
                pass  # Signals already disconnected or worker deleted
            self.worker = None

        # Store whether to auto-continue with lookup after this pass
        self._auto_continue_lookup = auto_continue

        # Get date range
        start_date = self.start_date.date().toString("yyyy-MM-dd")
        end_date = self.end_date.date().toString("yyyy-MM-dd")

        # Determine lookup setting
        if enable_lookup is None:
            enable_lookup = self.lookup_checkbox.isChecked()

        # Show progress and cancel button
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.progress_bar.show()
        self.cancel_btn.show()

        # Run analysis in background thread
        self.worker = AnalysisWorker(
            self.current_file,
            enable_lookup=enable_lookup,
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
        self.analysis_data = data

        # Set date range from file data on first load
        # Default to last 30 days from max date in file (or full range if shorter)
        date_range = data.get('date_range', {})
        if date_range.get('min') and not getattr(self, '_dates_initialized', False):
            self._dates_initialized = True
            # Block signals to prevent re-running analysis
            self.start_date.blockSignals(True)
            self.end_date.blockSignals(True)

            min_date = QDate.fromString(date_range['min'], 'yyyy-MM-dd')
            max_date = QDate.fromString(date_range['max'], 'yyyy-MM-dd')

            # Set end date to max date in file
            self.end_date.setDate(max_date)

            # Set start date to 30 days before max date, or min date if range is shorter
            start_30_days = max_date.addDays(-30)
            if start_30_days < min_date:
                self.start_date.setDate(min_date)
            else:
                self.start_date.setDate(start_30_days)

            self.start_date.blockSignals(False)
            self.end_date.blockSignals(False)

        # Update unified trip view
        trips = data.get('trips', [])
        self.unified_view.load_trips(trips)

        # Update summary
        self.summary_widget.update_stats(data)

        # Update weekly breakdown text
        self._update_weekly_text(data)

        # Update map with trips
        self.map_view.show_trips(trips)

        # Count destinations needing names
        needs_name_count = sum(1 for g in self.unified_view.grouped_data if g['status'] in ['Needs Name', 'Unconfirmed Business'])

        # Check if we should auto-continue with business lookup
        if getattr(self, '_auto_continue_lookup', False) and self.lookup_checkbox.isChecked():
            self._auto_continue_lookup = False  # Reset flag
            self.status_bar.showMessage(f"Loaded {len(trips)} trips. Starting business lookup...")
            # Run again with lookup enabled
            self._run_analysis(enable_lookup=True, auto_continue=False)
        else:
            self.progress_bar.hide()
            self.cancel_btn.hide()
            dest_count = len(self.unified_view.grouped_data)
            self.status_bar.showMessage(f"Analysis complete. {len(trips)} trips to {dest_count} destinations. {needs_name_count} need names.")

    def _on_analysis_error(self, error: str):
        """Handle analysis error"""
        self.progress_bar.hide()
        self.cancel_btn.hide()
        QMessageBox.critical(self, "Analysis Error", f"Error during analysis:\n{error}")
        self.status_bar.showMessage("Analysis failed.")

    def _cancel_operation(self):
        """Cancel the current background operation"""
        if hasattr(self, 'worker') and self.worker is not None and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait(3000)  # Wait up to 3 seconds
            self.worker = None
        self.progress_bar.hide()
        self.cancel_btn.hide()
        self.status_bar.showMessage("Operation cancelled.")

    def _update_weekly_text(self, data: dict):
        """Update the weekly breakdown with HTML formatting"""
        weekly_stats = data.get('weekly_stats', {})

        html = """
        <html>
        <head>
        <style>
            body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; background: #fafafa; }
            h2 { color: #333; margin-bottom: 20px; }
            table { border-collapse: collapse; width: 100%; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
            th { background: #424242; color: white; padding: 12px 15px; text-align: right; font-weight: 600; }
            th:first-child { text-align: left; }
            td { padding: 10px 15px; border-bottom: 1px solid #eee; text-align: right; }
            td:first-child { text-align: left; font-weight: 500; }
            tr:hover { background: #f5f5f5; }
            tr:last-child td { border-bottom: none; }
            .commute { color: #1565c0; }
            .business { color: #2e7d32; }
            .personal { color: #e65100; }
            .total { font-weight: bold; color: #333; }
            .trips { color: #666; font-size: 0.9em; }
            tfoot td { background: #f5f5f5; font-weight: bold; border-top: 2px solid #424242; }
        </style>
        </head>
        <body>
        <h2>Weekly Mileage Breakdown</h2>
        <table>
        <thead>
            <tr>
                <th>Week Starting</th>
                <th>Trips</th>
                <th>Commute</th>
                <th>Business</th>
                <th>Personal</th>
                <th>Total</th>
            </tr>
        </thead>
        <tbody>
        """

        total_trips = 0
        total_commute = 0
        total_business = 0
        total_personal = 0
        total_all = 0

        for week in sorted(weekly_stats.keys()):
            stats = weekly_stats[week]
            trips = len(stats.get('trips', []))
            commute = stats.get('commute', 0)
            business = stats.get('business', 0)
            personal = stats.get('personal', 0)
            total = stats.get('total', 0)

            total_trips += trips
            total_commute += commute
            total_business += business
            total_personal += personal
            total_all += total

            html += f"""
            <tr>
                <td>{week}</td>
                <td class="trips">{trips}</td>
                <td class="commute">{commute:.1f}</td>
                <td class="business">{business:.1f}</td>
                <td class="personal">{personal:.1f}</td>
                <td class="total">{total:.1f}</td>
            </tr>
            """

        html += f"""
        </tbody>
        <tfoot>
            <tr>
                <td>TOTALS</td>
                <td class="trips">{total_trips}</td>
                <td class="commute">{total_commute:.1f}</td>
                <td class="business">{total_business:.1f}</td>
                <td class="personal">{total_personal:.1f}</td>
                <td class="total">{total_all:.1f}</td>
            </tr>
        </tfoot>
        </table>
        </body>
        </html>
        """

        self.weekly_text.setHtml(html)

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
        """Handle trip selection - show route on map"""
        self.right_tabs.setCurrentIndex(0)  # Switch to Map tab
        # Show the full route with directions
        self.map_view.show_route(trip)
        # Show what's being viewed in status bar
        start = trip.get('start_address', 'Unknown')[:30]
        end = trip.get('end_address', 'Unknown')[:30]
        dist = trip.get('distance', 0)
        self.status_bar.showMessage(f"Showing route: {start}... -> {end}... ({dist:.1f} mi)")

    def _on_business_selected_from_map(self, business_name: str):
        """Handle business name selected from map click"""
        if not business_name:
            return

        # Get currently selected trip
        selected_trip = self.map_view.selected_trip
        if not selected_trip:
            self.status_bar.showMessage(f"Selected business: {business_name} (no trip selected to update)")
            return

        # Update the trip's business name
        selected_trip['business_name'] = business_name

        # Also save to business mapping for future auto-categorization
        end_address = selected_trip.get('end_address', '')
        if end_address:
            self._save_business_mapping(end_address, business_name)

        # Refresh the view
        self.unified_view._refresh_table()
        self.status_bar.showMessage(f"Set business name to: {business_name}", 5000)

    def _backup_file(self, file_path: str):
        """Create a backup of a file before modifying it"""
        if not os.path.exists(file_path):
            return

        try:
            # Create backup with .bak extension
            backup_path = file_path + '.bak'
            import shutil
            shutil.copy2(file_path, backup_path)
        except Exception as e:
            # Backup failure shouldn't prevent saving
            print(f"Warning: Could not create backup: {e}")

    def _save_business_mapping(self, address: str, business_name: str, category: str = None, track_undo: bool = True):
        """Save address to business name mapping"""
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')

        # Load existing mapping
        mappings = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except json.JSONDecodeError as e:
                self.status_bar.showMessage(f"Warning: Corrupt mapping file, starting fresh", 5000)
            except Exception as e:
                self.status_bar.showMessage(f"Warning: Could not load mappings: {e}", 5000)

        # Track previous value for undo (before making changes)
        if track_undo:
            old_value = mappings.get(address)  # None if didn't exist
            self._undo_stack.append((address, old_value))
            self._redo_stack.clear()  # Clear redo when new action is performed
            self._update_undo_actions()

        # Build entry with source=manual
        entry = {"name": business_name, "source": "manual"}
        if category:
            entry["category"] = category

        mappings[address] = entry

        try:
            # Create backup before saving
            self._backup_file(mapping_file)

            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, indent=2, ensure_ascii=False)
            self.status_bar.showMessage(f"Saved mapping: {business_name}", 3000)
        except PermissionError:
            QMessageBox.warning(self, "Save Error",
                f"Cannot save mapping - file is in use.\n\nClose any other programs using:\n{mapping_file}")
        except Exception as e:
            QMessageBox.warning(self, "Save Error", f"Failed to save mapping:\n{e}")

    def _update_undo_actions(self):
        """Update the enabled state of undo/redo menu actions"""
        self.undo_action.setEnabled(len(self._undo_stack) > 0)
        self.redo_action.setEnabled(len(self._redo_stack) > 0)

    def _undo_mapping(self):
        """Undo the last mapping change"""
        if not self._undo_stack:
            return

        address, old_value = self._undo_stack.pop()
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')

        # Load current mappings
        mappings = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except:
                pass

        # Save current value to redo stack
        current_value = mappings.get(address)
        self._redo_stack.append((address, current_value))

        # Restore old value
        if old_value is None:
            # Remove the mapping (it didn't exist before)
            if address in mappings:
                del mappings[address]
            self.status_bar.showMessage(f"Undone: Removed mapping for address", 3000)
        else:
            mappings[address] = old_value
            name = old_value.get('name', '') if isinstance(old_value, dict) else old_value
            self.status_bar.showMessage(f"Undone: Restored mapping to '{name}'", 3000)

        # Save mappings
        try:
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, indent=2, ensure_ascii=False)
        except Exception as e:
            QMessageBox.warning(self, "Undo Error", f"Failed to undo:\n{e}")

        self._update_undo_actions()

        # Refresh the analysis to show updated categorization
        if self.current_file:
            self._run_analysis()

    def _redo_mapping(self):
        """Redo the last undone mapping change"""
        if not self._redo_stack:
            return

        address, old_value = self._redo_stack.pop()
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')

        # Load current mappings
        mappings = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except:
                pass

        # Save current value to undo stack
        current_value = mappings.get(address)
        self._undo_stack.append((address, current_value))

        # Restore the redo value
        if old_value is None:
            if address in mappings:
                del mappings[address]
            self.status_bar.showMessage(f"Redone: Removed mapping", 3000)
        else:
            mappings[address] = old_value
            name = old_value.get('name', '') if isinstance(old_value, dict) else old_value
            self.status_bar.showMessage(f"Redone: Restored mapping to '{name}'", 3000)

        # Save mappings
        try:
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, indent=2, ensure_ascii=False)
        except Exception as e:
            QMessageBox.warning(self, "Redo Error", f"Failed to redo:\n{e}")

        self._update_undo_actions()

        # Refresh the analysis to show updated categorization
        if self.current_file:
            self._run_analysis()

    def _on_show_daily_journey(self, trips: list):
        """Handle show daily journey request"""
        if trips:
            self.right_tabs.setCurrentIndex(0)  # Switch to Map tab
            # Get the date for status message
            trip_date = trips[0].get('started')
            date_str = trip_date.strftime('%Y-%m-%d') if hasattr(trip_date, 'strftime') else str(trip_date)
            self.status_bar.showMessage(f"Showing {len(trips)} trips for {date_str}", 5000)
            self.map_view.show_daily_journey(trips)

    def _on_address_selected_for_map(self, address: str, lat: float, lng: float):
        """Handle address selection from unresolved list - show on embedded map"""
        if address:
            self.right_tabs.setCurrentIndex(0)  # Switch to Map tab
            self.map_view.show_address(address)

    def _on_mapping_saved(self):
        """Handle when a business mapping is saved - refresh analysis"""
        if self.current_file:
            self._run_analysis()

    def _on_trip_updated(self, trip: dict, field: str, value: str):
        """Handle when a trip is updated in the table - recalculate summary"""
        if self.analysis_data:
            # Recalculate totals from the trips data
            trips = self.analysis_data.get('trips', [])
            totals = {'business_miles': 0, 'personal_miles': 0, 'commute_miles': 0, 'total_miles': 0}

            for t in trips:
                distance = t.get('distance', 0)
                category = t.get('computed_category', 'PERSONAL')
                totals['total_miles'] += distance
                if category == 'BUSINESS':
                    totals['business_miles'] += distance
                elif category == 'PERSONAL':
                    totals['personal_miles'] += distance
                elif category == 'COMMUTE':
                    totals['commute_miles'] += distance

            # Calculate percentages
            total_all = totals['total_miles']
            totals['business_pct'] = (totals['business_miles'] / total_all * 100) if total_all > 0 else 0
            totals['personal_pct'] = (totals['personal_miles'] / total_all * 100) if total_all > 0 else 0
            totals['commute_pct'] = (totals['commute_miles'] / total_all * 100) if total_all > 0 else 0

            self.analysis_data['totals'] = totals
            self.summary_widget.update_stats(self.analysis_data)

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

                # Merge notes into trips data for export
                notes = load_trip_notes()
                for trip in trips:
                    trip_key = get_trip_key(trip)
                    trip['notes'] = notes.get(trip_key, '')

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

    def _show_settings(self):
        """Open settings dialog"""
        dialog = SettingsDialog(self)
        dialog.settings_changed.connect(self._on_settings_changed)
        dialog.exec()

    def _on_settings_changed(self):
        """Handle settings changes"""
        # Reload config in analyze_mileage module if it's been imported
        try:
            from analyze_mileage import load_config
            load_config()
        except:
            pass
        self.status_bar.showMessage("Settings updated. Re-run analysis to apply changes.")

    def _edit_mappings(self):
        """Open business mappings editor"""
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        # Create file if it doesn't exist
        if not os.path.exists(mapping_file):
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump({}, f)

        self.mapping_editor = JsonEditorDialog(
            mapping_file,
            "Business Mappings Editor",
            self
        )
        self.mapping_editor.data_saved.connect(self._on_mapping_saved)
        self.mapping_editor.show()

    def _clear_api_lookups(self):
        """Clear API lookup entries from business mapping, keeping manual entries"""
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        if not os.path.exists(mapping_file):
            QMessageBox.information(self, "No Mappings", "Business mapping file is empty.")
            return

        # Load current mapping
        try:
            with open(mapping_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except:
            QMessageBox.critical(self, "Error", "Failed to load business mapping file.")
            return

        # Count API entries
        api_count = 0
        manual_count = 0
        for addr, value in data.items():
            if isinstance(value, dict):
                source = value.get('source', 'manual')
                if source in ['google_api', 'osm_api']:
                    api_count += 1
                else:
                    manual_count += 1
            else:
                manual_count += 1

        if api_count == 0:
            QMessageBox.information(self, "No API Entries",
                "No API lookup entries found.\n"
                f"All {manual_count} entries are manual.")
            return

        reply = QMessageBox.question(
            self, "Confirm Clear API Lookups",
            f"This will remove {api_count} API lookup entries.\n"
            f"{manual_count} manual entries will be preserved.\n\n"
            "The next analysis will re-lookup the cleared addresses.\n\n"
            "Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            # Keep only manual entries
            new_data = {}
            for addr, value in data.items():
                if isinstance(value, dict):
                    source = value.get('source', 'manual')
                    if source not in ['google_api', 'osm_api']:
                        new_data[addr] = value
                else:
                    new_data[addr] = value

            try:
                with open(mapping_file, 'w', encoding='utf-8') as f:
                    json.dump(new_data, f, indent=2, ensure_ascii=False)
                QMessageBox.information(self, "API Lookups Cleared",
                    f"Removed {api_count} API lookup entries.\n"
                    f"Kept {len(new_data)} manual entries."
                )
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save:\n{e}")

    def _toggle_dark_mode(self, enabled: bool):
        """Toggle dark mode on/off"""
        if enabled:
            self._apply_dark_style()
        else:
            self._apply_light_style()
        # Save preference
        settings = self._get_settings()
        settings.setValue("darkMode", enabled)

    def _apply_dark_style(self):
        """Apply dark mode stylesheet"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1e1e1e;
            }
            QWidget {
                background-color: #1e1e1e;
                color: #e0e0e0;
            }
            QPushButton {
                background-color: #0d47a1;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #1565c0;
            }
            QPushButton:pressed {
                background-color: #1976d2;
            }
            QPushButton:disabled {
                background-color: #424242;
                color: #757575;
            }
            QPushButton#exportBtn {
                background-color: #2e7d32;
            }
            QPushButton#exportBtn:hover {
                background-color: #388e3c;
            }
            QDateEdit, QComboBox, QLineEdit {
                padding: 5px 8px;
                border: 1px solid #424242;
                border-radius: 4px;
                background: #2d2d2d;
                color: #e0e0e0;
            }
            QDateEdit:focus, QComboBox:focus, QLineEdit:focus {
                border-color: #1976d2;
            }
            QComboBox QAbstractItemView {
                background: #2d2d2d;
                border: 1px solid #424242;
                selection-background-color: #0d47a1;
                selection-color: white;
            }
            QTabWidget::pane {
                border: 1px solid #424242;
                border-radius: 4px;
                background: #252525;
            }
            QTabBar::tab {
                background: #2d2d2d;
                color: #b0b0b0;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background: #252525;
                color: white;
                border-bottom: 2px solid #1976d2;
            }
            QTabBar::tab:hover:!selected {
                background: #353535;
            }
            QTableWidget, QTreeWidget {
                background-color: #252525;
                alternate-background-color: #2d2d2d;
                color: #e0e0e0;
                gridline-color: #3d3d3d;
                border: 1px solid #424242;
            }
            QTableWidget::item:selected, QTreeWidget::item:selected {
                background-color: #0d47a1;
                color: white;
            }
            QHeaderView::section {
                background-color: #2d2d2d;
                color: #e0e0e0;
                padding: 6px;
                border: none;
                border-bottom: 1px solid #424242;
            }
            QProgressBar {
                border: none;
                border-radius: 4px;
                background: #3d3d3d;
                text-align: center;
            }
            QProgressBar::chunk {
                background: #1976d2;
                border-radius: 4px;
            }
            QStatusBar {
                background: #252525;
                border-top: 1px solid #424242;
                color: #b0b0b0;
            }
            QMenuBar {
                background: #2d2d2d;
                color: #e0e0e0;
            }
            QMenuBar::item:selected {
                background: #424242;
            }
            QMenu {
                background: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #424242;
            }
            QMenu::item:selected {
                background: #0d47a1;
            }
            QScrollBar:vertical {
                background: #2d2d2d;
                width: 12px;
                border: none;
            }
            QScrollBar::handle:vertical {
                background: #5d5d5d;
                min-height: 20px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical:hover {
                background: #6d6d6d;
            }
            QScrollBar:horizontal {
                background: #2d2d2d;
                height: 12px;
                border: none;
            }
            QScrollBar::handle:horizontal {
                background: #5d5d5d;
                min-width: 20px;
                border-radius: 6px;
            }
            QFrame {
                background: #252525;
                border-color: #424242;
            }
            QTextEdit {
                background: #252525;
                color: #e0e0e0;
                border: 1px solid #424242;
            }
            QLabel {
                color: #e0e0e0;
                background: transparent;
            }
            QCheckBox {
                color: #e0e0e0;
            }
            QToolTip {
                background-color: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #424242;
            }
            QSplitter::handle {
                background: #424242;
            }
            QSplitter::handle:hover {
                background: #1976d2;
            }
        """)

    def _apply_light_style(self):
        """Apply light mode stylesheet (restore default)"""
        # Restore the original light theme from _setup_ui
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QPushButton {
                background-color: #1976d2;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #1565c0;
            }
            QPushButton:pressed {
                background-color: #0d47a1;
            }
            QPushButton:disabled {
                background-color: #bdbdbd;
                color: #757575;
            }
            QPushButton#exportBtn {
                background-color: #388e3c;
            }
            QPushButton#exportBtn:hover {
                background-color: #2e7d32;
            }
            QDateEdit {
                padding: 5px 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background: white;
            }
            QDateEdit:focus {
                border-color: #1976d2;
            }
            QCheckBox {
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QTabWidget::pane {
                border: 1px solid #ddd;
                border-radius: 4px;
                background: white;
            }
            QTabBar::tab {
                background: #e0e0e0;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom: 2px solid #1976d2;
            }
            QTabBar::tab:hover:!selected {
                background: #eeeeee;
            }
            QProgressBar {
                border: none;
                border-radius: 4px;
                background: #e0e0e0;
                text-align: center;
            }
            QProgressBar::chunk {
                background: #1976d2;
                border-radius: 4px;
            }
            QStatusBar {
                background: #fafafa;
                border-top: 1px solid #e0e0e0;
            }
            QComboBox {
                padding: 5px 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background: white;
                min-width: 120px;
            }
            QComboBox:focus {
                border-color: #1976d2;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox QAbstractItemView {
                background: white;
                border: 1px solid #ccc;
                selection-background-color: #e3f2fd;
                selection-color: #1976d2;
                padding: 4px;
            }
            QLineEdit {
                padding: 5px 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background: white;
            }
            QLineEdit:focus {
                border-color: #1976d2;
            }
            QSplitter::handle {
                background: #e0e0e0;
            }
            QSplitter::handle:horizontal {
                width: 3px;
            }
            QSplitter::handle:hover {
                background: #1976d2;
            }
            QToolTip {
                background-color: #424242;
                color: white;
                border: none;
                padding: 5px 8px;
                border-radius: 4px;
                font-size: 12px;
            }
            QTableWidget {
                background-color: white;
                alternate-background-color: #fafafa;
                border: 1px solid #e0e0e0;
                gridline-color: #f0f0f0;
            }
            QTableWidget::item:selected {
                background-color: #e3f2fd;
                color: #1976d2;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 6px;
                border: none;
                border-bottom: 1px solid #e0e0e0;
                font-weight: 600;
                color: #424242;
            }
        """)

    def _show_about(self):
        """Show about dialog"""
        QMessageBox.about(
            self, "About Mileage Analyzer",
            "<h2>Mileage Analyzer</h2>"
            "<p>Version 3.0</p>"
            "<p>&copy; 2025 Kevin Whittenberger</p>"
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

    def _get_settings(self) -> QSettings:
        """Get QSettings object for storing application settings"""
        return QSettings("DewartRepresentatives", "MileageAnalyzer")

    def _save_window_state(self):
        """Save window geometry and splitter state"""
        settings = self._get_settings()
        settings.setValue("geometry", self.saveGeometry())
        settings.setValue("windowState", self.saveState())
        if hasattr(self, 'main_splitter'):
            settings.setValue("splitterState", self.main_splitter.saveState())
        if hasattr(self, 'right_tabs'):
            settings.setValue("rightTabIndex", self.right_tabs.currentIndex())

    def _restore_window_state(self):
        """Restore window geometry and splitter state"""
        settings = self._get_settings()

        # Restore window geometry
        geometry = settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)

        # Restore window state
        state = settings.value("windowState")
        if state:
            self.restoreState(state)

        # Restore splitter state
        if hasattr(self, 'main_splitter'):
            splitter_state = settings.value("splitterState")
            if splitter_state:
                self.main_splitter.restoreState(splitter_state)

        # Restore active tab on right side
        if hasattr(self, 'right_tabs'):
            tab_index = settings.value("rightTabIndex", type=int)
            if tab_index is not None and 0 <= tab_index < self.right_tabs.count():
                self.right_tabs.setCurrentIndex(tab_index)

        # Restore dark mode setting
        dark_mode = settings.value("darkMode", False, type=bool)
        if dark_mode:
            self.dark_mode_action.setChecked(True)
            self._apply_dark_style()

    def closeEvent(self, event):
        """Handle window close - save state before closing"""
        self._save_window_state()
        super().closeEvent(event)


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
