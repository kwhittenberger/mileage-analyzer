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
    QTreeWidget, QTreeWidgetItem, QAbstractItemView
)
from PyQt6.QtCore import Qt, QDate, QThread, pyqtSignal, QUrl
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

                // Draw the route
                const routeLine = new google.maps.Polyline({{
                    path: decodedPath,
                    geodesic: true,
                    strokeColor: color,
                    strokeWeight: 5,
                    strokeOpacity: 0.8
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
                        console.error('Route failed for trip', i, data.error);
                        continue;
                    }}

                    const route = data.routes[0];
                    const decodedPath = google.maps.geometry.encoding.decodePath(route.polyline.encodedPolyline);
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

        # Save all selected addresses with the same name
        self._save_multiple_to_mapping_file(self.selected_addresses, name)
        self.business_name_input.clear()
        self._remove_selected_from_list()
        self.mapping_saved.emit()

    def _mark_personal(self):
        """Mark all selected addresses as personal"""
        if not self.selected_addresses:
            return

        self._save_multiple_to_mapping_file(self.selected_addresses, "[PERSONAL]")
        self._remove_selected_from_list()
        self.mapping_saved.emit()

    def _save_multiple_to_mapping_file(self, addresses: List[str], name: str):
        """Save multiple address mappings to the business_mapping.json file"""
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')

        # Load existing
        mappings = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except:
                pass

        # Add all mappings
        for address in addresses:
            mappings[address] = name

        # Save
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


class JsonEditorDialog(QMainWindow):
    """Dialog for viewing and editing JSON files (address_cache, business_mapping)"""

    data_saved = pyqtSignal()

    def __init__(self, file_path: str, title: str, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.data = {}
        self.setWindowTitle(title)
        self.setMinimumSize(800, 600)
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

        add_btn = QPushButton("Add Entry")
        add_btn.clicked.connect(self._add_entry)
        toolbar.addWidget(add_btn)

        delete_btn = QPushButton("Delete Selected")
        delete_btn.clicked.connect(self._delete_selected)
        toolbar.addWidget(delete_btn)

        layout.addLayout(toolbar)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Address/Key", "Business Name/Value"])
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)

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
        self.table.setRowCount(len(self.data))

        for row, (key, value) in enumerate(sorted(self.data.items())):
            key_item = QTableWidgetItem(str(key))
            self.table.setItem(row, 0, key_item)

            # Handle different value types
            if isinstance(value, dict):
                # For address_cache entries which are dicts with 'name' key
                display_value = value.get('name', str(value))
            else:
                display_value = str(value)

            value_item = QTableWidgetItem(display_value)
            self.table.setItem(row, 1, value_item)

        self.table.blockSignals(False)
        self._update_stats()

    def _update_stats(self):
        """Update the stats label"""
        visible = sum(1 for row in range(self.table.rowCount()) if not self.table.isRowHidden(row))
        total = self.table.rowCount()
        self.stats_label.setText(f"Showing {visible} of {total} entries")

    def _filter_table(self, text: str):
        """Filter table rows based on search text"""
        text = text.lower()
        for row in range(self.table.rowCount()):
            key_item = self.table.item(row, 0)
            value_item = self.table.item(row, 1)
            key_text = key_item.text().lower() if key_item else ""
            value_text = value_item.text().lower() if value_item else ""

            show = text in key_text or text in value_text
            self.table.setRowHidden(row, not show)

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

    def _save_data(self):
        """Save data back to JSON file"""
        new_data = {}

        for row in range(self.table.rowCount()):
            key_item = self.table.item(row, 0)
            value_item = self.table.item(row, 1)

            if key_item and key_item.text().strip():
                key = key_item.text().strip()
                value = value_item.text().strip() if value_item else ""

                # Check if original was a dict (address_cache format)
                if key in self.data and isinstance(self.data[key], dict):
                    # Preserve dict structure, update 'name'
                    new_data[key] = self.data[key].copy()
                    new_data[key]['name'] = value
                else:
                    new_data[key] = value

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

        # Filter bar - two rows to prevent overlap when resizing
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
        self.view_mode_combo.addItems(["By Destination", "All Trips", "By Day"])
        self.view_mode_combo.currentTextChanged.connect(self._on_view_mode_changed)
        row1.addWidget(self.view_mode_combo)

        row1.addWidget(QLabel("Category:"))
        self.category_filter = QComboBox()
        self.category_filter.addItems(["All", "Business", "Personal", "Commute"])
        self.category_filter.currentTextChanged.connect(self._apply_filters)
        row1.addWidget(self.category_filter)

        row1.addWidget(QLabel("Status:"))
        self.status_filter = QComboBox()
        self.status_filter.addItems(["All", "Unresolved", "Resolved", "Unconfirmed Business"])
        self.status_filter.currentTextChanged.connect(self._apply_filters)
        row1.addWidget(self.status_filter)

        row1.addStretch()

        # Stats label on right of row 1
        self.stats_label = QLabel("0 items")
        self.stats_label.setStyleSheet("font-weight: bold; color: #555;")
        row1.addWidget(self.stats_label)

        filter_vlayout.addLayout(row1)

        # Row 2: Business filter and Search
        row2 = QHBoxLayout()
        row2.setSpacing(8)

        row2.addWidget(QLabel("Business:"))
        self.business_filter = QComboBox()
        self.business_filter.addItem("All")
        self.business_filter.setSizePolicy(self.business_filter.sizePolicy().horizontalPolicy(), self.business_filter.sizePolicy().verticalPolicy())
        self.business_filter.currentTextChanged.connect(self._apply_filters)
        row2.addWidget(self.business_filter, 1)  # Stretch factor 1

        row2.addWidget(QLabel("Search:"))
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search...")
        self.search_box.textChanged.connect(self._apply_filters)
        row2.addWidget(self.search_box, 1)  # Stretch factor 1

        filter_vlayout.addLayout(row2)

        layout.addWidget(filter_frame)

        # Main table
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.table.setSortingEnabled(True)
        # Keep selection visible even when table loses focus
        self.table.setStyleSheet("""
            QTableWidget::item:selected {
                background-color: #0078d4;
                color: white;
            }
        """)

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
        # Keep selection visible even when tree loses focus
        self.tree.setStyleSheet("""
            QTreeWidget::item:selected {
                background-color: #0078d4;
                color: white;
            }
        """)
        self.tree.setHeaderLabels(["Date / Trip", "Time", "Category", "Miles", "Destination"])
        self.tree.setColumnWidth(0, 180)
        self.tree.setColumnWidth(1, 60)
        self.tree.setColumnWidth(2, 80)
        self.tree.setColumnWidth(3, 70)
        self.tree.header().setStretchLastSection(True)
        self.tree.itemClicked.connect(self._on_tree_item_clicked)
        self.tree.itemDoubleClicked.connect(self._on_tree_item_double_clicked)
        self.tree.itemExpanded.connect(self._on_tree_item_expanded)
        self.tree.itemSelectionChanged.connect(self._on_tree_selection_changed)
        self.tree.hide()  # Hidden by default
        layout.addWidget(self.tree)

        # Set up for grouped view initially
        self._setup_grouped_columns()

    def _setup_grouped_columns(self):
        """Set up columns for grouped (by destination) view"""
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Destination", "Business Name", "Category", "Trips", "Miles", "Status", "Street"
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(2, 80)
        self.table.setColumnWidth(3, 50)
        self.table.setColumnWidth(4, 70)
        self.table.setColumnWidth(5, 100)
        self.table.setColumnWidth(6, 100)

    def _setup_individual_columns(self):
        """Set up columns for individual trips view"""
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Date", "Day", "Category", "Distance", "From", "To", "Business Name"
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 130)
        self.table.setColumnWidth(1, 50)
        self.table.setColumnWidth(2, 80)
        self.table.setColumnWidth(3, 70)

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
        else:
            self.view_mode = "individual"
            self.table.show()
            self.tree.hide()
            self._setup_individual_columns()
        self._refresh_table()

    def load_trips(self, trips: List[Dict]):
        """Load trip data and group by destination"""
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
            # Destination address
            addr_item = QTableWidgetItem(data['address'])
            addr_item.setData(Qt.ItemDataRole.UserRole, row)  # Store index
            self.table.setItem(row, 0, addr_item)

            # Business name
            name = data.get('business_name', '')
            name_item = QTableWidgetItem(name)
            if data['status'] == 'Unconfirmed Business':
                name_item.setText('[Unconfirmed]')
                name_item.setForeground(QColor('#ff8f00'))
                name_item.setFont(QFont('', -1, -1, True))
            elif data['status'] == 'Needs Name':
                name_item.setText('')
            self.table.setItem(row, 1, name_item)

            # Category
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
            self.table.setItem(row, 2, cat_item)

            # Trip count (use NumericTableWidgetItem for proper sorting)
            trip_count = data['trip_count']
            count_item = NumericTableWidgetItem(str(trip_count), trip_count)
            self.table.setItem(row, 3, count_item)

            # Total miles (use NumericTableWidgetItem for proper sorting)
            total_miles = data['total_miles']
            miles_item = NumericTableWidgetItem(f"{total_miles:.1f}", total_miles)
            self.table.setItem(row, 4, miles_item)

            # Status
            status_item = QTableWidgetItem(data['status'])
            if data['status'] == 'Needs Name':
                status_item.setForeground(QColor('#d32f2f'))
            elif data['status'] == 'Unconfirmed Business':
                status_item.setForeground(QColor('#ff8f00'))
            else:
                status_item.setForeground(QColor('#388e3c'))
            self.table.setItem(row, 5, status_item)

            # Street
            self.table.setItem(row, 6, QTableWidgetItem(data.get('street', '')))

    def _populate_individual_view(self):
        """Populate table with individual trip data"""
        self.table.setRowCount(len(self.trips_data))

        for row, trip in enumerate(self.trips_data):
            # Date
            date_str = trip['started'].strftime('%Y-%m-%d %H:%M')
            date_item = QTableWidgetItem(date_str)
            date_item.setData(Qt.ItemDataRole.UserRole, row)
            self.table.setItem(row, 0, date_item)

            # Day
            self.table.setItem(row, 1, QTableWidgetItem(trip['started'].strftime('%a')))

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
            self.table.setItem(row, 2, cat_item)

            # Distance (use NumericTableWidgetItem for proper sorting)
            dist_val = trip.get('distance', 0)
            dist_item = NumericTableWidgetItem(f"{dist_val:.1f} mi", dist_val)
            self.table.setItem(row, 3, dist_item)

            # From/To
            self.table.setItem(row, 4, QTableWidgetItem(trip.get('start_address', '')))
            self.table.setItem(row, 5, QTableWidgetItem(trip.get('end_address', '')))

            # Business name
            name = trip.get('business_name', '')
            name_item = QTableWidgetItem(name)
            if cat == 'BUSINESS' and not name:
                name_item.setText('[Unconfirmed]')
                name_item.setForeground(QColor('#ff8f00'))
                name_item.setFont(QFont('', -1, -1, True))
            self.table.setItem(row, 6, name_item)

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
            # Single trip selected - show it on map
            self.trip_selected.emit(user_data['data'])
        elif user_data['type'] == 'day':
            # Day selected - show all trips for that day
            day_data = user_data['data']
            if day_data['trips']:
                self.show_daily_journey.emit(day_data['trips'])

    def _on_tree_item_double_clicked(self, item: QTreeWidgetItem, column: int):
        """Handle double click on tree item - expand/collapse or show route"""
        user_data = item.data(0, Qt.ItemDataRole.UserRole)
        if not user_data:
            return

        if user_data['type'] == 'day':
            # Toggle expansion
            item.setExpanded(not item.isExpanded())
        elif user_data['type'] == 'trip':
            # Show trip on map
            self.trip_selected.emit(user_data['data'])

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
            # Show single trip on map
            self.trip_selected.emit(user_data['data'])
        elif user_data['type'] == 'day':
            # Show all trips for the day
            day_data = user_data['data']
            if day_data.get('trips'):
                self.show_daily_journey.emit(day_data['trips'])

    def _apply_filters(self):
        """Apply all filters to the table or tree"""
        category = self.category_filter.currentText()
        status = self.status_filter.currentText()
        business = self.business_filter.currentText()
        search = self.search_box.text().lower()

        # Auto-switch to All Trips view when filtering by Unresolved
        # because grouped view hides individual unresolved trips
        if status == "Unresolved" and self.view_mode == "grouped":
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
                        elif status == "Unconfirmed Business" and data['status'] != 'Unconfirmed Business':
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
                    # Category filter
                    if category != "All" and trip.get('computed_category', '') != category.upper():
                        show = False
                    # Status filter
                    if status != "All":
                        has_name = bool(trip.get('business_name', ''))
                        is_business = trip.get('computed_category') == 'BUSINESS'
                        if status == "Unresolved" and has_name:
                            show = False
                        elif status == "Resolved" and not has_name:
                            show = False
                        elif status == "Unconfirmed Business" and not (is_business and not has_name):
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

            # Update table cell
            name_item = self.table.item(row, 1)
            if name_item:
                if name:
                    name_item.setText(name)
                    name_item.setForeground(QColor('#000000'))
                    name_item.setFont(QFont())
                elif data['status'] == 'Unconfirmed Business':
                    name_item.setText('[Unconfirmed]')
                    name_item.setForeground(QColor('#ff8f00'))
                    name_item.setFont(QFont('', -1, -1, True))
                else:
                    name_item.setText('')

            # Update status cell
            status_item = self.table.item(row, 5)
            if status_item:
                status_item.setText(data['status'])

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

            # Update table
            cat_item = self.table.item(row, 2)
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
            name_item = self.table.item(row, 6)
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
            cat_item = self.table.item(row, 2)
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

        # Existing names
        existing_names = self._get_existing_business_names()
        name_actions = {}
        if existing_names:
            recent_menu = assign_menu.addMenu("Existing Names")
            for name in sorted(existing_names)[:15]:
                action = recent_menu.addAction(name)
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
        elif action == view_map_action:
            self._view_on_map(row)
        elif action == show_day_action:
            self._show_day_journey(row)
        elif action == open_gmaps_action:
            self._open_in_google_maps(row)

    def _apply_name_to_selected(self, data_indices: List[int], name: str):
        """Apply business name to selected rows (data_indices are original data indices)"""
        for data_index in data_indices:
            if self.view_mode == "grouped" and data_index < len(self.grouped_data):
                data = self.grouped_data[data_index]
                for trip in data['trips']:
                    trip['business_name'] = name
                data['business_name'] = name
                data['status'] = 'Has Name'
                self._save_business_mapping(data['address'], name)

            elif self.view_mode == "individual" and data_index < len(self.trips_data):
                trip = self.trips_data[data_index]
                trip['business_name'] = name
                self._save_business_mapping(trip.get('end_address', ''), name)

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

    def _apply_category_to_selected(self, data_indices: List[int], category: str):
        """Apply category to selected rows (data_indices are original data indices)"""
        for data_index in data_indices:
            if self.view_mode == "grouped" and data_index < len(self.grouped_data):
                data = self.grouped_data[data_index]
                for trip in data['trips']:
                    trip['computed_category'] = category
                data['primary_category'] = category

            elif self.view_mode == "individual" and data_index < len(self.trips_data):
                trip = self.trips_data[data_index]
                trip['computed_category'] = category

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

    def _open_in_google_maps(self, row: int):
        """Open in external Google Maps"""
        import webbrowser
        addr = ""
        if self.view_mode == "grouped" and row < len(self.grouped_data):
            addr = self.grouped_data[row]['address']
        elif self.view_mode == "individual" and row < len(self.trips_data):
            addr = self.trips_data[row].get('end_address', '')

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
        
        if trip_date:
            # Find all trips on the same date
            target_date = trip_date.date() if hasattr(trip_date, 'date') else trip_date
            day_trips = [
                t for t in self.trips_data 
                if hasattr(t.get('started'), 'date') and t['started'].date() == target_date
            ]
            
            if day_trips:
                self.show_daily_journey.emit(day_trips)

    def _save_business_mapping(self, address: str, name: str):
        """Save business mapping to file"""
        if not address or not name:
            return

        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        mappings = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except:
                pass

        mappings[address] = name

        try:
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, indent=2, ensure_ascii=False)
        except:
            pass

    def _get_existing_business_names(self) -> set:
        """Get existing business names"""
        names = set()
        skip = {'Home', 'Office', '[PERSONAL]', 'Unknown', '', 'NO_BUSINESS_FOUND'}

        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    for name in json.load(f).values():
                        if name and name not in skip:
                            names.add(name)
            except:
                pass

        cache_file = os.path.join(get_app_dir(), 'address_cache.json')
        if os.path.exists(cache_file):
            try:
                with open(cache_file, 'r', encoding='utf-8') as f:
                    for entry in json.load(f).values():
                        name = entry.get('name', '') if isinstance(entry, dict) else str(entry)
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

        # Add existing business names from mapping
        existing_names = self._get_existing_business_names()
        name_actions = {}
        if existing_names:
            recent_menu = assign_menu.addMenu("Existing Names")
            for name in sorted(existing_names)[:15]:
                action = recent_menu.addAction(name)
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

    def _save_business_mapping(self, address: str, name: str):
        """Save a business name mapping to the mapping file"""
        if not address or not name:
            return

        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        mappings = {}

        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except:
                pass

        mappings[address] = name

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
        self._setup_ui()
        self._setup_menu()

    def _setup_ui(self):
        self.setWindowTitle("Mileage Analyzer")
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

        # Set initial sizes (left=40%, right=60%)
        splitter.setSizes([480, 720])

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

        # Business lookup checkbox - default to enabled for automatic lookup
        self.lookup_checkbox = QCheckBox("Enable Business Lookup")
        self.lookup_checkbox.setChecked(True)  # Enable by default
        self.lookup_checkbox.setToolTip("Use online services to identify businesses (uses cache for speed)")
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

        mapping_action = QAction("Edit Business &Mappings...", self)
        mapping_action.triggered.connect(self._edit_mappings)
        tools_menu.addAction(mapping_action)

        cache_action = QAction("Edit Address &Cache...", self)
        cache_action.triggered.connect(self._edit_address_cache)
        tools_menu.addAction(cache_action)

        tools_menu.addSeparator()

        clear_cache_action = QAction("Clear Address Cache", self)
        clear_cache_action.triggered.connect(self._clear_address_cache)
        tools_menu.addAction(clear_cache_action)

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
            self.status_bar.showMessage(f"Loading: {os.path.basename(file_path)}...")
            # First load trips quickly without lookup, then auto-start lookup
            self._run_analysis(enable_lookup=False, auto_continue=True)

    def _run_analysis(self, enable_lookup=None, auto_continue=False):
        """Run the mileage analysis

        Args:
            enable_lookup: Override for business lookup. If None, uses checkbox state.
            auto_continue: If True, automatically run business lookup after initial load.
        """
        if not self.current_file:
            QMessageBox.warning(self, "No File", "Please open a trip log file first.")
            return

        # Store whether to auto-continue with lookup after this pass
        self._auto_continue_lookup = auto_continue

        # Get date range
        start_date = self.start_date.date().toString("yyyy-MM-dd")
        end_date = self.end_date.date().toString("yyyy-MM-dd")

        # Determine lookup setting
        if enable_lookup is None:
            enable_lookup = self.lookup_checkbox.isChecked()

        # Show progress
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.progress_bar.show()

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
            dest_count = len(self.unified_view.grouped_data)
            self.status_bar.showMessage(f"Analysis complete. {len(trips)} trips to {dest_count} destinations. {needs_name_count} need names.")

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

    def _save_business_mapping(self, address: str, business_name: str):
        """Save address to business name mapping"""
        mapping_file = os.path.join(get_app_dir(), 'business_mapping.json')
        mappings = {}
        if os.path.exists(mapping_file):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    mappings = json.load(f)
            except:
                pass
        mappings[address] = business_name
        try:
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump(mappings, f, indent=2, ensure_ascii=False)
        except Exception as e:
            self.status_bar.showMessage(f"Failed to save mapping: {e}", 5000)

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

    def _edit_address_cache(self):
        """Open address cache editor"""
        cache_file = os.path.join(get_app_dir(), 'address_cache.json')
        if not os.path.exists(cache_file):
            QMessageBox.information(
                self, "No Cache",
                "Address cache is empty.\n"
                "Run an analysis with business lookup to populate the cache."
            )
            return

        self.cache_editor = JsonEditorDialog(
            cache_file,
            "Address Cache Editor",
            self
        )
        self.cache_editor.show()

    def _clear_address_cache(self):
        """Clear the address cache"""
        cache_file = os.path.join(get_app_dir(), 'address_cache.json')
        if not os.path.exists(cache_file):
            QMessageBox.information(self, "No Cache", "Address cache is already empty.")
            return

        reply = QMessageBox.question(
            self, "Confirm Clear Cache",
            "This will delete all cached address lookups.\n"
            "The next analysis will need to re-lookup all business names.\n\n"
            "Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                os.remove(cache_file)
                QMessageBox.information(self, "Cache Cleared", "Address cache has been cleared.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to clear cache:\n{e}")

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
