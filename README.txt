================================================================================
MILEAGE ANALYSIS TOOL - USER GUIDE
D'Ewart Representatives, L.L.C.
================================================================================

OVERVIEW
--------
This tool analyzes your vehicle trip data and automatically categorizes trips
as business, personal, or commute based on the requirements you specified.

The analysis includes:
- Weekly mileage summaries
- Business vs Personal mileage breakdown
- Detailed trip log with business location names
- Automatic identification of commutes, gas stations, and business trips


REQUIREMENTS
------------
- Python 3.x (download from https://www.python.org/downloads/ if needed)
- Your trip log CSV file (exported from your vehicle tracking app)

OPTIONAL (for enhanced business lookup):
- geopy library: pip install geopy (free, basic business lookup)
- googlemaps library: pip install googlemaps (enhanced business lookup)
- Google Places API key (free tier, see GOOGLE_API_SETUP.txt)


INSTALLATION
------------
1. Ensure Python is installed on your computer
2. Place all files in the same folder:
   - analyze_mileage.py (the analysis script)
   - run_mileage_analysis.bat (easy run script for Windows)
   - volvo-trips-log.csv (your trip data)


HOW TO RUN
----------

OPTION 1 - Easy Method (Windows):
   Simply double-click: run_mileage_analysis.bat

OPTION 2 - Command Line Method:
   Open Command Prompt in this folder and run:
   python analyze_mileage.py volvo-trips-log.csv

OPTION 3 - With Business Lookup (slower, but more accurate):
   python analyze_mileage.py volvo-trips-log.csv --lookup

   Note: Business lookup uses OpenStreetMap to identify businesses at addresses.
   Requires internet connection and takes longer on first run.
   Results are cached for faster subsequent runs.

OPTION 4 - Save Text Report to Custom File:
   python analyze_mileage.py volvo-trips-log.csv > my_report.txt


OUTPUT
------
The analysis generates both text reports and CSV files:

TEXT REPORT (displayed on screen):
1. WEEKLY MILEAGE SUMMARY
   - Shows commute, business, and personal miles for each week
   - Weeks run Monday to Sunday

2. BUSINESS VS PERSONAL SUMMARY
   - Total miles driven
   - Business miles percentage
   - Personal miles breakdown (commute + other personal)

3. DETAILED TRIP LOG
   - Every trip listed chronologically
   - Shows date, time, category, distance, and locations
   - Business trips include the business name/type

CSV FILES (Excel-compatible):
1. weekly_summary.csv
   - Weekly breakdown of mileage by category
   - Perfect for creating charts and graphs

2. detailed_trips.csv
   - Complete trip log with all details
   - Includes: date, time, day of week, category, distance, addresses,
     business name, odometer readings, and week number
   - Easily filter and sort in Excel

3. summary.csv
   - Overall totals and percentages
   - Quick overview of business vs personal miles

All CSV files can be opened in Excel, Google Sheets, or any spreadsheet program.


CATEGORIZATION RULES
--------------------
The tool automatically applies these rules:

COMMUTE:
- Trips between home (15815 61st Ln NE, Kenmore) and
  work (9227 NE 180th St, Bothell)

BUSINESS:
- All gas station trips
- Trips over 8 miles during weekdays
- Trips to/from Whidbey Island (when not weekend personal travel)
- Short trips in Bothell/Kenmore area during weekdays

PERSONAL:
- Weekend trips (Friday evening to Monday morning)
- Short local trips on weekends
- Any trip not categorized as commute or business


UPDATING YOUR DATA
------------------
To analyze new trip data:
1. Export your updated CSV file from your tracking app
2. Replace volvo-trips-log.csv with the new file
3. Run the analysis again


BUSINESS LOOKUP FEATURE
-----------------------
The tool has a multi-tier system for identifying businesses:

TIER 1 - GOOGLE PLACES API (BEST, optional):
- Most accurate business names (offices, warehouses, business parks)
- Fast lookups (~300ms per address)
- Requires free Google API key (see GOOGLE_API_SETUP.txt)
- FREE for typical use ($200/month credit = ~11,000 lookups)
- Setup time: 10-15 minutes (one-time)

TIER 2 - OPENSTREETMAP (GOOD, free):
- Automatically identifies gas stations, restaurants, stores
- No API key needed
- Slower (~1 second per address)
- Falls back to this if Google API not configured

TIER 3 - MANUAL MAPPING (ALWAYS AVAILABLE):
- Edit business_mapping.json to add custom entries:
  "123 Main St, Bothell": "ABC Company Customer Site"
- Manual mappings have highest priority
- Perfect for frequent customer locations

HOW IT WORKS:
1. First checks business_mapping.json (instant, manual entries)
2. Then checks address_cache.json (instant, previous lookups)
3. If --lookup enabled and Google API key configured: Uses Google Places
4. If --lookup enabled but no Google key: Uses OpenStreetMap
5. Falls back to street address if nothing found

Without --lookup flag, uses pattern matching only (fast, basic names).

RECOMMENDATION:
- For best results: Set up Google Places API (see GOOGLE_API_SETUP.txt)
- For good results: Use --lookup without Google (uses OpenStreetMap)
- For quick results: Use manual mapping only (no --lookup needed)


CUSTOMIZATION
-------------
To modify the analysis rules, edit these settings in analyze_mileage.py:

Lines 26-28: Home and work addresses
Line 31: Business distance threshold (currently 8.0 miles)
Lines 34-38: Business location keywords


TROUBLESHOOTING
---------------
Problem: "Python is not installed or not in PATH"
Solution: Install Python from https://www.python.org/downloads/
          Make sure to check "Add Python to PATH" during installation

Problem: "volvo-trips-log.csv not found"
Solution: Make sure your CSV file is in the same folder as the script
          and named exactly: volvo-trips-log.csv

Problem: "No trips found in CSV file"
Solution: Verify your CSV file is not empty and is in the correct format

Problem: "geopy not installed" warning
Solution: This is just a warning. The script will work without geopy,
          but business lookup will be disabled.
          To enable business lookup: pip install geopy

Problem: Business lookup is slow
Solution: This is normal on first run with OpenStreetMap (~1 sec/address).
          For faster lookups, set up Google Places API (see GOOGLE_API_SETUP.txt).
          All lookups are cached for instant subsequent runs.

Problem: "Using: OpenStreetMap" but I want Google Places
Solution: Add your Google API key to config.json file.
          See GOOGLE_API_SETUP.txt for complete instructions.


SUPPORT
-------
For questions or issues, contact:
Doug D'Ewart
D'Ewart Representatives, L.L.C.
Email: doug@dewart.com
Office: 425.485.6545
Cell: 206.510.9692


VERSION HISTORY
---------------
Version 2.1 - November 2025
- Added Google Places API integration for superior business lookup
- Multi-tier lookup system (Google > OpenStreetMap > Manual)
- Automatic fallback if API key not configured
- Displays which lookup method is active
- Comprehensive setup guide for Google API

Version 2.0 - November 2025
- Added CSV export (weekly_summary.csv, detailed_trips.csv, summary.csv)
- Added online business lookup feature using OpenStreetMap
- Address caching for faster subsequent runs
- Excel-compatible output files
- Enhanced batch file with business lookup option

Version 1.0 - November 2025 - Initial release
- Automatic trip categorization
- Weekly summaries
- Business location identification
- Business vs personal calculations
- Text report generation


================================================================================
