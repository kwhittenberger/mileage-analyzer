# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Mileage Analyzer GUI
D'Ewart Representatives, L.L.C.

This builds a standalone Windows executable with embedded PyQt6 and WebEngine.
"""

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect all PyQt6 and WebEngine data files
datas = []
datas += collect_data_files('PyQt6')
datas += collect_data_files('PyQt6.QtWebEngineCore')

# Hidden imports for PyQt6 WebEngine
hiddenimports = [
    'PyQt6.QtWebEngineWidgets',
    'PyQt6.QtWebEngineCore',
    'PyQt6.QtWebChannel',
    'PyQt6.QtPositioning',
    'PyQt6.QtPrintSupport',
    'PyQt6.sip',
    'analyze_mileage',  # Our analysis module
    'openpyxl',
    'geopy',
    'geopy.geocoders',
    'googlemaps',
]
hiddenimports += collect_submodules('PyQt6')

a = Analysis(
    ['mileage_gui.py'],
    pathex=['.'],  # Add current directory to path
    binaries=[],
    datas=datas + [
        ('config.json', '.'),
        ('business_mapping.json', '.'),
        ('address_cache.json', '.'),
    ],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MileageAnalyzerGUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon path here if you have one: icon='app_icon.ico'
)
