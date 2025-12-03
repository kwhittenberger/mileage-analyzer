#!/usr/bin/env python3
"""
Analyze addresses where no business was found
Helps identify patterns and potential manual mappings
"""

import json
import os
from collections import defaultdict

def normalize_address(address):
    """Normalize address for comparison"""
    if not address:
        return ""
    return address.strip().lower()

def analyze_unresolved():
    """Analyze addresses with NO_BUSINESS_FOUND"""
    cache_file = "address_cache.json"

    if not os.path.exists(cache_file):
        print("No address_cache.json found!")
        return

    # Load cache
    with open(cache_file, 'r', encoding='utf-8') as f:
        cache = json.load(f)

    # Find all NO_BUSINESS_FOUND entries
    unresolved = [addr for addr, value in cache.items() if value == "NO_BUSINESS_FOUND"]
    resolved = {addr: value for addr, value in cache.items() if value != "NO_BUSINESS_FOUND"}

    print("=" * 80)
    print("UNRESOLVED ADDRESS ANALYSIS")
    print("=" * 80)
    print()
    print(f"Total addresses in cache: {len(cache)}")
    print(f"  - Resolved (with business names): {len(resolved)}")
    print(f"  - Unresolved (NO_BUSINESS_FOUND): {len(unresolved)}")
    print()

    if not unresolved:
        print("All addresses have been resolved!")
        return

    # Group by city/area
    print("-" * 80)
    print("UNRESOLVED ADDRESSES BY AREA")
    print("-" * 80)
    print()

    by_area = defaultdict(list)
    for addr in sorted(unresolved):
        norm = normalize_address(addr)
        if "kenmore" in norm:
            by_area["Kenmore"].append(addr)
        elif "bothell" in norm:
            by_area["Bothell"].append(addr)
        elif "coupeville" in norm or "oak harbor" in norm or "clinton" in norm or "whidbey" in norm or "langley" in norm or "freeland" in norm:
            by_area["Whidbey Island"].append(addr)
        elif "seattle" in norm:
            by_area["Seattle"].append(addr)
        elif "everett" in norm:
            by_area["Everett"].append(addr)
        elif "mukilteo" in norm:
            by_area["Mukilteo"].append(addr)
        elif "portland" in norm:
            by_area["Portland"].append(addr)
        else:
            by_area["Other"].append(addr)

    for area in sorted(by_area.keys()):
        addrs = by_area[area]
        print(f"{area} ({len(addrs)} addresses):")
        for addr in sorted(addrs)[:10]:  # Show first 10
            print(f"  - {addr}")
        if len(addrs) > 10:
            print(f"  ... and {len(addrs) - 10} more")
        print()

    # Look for patterns - residential streets
    print("-" * 80)
    print("LIKELY RESIDENTIAL ADDRESSES (Street names with numbers)")
    print("-" * 80)
    print()

    residential_patterns = ["ln ", "ct ", "pl ", "cir ", "way ", "dr ", "ave ", "st "]
    likely_residential = []

    for addr in sorted(unresolved):
        norm = normalize_address(addr)
        # Has a street number at the start and a residential street type
        if any(c.isdigit() for c in addr[:6]) and any(pattern in norm for pattern in residential_patterns):
            # Check if it looks like a residential address (not a business)
            if not any(keyword in norm for keyword in ["blvd", "hwy", "highway", "route", "state route"]):
                likely_residential.append(addr)

    print(f"Found {len(likely_residential)} likely residential addresses:")
    for addr in sorted(likely_residential)[:15]:
        print(f"  - {addr}")
    if len(likely_residential) > 15:
        print(f"  ... and {len(likely_residential) - 15} more")
    print()
    print("These are probably homes and may not need business names.")
    print()

    # Look for highways/routes
    print("-" * 80)
    print("HIGHWAYS & ROUTES (May be in transit)")
    print("-" * 80)
    print()

    highways = []
    for addr in sorted(unresolved):
        norm = normalize_address(addr)
        if any(keyword in norm for keyword in ["i-90", "wa-", "hwy", "highway", "route", "state route", "fry"]):
            highways.append(addr)

    if highways:
        print(f"Found {len(highways)} highway/route addresses:")
        for addr in sorted(highways):
            print(f"  - {addr}")
        print()
        print("These may be rest stops or simply points on routes.")
        print()

    # Suggest potential manual mappings
    print("-" * 80)
    print("RECOMMENDATIONS")
    print("-" * 80)
    print()
    print("1. RESIDENTIAL ADDRESSES:")
    print("   These don't need business names. The system will handle them correctly.")
    print()
    print("2. FREQUENT DESTINATIONS:")
    print("   If any of these addresses appear frequently in your trips,")
    print("   consider adding them to business_mapping.json with custom names.")
    print()
    print("3. CHECK AGAINST RESOLVED:")
    print("   Some addresses may be very close to resolved ones.")
    print("   Use fuzzy matching in business_mapping.json for similar addresses.")
    print()

if __name__ == '__main__':
    analyze_unresolved()
