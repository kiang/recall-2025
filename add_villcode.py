#!/usr/bin/env python3
import json
import os
import glob

def load_villcode_mapping():
    """Load VILLCODE mapping from taiwan basecode file"""
    with open('/home/kiang/public_html/taiwan_basecode/cunli/geo/20240807.json', 'r', encoding='utf-8') as f:
        basecode_data = json.load(f)
    
    # Create mapping dictionary
    villcode_map = {}
    for feature in basecode_data['features']:
        props = feature['properties']
        if props['VILLNAME'] is not None:
            # Create key with county/city + district/town + village
            key = f"{props['COUNTYNAME']}_{props['TOWNNAME']}_{props['VILLNAME']}"
            villcode_map[key] = {
                'VILLCODE': props['VILLCODE'],
                'COUNTYCODE': props['COUNTYCODE'],
                'TOWNCODE': props['TOWNCODE'],
                'COUNTYNAME': props['COUNTYNAME'],
                'TOWNNAME': props['TOWNNAME'],
                'VILLNAME': props['VILLNAME']
            }
    
    return villcode_map

def normalize_district_name(district_name):
    """Normalize district names to match basecode format"""
    # Handle special cases
    replacements = {
        '臺北市': '台北市',
        '臺中市': '台中市',
        '臺東縣': '台東縣',
        '臺南市': '台南市'
    }
    
    for old, new in replacements.items():
        district_name = district_name.replace(old, new)
    
    return district_name

def extract_county_from_recall_case(recall_case):
    """Extract county/city name from recall case string"""
    import re
    
    # Pattern to match county/city names in parentheses
    match = re.search(r'\(([^)]+)(市|縣)(?:第\d+)?選舉區\)', recall_case)
    if match:
        return match.group(1) + match.group(2)
    
    # Special case for mayor recall
    if '新竹市第11屆市長' in recall_case:
        return '新竹市'
    
    return None

def update_cunli_files(villcode_map):
    """Update cunli JSON files with VILLCODE"""
    json_files = glob.glob('cunli_json/*.json')
    updated_count = 0
    not_found = []
    
    for json_file in json_files:
        filename = os.path.basename(json_file)
        if filename in ['summary.json', 'villages_without_villcode.json']:
            continue
            
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Skip if not a dict or already has VILLCODE
        if not isinstance(data, dict) or 'VILLCODE' in data:
            continue
            
        # Extract district and village from cunli key
        cunli_parts = data['cunli'].split('_')
        if len(cunli_parts) >= 2:
            district = cunli_parts[0]
            village = cunli_parts[1]
            
            # Try to find matching VILLCODE
            found = False
            
            # Extract county from recall case
            if data['records']:
                first_record = data['records'][0]
                recall_case = first_record.get('recall_case', '')
                county = extract_county_from_recall_case(recall_case)
                
                if county:
                    # Normalize county name
                    county = normalize_district_name(county)
                    
                    # Create key
                    key = f"{county}_{district}_{village}"
                    
                    if key in villcode_map:
                        data.update(villcode_map[key])
                        found = True
            
            if found:
                # Save updated file
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                updated_count += 1
            else:
                not_found.append({
                    'cunli': data['cunli'],
                    'county': county if 'county' in locals() else 'Unknown',
                    'district': district,
                    'village': village
                })
    
    return updated_count, not_found

def main():
    print("Loading VILLCODE mapping...")
    villcode_map = load_villcode_mapping()
    print(f"Loaded {len(villcode_map)} village codes")
    
    print("\nUpdating cunli JSON files...")
    updated_count, not_found = update_cunli_files(villcode_map)
    
    print(f"\nUpdated {updated_count} files with VILLCODE")
    print(f"Could not find VILLCODE for {len(not_found)} villages")
    
    if not_found:
        print("\nVillages without VILLCODE (first 20):")
        for village in not_found[:20]:
            print(f"  - {village}")
        
        # Save not found list for review
        with open('cunli_json/villages_without_villcode.json', 'w', encoding='utf-8') as f:
            json.dump(not_found, f, ensure_ascii=False, indent=2)

if __name__ == '__main__':
    main()