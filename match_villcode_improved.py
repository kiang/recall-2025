#!/usr/bin/env python3
import json
import os
import glob
import re

def load_villcode_mapping():
    """Load VILLCODE mapping from taiwan basecode file"""
    with open('/home/kiang/public_html/taiwan_basecode/cunli/geo/20240807.json', 'r', encoding='utf-8') as f:
        basecode_data = json.load(f)
    
    # Create mapping dictionary with multiple key formats
    villcode_map = {}
    for feature in basecode_data['features']:
        props = feature['properties']
        if props['VILLNAME'] is not None:
            county = props['COUNTYNAME']
            town = props['TOWNNAME']
            village = props['VILLNAME']
            
            # Create multiple key variations for better matching
            keys = []
            
            # Original key
            keys.append(f"{county}_{town}_{village}")
            
            # Normalized variations (臺 <-> 台)
            county_norm = county.replace('臺', '台')
            town_norm = town.replace('臺', '台')
            keys.append(f"{county_norm}_{town_norm}_{village}")
            keys.append(f"{county_norm}_{town}_{village}")
            keys.append(f"{county}_{town_norm}_{village}")
            
            data = {
                'VILLCODE': props['VILLCODE'],
                'COUNTYCODE': props['COUNTYCODE'],
                'TOWNCODE': props['TOWNCODE'],
                'COUNTYNAME': county,
                'TOWNNAME': town,
                'VILLNAME': village
            }
            
            # Store under all key variations
            for key in keys:
                villcode_map[key] = data
    
    return villcode_map

def extract_county_from_recall_case(recall_case):
    """Extract county/city name from recall case string"""
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
        
        # Skip if not a dict
        if not isinstance(data, dict):
            continue
            
        # Extract district and village from cunli key
        cunli_parts = data['cunli'].split('_')
        if len(cunli_parts) >= 2:
            district = cunli_parts[0]
            village = cunli_parts[1]
            
            # Extract county from recall case
            county = None
            if data['records']:
                first_record = data['records'][0]
                recall_case = first_record.get('recall_case', '')
                county = extract_county_from_recall_case(recall_case)
            
            if county:
                # Try to find matching VILLCODE
                key = f"{county}_{district}_{village}"
                
                if key in villcode_map:
                    # Update with VILLCODE info (don't overwrite existing keys)
                    villcode_info = villcode_map[key]
                    for k, v in villcode_info.items():
                        if k not in data:
                            data[k] = v
                    
                    # Save updated file
                    with open(json_file, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                    updated_count += 1
                    print(f"Updated: {data['cunli']}")
                else:
                    not_found.append({
                        'cunli': data['cunli'],
                        'county': county,
                        'district': district,
                        'village': village,
                        'key_tried': key
                    })
    
    return updated_count, not_found

def main():
    print("Loading VILLCODE mapping...")
    villcode_map = load_villcode_mapping()
    print(f"Loaded {len(villcode_map)} village code mappings")
    
    print("\nUpdating cunli JSON files...")
    updated_count, not_found = update_cunli_files(villcode_map)
    
    print(f"\nUpdated {updated_count} files with VILLCODE")
    print(f"Could not find VILLCODE for {len(not_found)} villages")
    
    if not_found:
        print("\nVillages without VILLCODE (first 10):")
        for village in not_found[:10]:
            print(f"  - {village['cunli']} (tried: {village['key_tried']})")
        
        # Save not found list for review
        with open('cunli_json/villages_still_missing.json', 'w', encoding='utf-8') as f:
            json.dump(not_found, f, ensure_ascii=False, indent=2)

if __name__ == '__main__':
    main()