#!/usr/bin/env python3
import pandas as pd
import json
import os
from collections import defaultdict
import glob

def extract_excel_data(file_path):
    """Extract data from a single Excel file"""
    df = pd.read_excel(file_path, header=None)
    
    # Find the header row (contains '行政區別')
    header_row = None
    for i in range(10):
        if df.iloc[i, 0] == '行政區別':
            header_row = i
            break
    
    if header_row is None:
        print(f"Warning: Could not find header in {file_path}")
        return []
    
    # Extract recall target info from filename
    filename = os.path.basename(file_path)
    recall_info = filename.replace('表5-', '').replace('.xlsx', '')
    
    results = []
    current_district = None
    
    # Process data rows
    for idx in range(header_row + 5, len(df)):  # Skip empty rows after header
        row = df.iloc[idx]
        
        # Check if it's a district row
        if pd.notna(row[0]) and pd.isna(row[1]):
            current_district = str(row[0]).strip()
            continue
        
        # Check if it's a village data row
        if pd.notna(row[1]) and pd.notna(row[2]):
            try:
                record = {
                    'recall_case': recall_info,
                    'district': current_district,
                    'village': str(row[1]).strip(),
                    'polling_station': str(row[2]).strip(),
                    'agree_votes': int(row[3]) if pd.notna(row[3]) else 0,
                    'disagree_votes': int(row[4]) if pd.notna(row[4]) else 0,
                    'valid_votes': int(row[5]) if pd.notna(row[5]) else 0,
                    'invalid_votes': int(row[6]) if pd.notna(row[6]) else 0,
                    'total_voters': int(row[7]) if pd.notna(row[7]) else 0,
                    'ballots_not_cast': int(row[8]) if pd.notna(row[8]) else 0,
                    'ballots_issued': int(row[9]) if pd.notna(row[9]) else 0,
                    'unused_ballots': int(row[10]) if pd.notna(row[10]) else 0,
                    'eligible_voters': int(row[11]) if pd.notna(row[11]) else 0,
                    'turnout_rate': float(str(row[12]).replace('%', '')) if pd.notna(row[12]) else 0.0
                }
                results.append(record)
            except (ValueError, TypeError) as e:
                # Skip rows that don't have valid data
                continue
    
    return results

def load_villcode_mapping():
    """Load VILLCODE mapping from taiwan basecode file"""
    with open('/home/kiang/public_html/taiwan_basecode/cunli/geo/20250620.json', 'r', encoding='utf-8') as f:
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
    import re
    
    # Pattern to match county/city names in parentheses
    match = re.search(r'\(([^)]+)(市|縣)(?:第\d+)?選舉區\)', recall_case)
    if match:
        return match.group(1) + match.group(2)
    
    # Special case for mayor recall
    if '新竹市第11屆市長' in recall_case:
        return '新竹市'
    
    return None

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

def merge_by_cunli(all_data):
    """Merge data by cunli (village)"""
    cunli_data = defaultdict(list)
    
    for record in all_data:
        if record['district'] and record['village']:
            # Create a key combining district and village
            key = f"{record['district']}_{record['village']}"
            cunli_data[key].append(record)
    
    return dict(cunli_data)

def main():
    # Get all Excel files
    excel_files = glob.glob('raw/*.xlsx')
    print(f"Found {len(excel_files)} Excel files")
    
    # Extract data from all files
    all_data = []
    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")
        data = extract_excel_data(file_path)
        all_data.extend(data)
        print(f"  Extracted {len(data)} records")
    
    print(f"\nTotal records extracted: {len(all_data)}")
    
    # Merge by cunli
    cunli_merged = merge_by_cunli(all_data)
    print(f"Merged into {len(cunli_merged)} cunli groups")
    
    # Create output directory
    os.makedirs('cunli_json', exist_ok=True)
    
    # Load VILLCODE mapping
    villcode_map = load_villcode_mapping()
    
    # Save individual cunli JSON files
    villcode_files = []
    missing_villcode = []
    for cunli_key, records in cunli_merged.items():
        # Calculate sum fields
        sum_fields = {
            'agree_votes': sum(r['agree_votes'] for r in records),
            'disagree_votes': sum(r['disagree_votes'] for r in records),
            'valid_votes': sum(r['valid_votes'] for r in records),
            'invalid_votes': sum(r['invalid_votes'] for r in records),
            'total_voters': sum(r['total_voters'] for r in records),
            'ballots_not_cast': sum(r['ballots_not_cast'] for r in records),
            'ballots_issued': sum(r['ballots_issued'] for r in records),
            'unused_ballots': sum(r['unused_ballots'] for r in records),
            'eligible_voters': sum(r['eligible_voters'] for r in records)
        }
        
        # Calculate average turnout rate
        if sum_fields['eligible_voters'] > 0:
            sum_fields['average_turnout_rate'] = round(
                (sum_fields['total_voters'] / sum_fields['eligible_voters']) * 100, 2
            )
        else:
            sum_fields['average_turnout_rate'] = 0.0
        
        # Prepare data
        data = {
            'cunli': cunli_key,
            'total_records': len(records),
            'sum_fields': sum_fields,
            'records': records
        }
        
        # Try to get VILLCODE for filename
        district = cunli_key.split('_')[0]
        village = cunli_key.split('_')[1]
        
        # Extract county from recall case
        county = None
        if records:
            recall_case = records[0].get('recall_case', '')
            county = extract_county_from_recall_case(recall_case)
        
        villcode = None
        filename = cunli_key  # fallback
        
        if county:
            # Normalize county name
            county = normalize_district_name(county)
            # Try to find VILLCODE
            key = f"{county}_{district}_{village}"
            if key in villcode_map:
                villcode_info = villcode_map[key]
                # Add VILLCODE info to data
                for k, v in villcode_info.items():
                    data[k] = v
                villcode = villcode_info['VILLCODE']
                filename = villcode
                villcode_files.append(villcode)
            else:
                # Track villages without VILLCODE
                missing_villcode.append({
                    'cunli_key': cunli_key,
                    'district': district,
                    'village': village,
                    'county': county,
                    'villcode': '',  # Empty for manual filling
                    'note': 'Please fill VILLCODE manually'
                })
        
        output_file = f"cunli_json/{filename}.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    
    # Save summary file
    summary = {
        'total_files_processed': len(excel_files),
        'total_records': len(all_data),
        'total_cunli': len(cunli_merged),
        'villcode_files': len(villcode_files),
        'file_naming': 'VILLCODE (where available), cunli name (fallback)',
        'villcode_list': sorted(villcode_files),
        'cunli_list': sorted(list(cunli_merged.keys())),
        'note': 'Files are named using official Taiwan VILLCODE when available for precise geographic identification'
    }
    
    with open('cunli_json/summary.json', 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    
    # Save missing VILLCODE mapping file for manual completion
    if missing_villcode:
        with open('missing_villcode_mapping.json', 'w', encoding='utf-8') as f:
            json.dump(missing_villcode, f, ensure_ascii=False, indent=2)
        print(f"Created missing_villcode_mapping.json with {len(missing_villcode)} entries for manual completion")
    
    print("\nExtraction and merging completed!")
    print(f"Created {len(villcode_files)} files with VILLCODE names")
    print(f"Created {len(cunli_merged) - len(villcode_files)} files with cunli names (no VILLCODE found)")
    print(f"Output files saved in 'cunli_json' directory")

if __name__ == '__main__':
    main()