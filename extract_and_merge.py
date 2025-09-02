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

def load_manual_villcode_mapping():
    """Load manually filled VILLCODE mappings if file exists"""
    manual_map = {}
    if os.path.exists('missing_villcode_mapping.json'):
        try:
            with open('missing_villcode_mapping.json', 'r', encoding='utf-8') as f:
                manual_mappings = json.load(f)
            
            for mapping in manual_mappings:
                if mapping.get('villcode'):  # Only add if villcode is filled
                    key = mapping['cunli_key']
                    manual_map[key] = {
                        'VILLCODE': mapping['villcode'],
                        'COUNTYCODE': mapping['villcode'][:3] if len(mapping['villcode']) >= 3 else '',
                        'TOWNCODE': mapping['villcode'][:6] if len(mapping['villcode']) >= 6 else '',
                        'COUNTYNAME': mapping['county'],
                        'TOWNNAME': mapping['district'],
                        'VILLNAME': mapping['village']
                    }
            
            print(f"Loaded {len(manual_map)} manual VILLCODE mappings")
        except Exception as e:
            print(f"Error loading manual mappings: {e}")
    
    return manual_map

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
    os.makedirs('docs/cunli_json', exist_ok=True)
    
    # Load VILLCODE mapping
    villcode_map = load_villcode_mapping()
    
    # Load manual VILLCODE mappings if available
    manual_villcode_map = load_manual_villcode_mapping()
    
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
        
        # First check manual mappings
        if cunli_key in manual_villcode_map:
            villcode_info = manual_villcode_map[cunli_key]
            # Add VILLCODE info to data
            for k, v in villcode_info.items():
                data[k] = v
            villcode = villcode_info['VILLCODE']
            # Handle case where VILLCODE might be a list (combined villages)
            if isinstance(villcode, list):
                filename = villcode[0]  # Use first code for filename
                villcode_files.extend(villcode)  # Add all codes to list
            else:
                filename = villcode
                villcode_files.append(villcode)
        elif county:
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
        
        # Special case: Fix data error for village 65000040036 (新北市永和區光復里)
        # According to UDN news report, the agree and disagree votes were swapped
        # Reference: https://udn.com/news/story/124323/8903780
        if villcode == '65000040036' or (data.get('VILLCODE') == '65000040036'):
            print(f"Applying data correction for 65000040036 (新北市永和區光復里): exchanging agree/disagree votes")
            # Exchange agree and disagree votes in sum_fields
            original_agree = data['sum_fields']['agree_votes']
            data['sum_fields']['agree_votes'] = data['sum_fields']['disagree_votes']
            data['sum_fields']['disagree_votes'] = original_agree
            
            # Also exchange in individual records
            for record in data['records']:
                original_record_agree = record['agree_votes']
                record['agree_votes'] = record['disagree_votes']
                record['disagree_votes'] = original_record_agree
        
        output_file = f"docs/cunli_json/{filename}.json"
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
        'note': 'Files are named using official Taiwan VILLCODE for precise geographic identification',
        'manual_mappings_applied': len(manual_villcode_map)
    }
    
    with open('docs/cunli_json/summary.json', 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    
    # Generate cunli summary for faster map loading
    cunli_summary = {}
    
    for cunli_key, records in cunli_merged.items():
        # Check if this needs data correction first (before calculating sums)
        # Special case: village 65000040036 (新北市永和區光復里) has swapped agree/disagree votes
        # Reference: https://udn.com/news/story/124323/8903780
        needs_correction = (cunli_key == '永和區_光復里')
        
        # Apply correction to records if needed
        corrected_records = records
        if needs_correction:
            corrected_records = []
            for r in records:
                corrected_record = r.copy()
                # Exchange agree and disagree votes
                original_agree = corrected_record['agree_votes']
                corrected_record['agree_votes'] = corrected_record['disagree_votes']
                corrected_record['disagree_votes'] = original_agree
                corrected_records.append(corrected_record)
        
        # Calculate sum fields from (possibly corrected) records
        sum_fields = {
            'agree_votes': sum(r['agree_votes'] for r in corrected_records),
            'disagree_votes': sum(r['disagree_votes'] for r in corrected_records),
            'valid_votes': sum(r['valid_votes'] for r in corrected_records),
            'invalid_votes': sum(r['invalid_votes'] for r in corrected_records),
            'total_voters': sum(r['total_voters'] for r in corrected_records),
            'eligible_voters': sum(r['eligible_voters'] for r in corrected_records)
        }
        
        # Calculate average turnout rate
        if sum_fields['eligible_voters'] > 0:
            sum_fields['average_turnout_rate'] = round(
                (sum_fields['total_voters'] / sum_fields['eligible_voters']) * 100, 2
            )
        else:
            sum_fields['average_turnout_rate'] = 0.0
        
        # Get VILLCODE
        district = cunli_key.split('_')[0]
        village = cunli_key.split('_')[1]
        
        # Extract county from recall case
        county = None
        if records:
            recall_case = records[0].get('recall_case', '')
            county = extract_county_from_recall_case(recall_case)
        
        villcode = None
        
        # First check manual mappings
        if cunli_key in manual_villcode_map:
            villcode = manual_villcode_map[cunli_key]['VILLCODE']
            # Handle list of VILLCODEs for combined villages
            if isinstance(villcode, list):
                villcode = villcode[0]  # Use first code for summary
        elif county:
            # Normalize county name
            county = normalize_district_name(county)
            # Try to find VILLCODE
            key = f"{county}_{district}_{village}"
            if key in villcode_map:
                villcode = villcode_map[key]['VILLCODE']
        
        # Only add to summary if we have a VILLCODE
        if villcode:
            cunli_summary[villcode] = {
                'cunli': cunli_key,
                'district': district,
                'village': village,
                'county': county,
                'total_records': len(records),
                'sum_fields': sum_fields
            }
    
    # Save cunli summary file
    with open('docs/cunli_json/cunli_summary.json', 'w', encoding='utf-8') as f:
        json.dump(cunli_summary, f, ensure_ascii=False, indent=2)
    
    print(f"Created cunli_summary.json with {len(cunli_summary)} villages for fast map loading")
    
    # Generate recall cases list for dropdown
    recall_cases = set()
    for record in all_data:
        if record.get('recall_case'):
            recall_cases.add(record['recall_case'])
    
    # Clean up case names for better display
    def clean_case_name(case_name):
        """Clean case name for dropdown display"""
        import re
        
        # Remove file extensions and table prefixes
        name = case_name.replace('各投開票所投開票結果表', '').strip()
        
        # Extract key information using regex
        if '第11屆市長' in name:
            # Mayor recall case
            match = re.search(r'(\w+市)第11屆市長(\w+)罷免案', name)
            if match:
                return f"{match.group(1)}市長{match.group(2)}罷免案"
        else:
            # Legislator recall case
            match = re.search(r'第11屆立法委員\(([^)]+)\)(\w+)罷免案', name)
            if match:
                district = match.group(1)
                candidate = match.group(2)
                return f"立委{candidate}罷免案({district})"
        
        # Fallback: return original name
        return name
    
    # Create recall cases data structure with additional metadata
    # Create mapping between original and cleaned names
    case_mapping = {}
    cleaned_cases = []
    
    for case in recall_cases:
        cleaned_name = clean_case_name(case)
        cleaned_cases.append(cleaned_name)
        case_mapping[cleaned_name] = case  # Map cleaned name to original
    
    recall_cases_data = {
        'cases': sorted(cleaned_cases),
        'original_cases': sorted(list(recall_cases)),
        'case_mapping': case_mapping,  # Maps cleaned names to original names
        'total_cases': len(recall_cases),
        'generated_at': pd.Timestamp.now().isoformat(),
        'case_details': {}
    }
    
    # Add detailed statistics for each recall case
    for case in recall_cases:
        case_records = [r for r in all_data if r.get('recall_case') == case]
        case_villages = set()
        case_districts = set()
        case_villcodes = set()  # Track village codes for this case
        
        for record in case_records:
            if record.get('district') and record.get('village'):
                cunli_key = f"{record['district']}_{record['village']}"
                case_villages.add(cunli_key)
                case_districts.add(record['district'])
                
                # Find VILLCODE for this village
                district = record['district']
                village = record['village']
                
                # Extract county from recall case
                county = extract_county_from_recall_case(case)
                
                # Try to get VILLCODE
                if cunli_key in manual_villcode_map:
                    villcode = manual_villcode_map[cunli_key]['VILLCODE']
                    if isinstance(villcode, list):
                        case_villcodes.update(villcode)  # Add all codes from list
                    else:
                        case_villcodes.add(villcode)
                elif county:
                    # Normalize county name
                    county = normalize_district_name(county)
                    # Try to find VILLCODE
                    key = f"{county}_{district}_{village}"
                    if key in villcode_map:
                        case_villcodes.add(villcode_map[key]['VILLCODE'])
        
        # Calculate totals for this case
        case_totals = {
            'agree_votes': sum(r['agree_votes'] for r in case_records),
            'disagree_votes': sum(r['disagree_votes'] for r in case_records),
            'valid_votes': sum(r['valid_votes'] for r in case_records),
            'total_voters': sum(r['total_voters'] for r in case_records),
            'eligible_voters': sum(r['eligible_voters'] for r in case_records),
            'polling_stations': len(case_records),
            'villages': len(case_villages),
            'districts': len(case_districts),
            'village_codes': sorted(list(case_villcodes)),  # Add village codes list
            'cunli_keys': sorted(list(case_villages))  # Add cunli keys for fallback
        }
        
        # Calculate percentages
        if case_totals['valid_votes'] > 0:
            case_totals['agree_percentage'] = round(
                (case_totals['agree_votes'] / case_totals['valid_votes']) * 100, 2
            )
            case_totals['disagree_percentage'] = round(
                (case_totals['disagree_votes'] / case_totals['valid_votes']) * 100, 2
            )
        else:
            case_totals['agree_percentage'] = 0.0
            case_totals['disagree_percentage'] = 0.0
        
        if case_totals['eligible_voters'] > 0:
            case_totals['turnout_rate'] = round(
                (case_totals['total_voters'] / case_totals['eligible_voters']) * 100, 2
            )
        else:
            case_totals['turnout_rate'] = 0.0
        
        # Store details using cleaned name
        cleaned_name = clean_case_name(case)
        recall_cases_data['case_details'][cleaned_name] = case_totals
    
    # Save recall cases list
    with open('docs/cunli_json/recall_cases.json', 'w', encoding='utf-8') as f:
        json.dump(recall_cases_data, f, ensure_ascii=False, indent=2)
    
    print(f"Created recall_cases.json with {len(recall_cases)} unique recall cases")
    
    # Save missing VILLCODE mapping file for manual completion
    if missing_villcode:
        # Check if file already exists and preserve any manual entries
        existing_mappings = {}
        if os.path.exists('missing_villcode_mapping.json'):
            try:
                with open('missing_villcode_mapping.json', 'r', encoding='utf-8') as f:
                    existing_data = json.load(f)
                for item in existing_data:
                    if item.get('villcode'):  # If villcode was manually filled
                        existing_mappings[item['cunli_key']] = item
                print(f"Preserved {len(existing_mappings)} existing manual mappings")
            except Exception as e:
                print(f"Error reading existing mappings: {e}")
        
        # Update missing_villcode list with preserved mappings
        updated_missing = []
        for item in missing_villcode:
            key = item['cunli_key']
            if key in existing_mappings:
                # Use the existing manual mapping
                updated_missing.append(existing_mappings[key])
            else:
                # Use the new entry
                updated_missing.append(item)
        
        with open('missing_villcode_mapping.json', 'w', encoding='utf-8') as f:
            json.dump(updated_missing, f, ensure_ascii=False, indent=2)
        print(f"Created missing_villcode_mapping.json with {len(updated_missing)} entries for manual completion")
    
    print("\nExtraction and merging completed!")
    print(f"Created {len(villcode_files)} files with VILLCODE names")
    print(f"Created {len(cunli_merged) - len(villcode_files)} files with cunli names (no VILLCODE found)")
    print(f"Output files saved in 'docs/cunli_json' directory")

if __name__ == '__main__':
    main()