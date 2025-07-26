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
    
    # Save individual cunli JSON files
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
        
        output_file = f"cunli_json/{cunli_key}.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump({
                'cunli': cunli_key,
                'total_records': len(records),
                'sum_fields': sum_fields,
                'records': records
            }, f, ensure_ascii=False, indent=2)
    
    # Save summary file
    summary = {
        'total_files_processed': len(excel_files),
        'total_records': len(all_data),
        'total_cunli': len(cunli_merged),
        'cunli_list': sorted(list(cunli_merged.keys()))
    }
    
    with open('cunli_json/summary.json', 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    
    print("\nExtraction and merging completed!")
    print(f"Output files saved in 'cunli_json' directory")

if __name__ == '__main__':
    main()