#!/usr/bin/env python3
import json
import csv
import os
from pathlib import Path

def load_cunli_data(cunli_json_dir):
    """Load all cunli JSON files"""
    cunli_data = {}
    
    for json_file in Path(cunli_json_dir).glob("*.json"):
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
                # Check if required fields exist
                if 'COUNTYNAME' not in data or 'TOWNNAME' not in data or 'VILLNAME' not in data:
                    print(f"Skipping {json_file.name}: Missing required fields")
                    print(f"  Available keys: {list(data.keys())}")
                    continue
                
                # Create key using COUNTYNAME+TOWNNAME+VILLNAME
                key = data['COUNTYNAME'] + data['TOWNNAME'] + data['VILLNAME']
                cunli_data[key] = data
        except Exception as e:
            print(f"Error processing {json_file.name}: {e}")
            continue
    
    return cunli_data

def load_election_data(election_json_path):
    """Load 2024 election results"""
    with open(election_json_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def combine_data(cunli_data, election_data):
    """Combine cunli and election data using the key"""
    combined_results = []
    
    for cunli_code, election_info in election_data.items():
        # Extract location info from election data name
        # Format: "縣市+區+里"
        location_parts = election_info['name'].split('縣')
        if len(location_parts) == 2:
            county_name = location_parts[0] + '縣'
            remaining = location_parts[1]
        else:
            location_parts = election_info['name'].split('市')
            if len(location_parts) >= 2:
                county_name = location_parts[0] + '市'
                remaining = '市'.join(location_parts[1:])
            else:
                continue
        
        # Further split town and village
        if '區' in remaining:
            town_parts = remaining.split('區')
            town_name = town_parts[0] + '區'
            village_name = '區'.join(town_parts[1:])
        elif '鎮' in remaining:
            town_parts = remaining.split('鎮')
            town_name = town_parts[0] + '鎮'
            village_name = '鎮'.join(town_parts[1:])
        elif '鄉' in remaining:
            town_parts = remaining.split('鄉')
            town_name = town_parts[0] + '鄉'
            village_name = '鄉'.join(town_parts[1:])
        else:
            continue
        
        # Create matching key
        match_key = county_name + town_name + village_name
        
        # Find matching cunli data
        if match_key in cunli_data:
            cunli_info = cunli_data[match_key]
            
            # Get winner info
            winner_name = ""
            winner_party = ""
            winner_votes = 0
            
            for candidate, vote_info in election_info['votes'].items():
                if vote_info['votes'] > winner_votes:
                    winner_votes = vote_info['votes']
                    winner_name = vote_info['name']
                    winner_party = vote_info['party']
            
            # Handle multiple recall cases - filter to match election zone
            recall_data = cunli_info['sum_fields'].copy()
            recall_case_name = cunli_info['records'][0]['recall_case'] if cunli_info['records'] else ""
            
            # Check if there are multiple different recall cases
            recall_cases = set(record['recall_case'] for record in cunli_info['records'])
            
            if len(recall_cases) > 1:
                # Multiple recall cases - need to filter to matching zone
                matching_records = []
                election_zone = election_info['zone']
                
                # Try to match recall case to election zone
                for record in cunli_info['records']:
                    recall_case = record['recall_case']
                    
                    # For 新竹市: separate legislator vs mayor recalls
                    if county_name == '新竹市':
                        if '鄭正鈐罷免案' in recall_case:
                            matching_records.append(record)
                    # For 基隆市: match 基隆市選舉區
                    elif county_name == '基隆市' and '基隆市選舉區' in recall_case:
                        matching_records.append(record)
                    # For 臺北市: match zone number in recall case
                    elif county_name == '臺北市':
                        if election_zone == '臺北市第01選區' and '臺北市第1選舉區' in recall_case:
                            matching_records.append(record)
                        elif election_zone == '臺北市第02選區' and '臺北市第2選舉區' in recall_case:
                            matching_records.append(record)
                        elif election_zone == '臺北市第03選區' and '臺北市第3選舉區' in recall_case:
                            matching_records.append(record)
                        elif election_zone == '臺北市第04選區' and '臺北市第4選舉區' in recall_case:
                            matching_records.append(record)
                        elif election_zone == '臺北市第05選區' and '臺北市第5選舉區' in recall_case:
                            matching_records.append(record)
                        elif election_zone == '臺北市第06選區' and '臺北市第6選舉區' in recall_case:
                            matching_records.append(record)
                        elif election_zone == '臺北市第07選區' and '臺北市第7選舉區' in recall_case:
                            matching_records.append(record)
                        elif election_zone == '臺北市第08選區' and '臺北市第8選舉區' in recall_case:
                            matching_records.append(record)
                
                # Use matching records if found, otherwise fall back to first case
                if matching_records:
                    recall_data = {
                        'agree_votes': sum(r['agree_votes'] for r in matching_records),
                        'disagree_votes': sum(r['disagree_votes'] for r in matching_records),
                        'valid_votes': sum(r['valid_votes'] for r in matching_records),
                        'invalid_votes': sum(r['invalid_votes'] for r in matching_records),
                        'eligible_voters': sum(r['eligible_voters'] for r in matching_records),
                    }
                    # Calculate turnout rate
                    total_voters = sum(r['total_voters'] for r in matching_records)
                    recall_data['average_turnout_rate'] = (total_voters / recall_data['eligible_voters'] * 100) if recall_data['eligible_voters'] > 0 else 0
                    recall_case_name = matching_records[0]['recall_case']
            
            # Combine the data
            combined_record = {
                'VILLCODE': cunli_info['VILLCODE'],
                'COUNTYNAME': cunli_info['COUNTYNAME'],
                'TOWNNAME': cunli_info['TOWNNAME'],
                'VILLNAME': cunli_info['VILLNAME'],
                'election_zone': election_info['zone'],
                'election_zone_code': election_info['zoneCode'],
                'total_election_votes': election_info['total'],
                'eligible_voters_election': election_info['votes_all'],
                'winner_name': winner_name,
                'winner_party': winner_party,
                'winner_votes': winner_votes,
                'recall_agree_votes': recall_data['agree_votes'],
                'recall_disagree_votes': recall_data['disagree_votes'],
                'recall_valid_votes': recall_data['valid_votes'],
                'recall_invalid_votes': recall_data['invalid_votes'],
                'recall_eligible_voters': recall_data['eligible_voters'],
                'recall_turnout_rate': recall_data['average_turnout_rate'],
                'recall_case': recall_case_name
            }
            
            # Add individual candidate vote data
            for i, (candidate, vote_info) in enumerate(election_info['votes'].items(), 1):
                combined_record[f'candidate_{i}_name'] = vote_info['name']
                combined_record[f'candidate_{i}_party'] = vote_info['party']
                combined_record[f'candidate_{i}_votes'] = vote_info['votes']
                combined_record[f'candidate_{i}_number'] = vote_info['no']
            
            combined_results.append(combined_record)
    
    return combined_results

def write_csv(data, output_file):
    """Write combined data to CSV"""
    if not data:
        print("No data to write")
        return
    
    # Get all possible field names
    all_fields = set()
    for record in data:
        all_fields.update(record.keys())
    
    # Sort fields for consistent output
    base_fields = [
        'VILLCODE', 'COUNTYNAME', 'TOWNNAME', 'VILLNAME',
        'election_zone', 'election_zone_code', 'total_election_votes', 'eligible_voters_election',
        'winner_name', 'winner_party', 'winner_votes',
        'recall_agree_votes', 'recall_disagree_votes', 'recall_valid_votes', 
        'recall_invalid_votes', 'recall_eligible_voters', 'recall_turnout_rate', 'recall_case'
    ]
    
    candidate_fields = sorted([f for f in all_fields if f.startswith('candidate_')])
    fieldnames = base_fields + candidate_fields
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)

def main():
    # Paths
    cunli_json_dir = "/home/kiang/public_html/recall-2025/docs/cunli_json"
    election_json_path = "/home/kiang/public_html/db.cec.gov.tw/data/ly/2024_zone_cunli.json"
    output_csv = "/home/kiang/public_html/recall-2025/cunli_combined_results.csv"
    
    print("Loading cunli data...")
    cunli_data = load_cunli_data(cunli_json_dir)
    print(f"Loaded {len(cunli_data)} cunli records")
    
    print("Loading 2024 election data...")
    election_data = load_election_data(election_json_path)
    print(f"Loaded {len(election_data)} election records")
    
    print("Combining data...")
    combined_data = combine_data(cunli_data, election_data)
    print(f"Combined {len(combined_data)} matching records")
    
    print(f"Writing results to {output_csv}...")
    write_csv(combined_data, output_csv)
    
    print("Done!")
    print(f"Results written to: {output_csv}")

if __name__ == "__main__":
    main()