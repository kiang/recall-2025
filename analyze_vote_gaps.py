#!/usr/bin/env python3
import csv
import pandas as pd

def analyze_vote_gaps():
    # Read the combined CSV
    df = pd.read_csv('cunli_combined_results.csv')
    
    # Initialize lists for KMT and non-KMT results
    kmt_results = []
    non_kmt_results = []
    
    for _, row in df.iterrows():
        # Basic info for both cases
        base_info = {
            'VILLCODE': row['VILLCODE'],
            'COUNTYNAME': row['COUNTYNAME'],
            'TOWNNAME': row['TOWNNAME'],
            'VILLNAME': row['VILLNAME'],
            'election_zone': row['election_zone'],
            'winner_name': row['winner_name'],
            'winner_party': row['winner_party'],
            'winner_votes': row['winner_votes'],
            'recall_agree_votes': row['recall_agree_votes'],
            'recall_disagree_votes': row['recall_disagree_votes'],
            'recall_valid_votes': row['recall_valid_votes'],
            'recall_eligible_voters': row['recall_eligible_voters'],
            'eligible_voters_election': row['eligible_voters_election']
        }
        
        if row['winner_party'] == '中國國民黨':
            # For KMT winners: compare winner_votes with recall_disagree_votes
            gap = row['winner_votes'] - row['recall_disagree_votes']
            kmt_info = base_info.copy()
            kmt_info['compared_with'] = 'recall_disagree_votes'
            kmt_info['gap'] = gap
            kmt_results.append(kmt_info)
        else:
            # For non-KMT winners: compare winner_votes with recall_agree_votes  
            gap = row['winner_votes'] - row['recall_agree_votes']
            non_kmt_info = base_info.copy()
            non_kmt_info['compared_with'] = 'recall_agree_votes'
            non_kmt_info['gap'] = gap
            non_kmt_results.append(non_kmt_info)
    
    # Convert to DataFrames and sort by gap in descending order
    kmt_df = pd.DataFrame(kmt_results)
    non_kmt_df = pd.DataFrame(non_kmt_results)
    
    kmt_df = kmt_df.sort_values('gap', ascending=False)
    non_kmt_df = non_kmt_df.sort_values('gap', ascending=False)
    
    # Define output columns in desired order
    output_columns = [
        'VILLCODE', 'COUNTYNAME', 'TOWNNAME', 'VILLNAME', 'election_zone',
        'winner_name', 'winner_party', 'winner_votes', 
        'recall_agree_votes', 'recall_disagree_votes', 'compared_with', 'gap',
        'recall_valid_votes', 'recall_eligible_voters', 'eligible_voters_election'
    ]
    
    # Save to CSV files
    kmt_df[output_columns].to_csv('kmt_winners_vote_gaps.csv', index=False, encoding='utf-8')
    non_kmt_df[output_columns].to_csv('non_kmt_winners_vote_gaps.csv', index=False, encoding='utf-8')
    
    # Print summary statistics
    print(f"KMT Winners Analysis:")
    print(f"  Total records: {len(kmt_df)}")
    print(f"  Largest gap: {kmt_df['gap'].max():,}")
    print(f"  Smallest gap: {kmt_df['gap'].min():,}")
    print(f"  Average gap: {kmt_df['gap'].mean():.1f}")
    print()
    
    print(f"Non-KMT Winners Analysis:")
    print(f"  Total records: {len(non_kmt_df)}")
    print(f"  Largest gap: {non_kmt_df['gap'].max():,}")
    print(f"  Smallest gap: {non_kmt_df['gap'].min():,}")
    print(f"  Average gap: {non_kmt_df['gap'].mean():.1f}")
    print()
    
    print("Top 5 largest gaps for KMT winners:")
    for i, row in kmt_df.head().iterrows():
        print(f"  {row['COUNTYNAME']}{row['TOWNNAME']}{row['VILLNAME']}: {row['gap']:,}")
    print()
    
    print("Top 5 largest gaps for non-KMT winners:")
    for i, row in non_kmt_df.head().iterrows():
        print(f"  {row['COUNTYNAME']}{row['TOWNNAME']}{row['VILLNAME']}: {row['gap']:,}")
    
    print(f"\nFiles generated:")
    print(f"  kmt_winners_vote_gaps.csv ({len(kmt_df)} records)")
    print(f"  non_kmt_winners_vote_gaps.csv ({len(non_kmt_df)} records)")

if __name__ == "__main__":
    analyze_vote_gaps()