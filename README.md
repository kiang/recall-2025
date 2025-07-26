# Taiwan 2025 Recall Election Results Data Processing

This project processes and visualizes voting results data from Taiwan's 2025 recall elections.

## Project Overview

This project extracts voting data from Excel files containing 2025 recall election results and organizes them into JSON files named using Taiwan's official village codes (VILLCODE) for precise geographic identification.

## Features

- **Data Extraction**: Extracts 6,303 voting records from 25 Excel files
- **Geographic Mapping**: Maps voting data to 2,280 villages using official VILLCODE identifiers
- **Data Aggregation**: Aggregates polling station results by village (cunli)
- **Performance Optimization**: Generates summary files for faster map loading
- **Complete Coverage**: 100% of villages have corresponding VILLCODE identifiers

## File Structure

```
recall-2025/
├── raw/                           # Raw Excel files
│   ├── 表5-第11屆立法委員(...).xlsx
│   └── ...
├── docs/                          # Output data and web files
│   ├── index.html                 # Redirect page
│   └── cunli_json/               # Village voting data
│       ├── cunli_summary.json    # Fast loading summary (879KB)
│       ├── summary.json          # Overall statistics
│       ├── 63000020001.json      # Individual village data (named by VILLCODE)
│       └── ...
├── extract_and_merge.py          # Main data processing script
├── missing_villcode_mapping.json # Manual VILLCODE mappings
└── README.md
```

## Usage

### Requirements

```bash
pip install pandas openpyxl
```

### Run Data Processing

```bash
python3 extract_and_merge.py
```

The script will automatically:
1. Read all Excel files from the `raw/` directory
2. Extract voting data and group by village
3. Map to official VILLCODE (using taiwan_basecode data)
4. Generate individual village JSON files and summary files
5. Preserve any manually entered VILLCODE mappings

### Output Files

- **Individual Village Files**: `docs/cunli_json/{VILLCODE}.json`
  - Contains detailed data for all polling stations in the village
  - Includes aggregated statistics

- **Summary File**: `docs/cunli_json/cunli_summary.json`
  - Basic information and aggregated data for all villages
  - Used for fast map loading (879KB vs 11MB)

## Data Format

### Village Data Structure
```json
{
  "cunli": "信義區_西村里",
  "district": "信義區", 
  "village": "西村里",
  "county": "台北市",
  "villcode": "63000020001",
  "sum_fields": {
    "agree_votes": 691,
    "disagree_votes": 1541,
    "valid_votes": 2232,
    "invalid_votes": 14,
    "total_voters": 2246,
    "eligible_voters": 3532,
    "average_turnout_rate": 63.59
  },
  "records": [
    {
      "polling_station": "001",
      "recall_case": "第11屆立法委員(臺北市第7選舉區)徐巧芯罷免案",
      "agree_votes": 172,
      "disagree_votes": 395,
      "valid_votes": 567,
      "invalid_votes": 4,
      "total_voters": 571,
      "eligible_voters": 871,
      "turnout_rate": 65.56
    }
  ]
}
```

## Technical Features

### VILLCODE Mapping
- Uses Taiwan's official village codes for geographic precision
- Automatic mapping + manual completion achieves 100% coverage
- Supports Traditional Chinese character variants (台/臺)

### Performance Optimization
- Summary file generation reduces initial loading time (92% improvement)
- On-demand loading of detailed data
- Filters villages without data to save memory

### Data Quality
- Comprehensive error handling and validation
- Automatic detection of Excel header row positions
- Preserves manual mappings to prevent overwrites

## Web Display

Originally included an interactive map, now redirects to: https://tainan.olc.tw/p/recall-2025/

## License

MIT License - see [LICENSE](LICENSE) file for details

## Contributors

- Finjon Kiang - Initial development

## Data Sources

- 2025 Recall Election Results: Central Election Commission, Taiwan
- Taiwan Village Boundary Data: [taiwan_basecode](https://github.com/kiang/taiwan_basecode)