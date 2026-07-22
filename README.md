# CPE Sector Alignment & Issue Analyzer

A toolset designed to audit CPE (Customer Premises Equipment) alignment against serving sectors, identify anomalies, and automatically categorize root causes of misalignment (e.g. database errors vs. coordinates errors).

---

## Important Notice: Security & Placeholders

For security, privacy, and sharing readiness:
1. **API Configurations & Endpoint Settings**: This tool operates offline using local database exports.
2. **Demo Database Data**: All input Excel, CSV, and JSON data files (`cpe_data_hybrid.xlsx`, `sector_details_complete.csv`, `towers_cache.json`, and `subnets.csv`) are populated with **mock placeholder data** (`TOWER-A`, `TOWER-B`, and dummy clients) instead of real production coordinates or customer information. Use these mock entries to test the tool, or replace them with your own network datasets.

---

## Features

1. **CPE Sector Alignment Audit (`analyze_cpe_sectors.py`)**:
   - Calculates compass bearings (0–359°) from the home tower to each CPE.
   - Cross-references the CPE's actual bearing against the serving sector's azimuth and beamwidth.
   - Audits whether the CPE is physically facing the sector it is registered to in cnMaestro or IPAM.
   - Highlights anomalies: e.g. CPEs connected to sectors facing the opposite direction.
   - Generates two output sheets:
     - `cpe_data_hybrid_analyzed.xlsx`: Full audit database with bearing calculations and flags.
     - `cpe_data_hybrid_issues.xlsx`: Filtered sheet containing only devices failing alignment or distance criteria.

2. **Issue Categorization (`categorize_issues.py`)**:
   - Group-analyzes failures at a sector level.
   - **Status Categories**:
     - *Wrong Azimuth (Master DB Error)*: If a high percentage of CPEs connected to a sector fail the facing audit, the sector's azimuth configuration in the master database is likely wrong.
     - *Healthy DB (Check individual CPEs)*: If most CPEs face the sector correctly, but a small handful fail, the individual failing CPEs likely have incorrect GPS coordinates entered.
     - *Mixed Results (Review required)*: Borderline scenarios requiring engineering review.
   - Generates a final spreadsheet report: `issue_categorization_report.xlsx`.

---

## File Structure

```
Sector Analyzer/
├── analyze_cpe_sectors.py   # Primary alignment auditor
├── categorize_issues.py      # Automated root-cause classifier
├── cpe_data_hybrid.xlsx      # Mock workbook of CPE client coordinates and connected AP names
├── sector_details_complete.csv # Mock workbook of sector azimuths and beamwidths (Excel format)
├── towers_cache.json        # Mock database of tower coordinates
├── subnets.csv              # Mock subnet prefix-to-tower map
├── requirements.txt         # Python package dependencies
└── README.md                # Project documentation
```

---

## Installation & Setup

1. Clone or download this directory.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage Instructions (Using Demo Data)

You can run these commands out of the box using the included demo files.

### 1. Run the CPE Sector Alignment Audit
```bash
python analyze_cpe_sectors.py
```
This reads `cpe_data_hybrid.xlsx`, maps each CPE geographically, audits its alignment, and outputs:
- `cpe_data_hybrid_analyzed.xlsx`
- `cpe_data_hybrid_issues.xlsx`

### 2. Run the Root Cause Categorizer
```bash
python categorize_issues.py
```
This reads `cpe_data_hybrid_analyzed.xlsx` and outputs a health summary and specific human-error candidates in:
- `issue_categorization_report.xlsx`

---

## Data Schema & Format

### `sector_details_complete.csv`
*Note: Although it has a `.csv` extension, it is physically saved as an Excel workbook.*
Required headers:
- `Site Name` (Tower location)
- `Azimut` (Sector boresight heading)
- `Beam width` (Sector coverage arc in degrees)
- `Sector Name` (System name of the sector AP)

### `cpe_data_hybrid.xlsx`
Required headers:
- `IP Address`
- `IPAM Name`
- `IPAM Sector`
- `cnMaestro AP Name`
- `IPAM SM Latitude`
- `IPAM SM Longitude`

---

*Disclaimer: Internal Network Tools — Use with appropriate credentials.*
