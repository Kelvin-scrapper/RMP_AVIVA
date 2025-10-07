# Aviva Fund Data Extraction Pipeline

Automated pipeline for scraping Aviva Investors fund data and extracting key metrics from PDF factsheets.

## Overview

This project automates the complete workflow:
1. **Web Scraping**: Downloads fund factsheet PDFs from the Aviva website
2. **Data Extraction**: Parses PDFs and extracts portfolio statistics, country duration, and FX weights
3. **CSV Output**: Generates a structured CSV file with all extracted data

## Files

- **`orchestrator.py`** - Main pipeline orchestrator that runs the complete workflow
- **`main.py`** - Web scraper that downloads PDF factsheets from Aviva website
- **`map.py`** - PDF parser that extracts data from downloaded PDFs
- **`RMP_AVIVA_DATA_EXTRACTED.csv`** - Output file with extracted data

## Requirements

### Python Dependencies
Install all required packages using:

```bash
pip install -r requirements.txt
```

Or install manually:
```bash
pip install selenium undetected-chromedriver pandas pdfplumber
```

### System Requirements
- Python 3.8 or higher
- Google Chrome browser (compatible version will be downloaded automatically by undetected-chromedriver)

## Usage

### Run Complete Pipeline
```bash
python orchestrator.py
```

This will:
1. Scrape the Aviva website and download all fund factsheets
2. Extract data from all PDFs in the directory
3. Generate `RMP_AVIVA_DATA_EXTRACTED.csv` with the results

### Run Individual Scripts

**Download PDFs only:**
```bash
python main.py
```

**Extract data from existing PDFs:**
```bash
# Process all PDFs in current directory
python map.py

# Process specific PDF
python map.py download.pdf

# Process PDFs in specific directory
python map.py /path/to/pdfs/
```

## Configuration

### main.py Settings
- `HEADLESS_MODE = True` - Run Chrome in headless mode (no browser window)
- `DOWNLOAD_DIR` - Directory where PDFs will be saved (default: current directory)

### Behavior Features
- Human-like scrolling with random delays (0.5-1.5s)
- Random scroll amounts (200-500px)
- Automatic PDF download in both visible and headless modes

## Data Extracted

### Portfolio Statistics
- Yield to maturity (%)
- Modified duration
- Time to maturity
- Spread duration

### Country Duration (Top 5 overweights & underweights)
- Duration values by country
- Relative to benchmark values

### FX Weights (Top 5 overweights & underweights)
- Fund percentage by currency
- Relative to benchmark values

## Output Format

CSV file with two header rows:
1. **Data codes**: Machine-readable identifiers (e.g., `RMP.AVIVA.PORTSTST.YIELD.M`)
2. **Descriptive headers**: Human-readable descriptions

Data rows sorted chronologically by time period.

## Robustness Features

- **Order-independent table parsing**: Handles tables in any order
- **Dynamic country/currency mapping**: Adapts to new countries/currencies not in predefined maps
- **Flexible cell parsing**: Handles both combined cells ("Brazil 3.75") and separate columns
- **Multi-PDF processing**: Processes all PDFs in directory and combines into single CSV
- **Fixed output structure**: Maintains consistent column order regardless of PDF changes

## Notes

- Date is extracted from "AS AT" field (fallback to first date match)
- Output structure remains fixed even if PDF format changes
- Script automatically generates codes for new countries/currencies not in the predefined map
