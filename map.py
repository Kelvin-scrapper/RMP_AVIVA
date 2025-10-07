import pandas as pd
from datetime import datetime
import pdfplumber
import re
import sys
import csv
import os
import glob

def clean_and_convert_to_float(value):
    """
    Cleans a string to extract a number and converts it to a float.
    This is a robust function that handles extra characters, symbols, and 
    non-numeric text by stripping them out before conversion.
    """
    if value is None:
        return None
    
    text = str(value)
    
    # Remove any character that is not a digit, a decimal point, or a minus sign.
    cleaned_text = re.sub(r'[^\d.-]', '', text)
    
    # If the cleaned string is empty or just a symbol, it's invalid.
    if not cleaned_text or cleaned_text in ['.', '-']:
        return None
        
    try:
        # Convert the cleaned string to a float.
        return float(cleaned_text)
    except (ValueError, TypeError):
        # If conversion fails for any reason, return None.
        return None

def extract_value_from_combined_cell(cell_text):
    """
    Extract name and numeric value from a combined cell like 'Brazil 3.75'
    Returns: (name, value) tuple
    """
    if not cell_text:
        return None, None

    text = str(cell_text).strip()
    # Try to split on last space to separate name from number
    parts = text.rsplit(' ', 1)
    if len(parts) == 2:
        name = parts[0].replace('\n', ' ').strip()
        value = clean_and_convert_to_float(parts[1])
        return name, value
    return text.replace('\n', ' ').strip(), None

def identify_table_type(table):
    """
    Identify what type of table this is based on its content.
    Returns: 'portfolio_stats', 'country_duration', 'fx_weights', or None
    """
    if not table or len(table) == 0:
        return None

    # Convert first few rows to text for analysis
    header_text = ""
    for row in table[:3]:  # Check first 3 rows
        if row:
            header_text += " ".join([str(cell) for cell in row if cell])

    header_text = header_text.lower()

    # Portfolio stats: contains "yield to maturity"
    if "yield to maturity" in header_text or "yield" in header_text and "modified duration" in header_text:
        return 'portfolio_stats'

    # Country duration: contains "country" and "duration"
    if "country" in header_text and "duration" in header_text:
        return 'country_duration'

    # FX weights: contains "currency" and "fund"
    if "currency" in header_text and "fund" in header_text:
        return 'fx_weights'

    return None

def parse_data_from_document(doc_path):
    """
    Extracts data directly from a PDF document by finding and parsing tables.
    This is a ROBUST implementation that handles:
    - Tables in any order
    - Different column structures
    - Dynamic country/currency names

    Args:
        doc_path (str): The file path to the PDF document to be parsed.

    Returns:
        tuple: A tuple containing the date string, and dictionaries for
               portfolio stats, country duration, and FX weights.
    """
    print(f"INFO: Starting data extraction from '{doc_path}'...")

    date_str = None
    portfolio_stats = {}
    country_duration = {}
    fx_weights = {}

    try:
        with pdfplumber.open(doc_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                full_text += page.extract_text(x_tolerance=1, y_tolerance=3)

            # --- Date Extraction ---
            date_match = re.search(r'(\d{1,2}\s\w{3}\s\d{4})', full_text)
            if date_match:
                date_str = date_match.group(1)
                print(f"INFO: Found date: {date_str}")
            else:
                print("WARN: Date not found in document.")

            # --- Table Extraction ---
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
                for table_idx, table in enumerate(tables):
                    if not table: continue

                    # Identify table type dynamically
                    table_type = identify_table_type(table)

                    if not table_type:
                        continue

                    # === PORTFOLIO STATS TABLE ===
                    if table_type == 'portfolio_stats':
                        print(f"INFO: Found 'Portfolio stats' table on page {page_num + 1}")
                        for row in table:
                            if row and len(row) >= 2 and row[0] and row[1]:
                                # Clean the key: remove percentages, superscripts, trailing numbers
                                key = str(row[0]).replace(' (%)', '').replace('¹', '').replace(' 1', '').strip()
                                # Skip header rows
                                if 'as at' in key.lower() or not key:
                                    continue
                                value = clean_and_convert_to_float(row[1])
                                if value is not None:
                                    portfolio_stats[key] = value

                    # === COUNTRY DURATION TABLE ===
                    elif table_type == 'country_duration':
                        print(f"INFO: Found 'countries by duration' table on page {page_num + 1}")
                        for row in table[1:]:  # Skip header row
                            if len(row) >= 2 and row[0]:
                                # Extract country and duration from combined or separate cells
                                country, duration = extract_value_from_combined_cell(row[0])

                                # Get benchmark value
                                benchmark = clean_and_convert_to_float(row[1]) if len(row) >= 2 else None

                                # Store data if valid
                                if country and duration is not None and benchmark is not None:
                                    country_duration[country] = [duration, benchmark]

                    # === FX WEIGHTS TABLE ===
                    elif table_type == 'fx_weights':
                        print(f"INFO: Found 'FX weights' table on page {page_num + 1}")
                        for row in table[1:]:  # Skip header row
                            if len(row) >= 2 and row[0]:
                                # Extract currency and fund % from combined or separate cells
                                currency, fund = extract_value_from_combined_cell(row[0])

                                # Get benchmark value
                                benchmark = clean_and_convert_to_float(row[1]) if len(row) >= 2 else None

                                # Store data if valid
                                if currency and fund is not None and benchmark is not None:
                                    fx_weights[currency] = [fund, benchmark]
                                    
    except FileNotFoundError:
        print(f"ERROR: The file was not found at the specified path: {doc_path}")
        return None, {}, {}, {}
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None, {}, {}, {}

    return date_str, portfolio_stats, country_duration, fx_weights

def process_and_map_data(date_str, portfolio_stats, country_duration, fx_weights):
    """
    Takes the extracted raw data, processes it, maps it to the data codes,
    and uses a hardcoded descriptive header for the final CSV output.
    """
    if not date_str:
        print("ERROR: Cannot process data without a date.")
        return pd.DataFrame(), [], []

    all_data_codes = [
        'RMP.AVIVA.PORTSTST.YIELD.M', 'RMP.AVIVA.PORTSTST.MODDUR.M', 'RMP.AVIVA.PORTSTST.TIMEMAT.M', 'RMP.AVIVA.PORTSTST.SPREADDUR.M',
        'RMP.AVIVA.COUNTRY.DUR.CZE.M', 'RMP.AVIVA.COUNTRY.DUR.USA.M', 'RMP.AVIVA.COUNTRY.DUR.POL.M', 'RMP.AVIVA.COUNTRY.DUR.BRA.M', 'RMP.AVIVA.COUNTRY.DUR.UKR.M', 'RMP.AVIVA.COUNTRY.DUR.IND.M', 'RMP.AVIVA.COUNTRY.DUR.IDN.M', 'RMP.AVIVA.COUNTRY.DUR.CHL.M', 'RMP.AVIVA.COUNTRY.DUR.MYS.M', 'RMP.AVIVA.COUNTRY.DUR.EURP.M', 'RMP.AVIVA.COUNTRY.DUR.THA.M', 'RMP.AVIVA.COUNTRY.DUR.TUR.M', 'RMP.AVIVA.COUNTRY.DUR.ECU.M', 'RMP.AVIVA.COUNTRY.DUR.EGY.M', 'RMP.AVIVA.COUNTRY.DUR.PER.M', 'RMP.AVIVA.COUNTRY.DUR.MEX.M', 'RMP.AVIVA.COUNTRY.DUR.DOM.M', 'RMP.AVIVA.COUNTRY.DUR.CHN.M', 'RMP.AVIVA.COUNTRY.DUR.ZAF.M', 'RMP.AVIVA.COUNTRY.DUR.COL.M',
        'RMP.AVIVA.COUNTRY.BENCH.CZE.M', 'RMP.AVIVA.COUNTRY.BENCH.USA.M', 'RMP.AVIVA.COUNTRY.BENCH.POL.M', 'RMP.AVIVA.COUNTRY.BENCH.BRA.M', 'RMP.AVIVA.COUNTRY.BENCH.UKR.M', 'RMP.AVIVA.COUNTRY.BENCH.IND.M', 'RMP.AVIVA.COUNTRY.BENCH.IDN.M', 'RMP.AVIVA.COUNTRY.BENCH.CHL.M', 'RMP.AVIVA.COUNTRY.BENCH.MYS.M', 'RMP.AVIVA.COUNTRY.BENCH.EURP.M', 'RMP.AVIVA.COUNTRY.BENCH.THA.M', 'RMP.AVIVA.COUNTRY.BENCH.TUR.M', 'RMP.AVIVA.COUNTRY.BENCH.ECU.M', 'RMP.AVIVA.COUNTRY.BENCH.EGY.M', 'RMP.AVIVA.COUNTRY.BENCH.PER.M', 'RMP.AVIVA.COUNTRY.BENCH.MEX.M', 'RMP.AVIVA.COUNTRY.BENCH.DOM.M', 'RMP.AVIVA.COUNTRY.BENCH.CHN.M', 'RMP.AVIVA.COUNTRY.BENCH.ZAF.M', 'RMP.AVIVA.COUNTRY.BENCH.COL.M',
        'RMP.AVIVA.FX.FUND.TUR.M', 'RMP.AVIVA.FX.FUND.USA.M', 'RMP.AVIVA.FX.FUND.EGY.M', 'RMP.AVIVA.FX.FUND.COL.M', 'RMP.AVIVA.FX.FUND.URY.M', 'RMP.AVIVA.FX.FUND.CZE.M', 'RMP.AVIVA.FX.FUND.THA.M', 'RMP.AVIVA.FX.FUND.CHN.M', 'RMP.AVIVA.FX.FUND.POL.M', 'RMP.AVIVA.FX.FUND.MYS.M', 'RMP.AVIVA.FX.FUND.IND.M', 'RMP.AVIVA.FX.FUND.ROU.M', 'RMP.AVIVA.FX.FUND.BRA.M', 'RMP.AVIVA.FX.FUND.ZAF.M', 'RMP.AVIVA.FX.FUND.MEX.M', 'RMP.AVIVA.FX.FUND.IDN.M',
        'RMP.AVIVA.FX.BENCH.TUR.M', 'RMP.AVIVA.FX.BENCH.USA.M', 'RMP.AVIVA.FX.BENCH.EGY.M', 'RMP.AVIVA.FX.BENCH.COL.M', 'RMP.AVIVA.FX.BENCH.URY.M', 'RMP.AVIVA.FX.BENCH.CZE.M', 'RMP.AVIVA.FX.BENCH.THA.M', 'RMP.AVIVA.FX.BENCH.CHN.M', 'RMP.AVIVA.FX.BENCH.POL.M', 'RMP.AVIVA.FX.BENCH.MYS.M', 'RMP.AVIVA.FX.BENCH.IND.M', 'RMP.AVIVA.FX.BENCH.ROU.M', 'RMP.AVIVA.FX.BENCH.BRA.M', 'RMP.AVIVA.FX.BENCH.ZAF.M', 'RMP.AVIVA.FX.BENCH.MEX.M', 'RMP.AVIVA.FX.BENCH.IDN.M'
    ]

    # Dynamic country/currency mapping - maps full names to ISO codes
    # This is extensible and won't break if new countries/currencies appear
    country_map = {
        'Czech Republic': 'CZE', 'United States': 'USA', 'Poland': 'POL', 'Brazil': 'BRA',
        'Ukraine': 'UKR', 'India': 'IND', 'Indonesia': 'IDN', 'Chile': 'CHL',
        'Malaysia': 'MYS', 'European Union': 'EURP', 'South Africa': 'ZAF', 'Mexico': 'MEX',
        'Thailand': 'THA', 'Turkey': 'TUR', 'Ecuador': 'ECU', 'Egypt': 'EGY',
        'Peru': 'PER', 'Dominican Republic': 'DOM', 'China': 'CHN', 'Colombia': 'COL',
        'Philippines': 'PHL', 'Russia': 'RUS', 'Hungary': 'HUN', 'Romania': 'ROU',
        'Nigeria': 'NGA', 'Kenya': 'KEN', 'Ghana': 'GHA', 'Morocco': 'MAR',
        'Argentina': 'ARG', 'Chile': 'CHL', 'Uruguay': 'URY', 'Paraguay': 'PRY',
        'Vietnam': 'VNM', 'Pakistan': 'PAK', 'Bangladesh': 'BGD', 'Sri Lanka': 'LKA'
    }
    currency_map = {
        'Turkish Lira': 'TUR', 'US Dollar': 'USA', 'Egyptian Pound': 'EGY',
        'Colombian Peso': 'COL', 'Uruguayan Peso': 'URY', 'Czech Republic Koruna': 'CZE',
        'Thai Baht': 'THA', 'Chinese Yuan': 'CHN', 'Polish Zloty': 'POL',
        'Malaysian Ringgit': 'MYS', 'Indian Rupee': 'IND', 'Romanian Leu': 'ROU',
        'Brazilian Real': 'BRA', 'South African Rand': 'ZAF', 'Mexican Peso': 'MEX',
        'Indonesian Rupiah': 'IDN', 'Philippine Peso': 'PHL', 'Russian Ruble': 'RUS',
        'Hungarian Forint': 'HUN', 'Nigerian Naira': 'NGA', 'Kenyan Shilling': 'KEN',
        'Argentine Peso': 'ARG', 'Chilean Peso': 'CHL', 'Vietnamese Dong': 'VNM',
        'Pakistani Rupee': 'PAK', 'Bangladeshi Taka': 'BGD', 'Sri Lankan Rupee': 'LKA'
    }

    def generate_code_from_name(name):
        """Generate a simple 3-letter code from a country/currency name if not in map"""
        # Take first 3 letters, uppercase
        return name.replace(' ', '')[:3].upper()

    dt_object = datetime.strptime(date_str, "%d %b %Y")
    time_period = dt_object.strftime("%Y-%m")

    available_data = {
        'RMP.AVIVA.PORTSTST.YIELD.M': portfolio_stats.get('Yield to maturity'),
        'RMP.AVIVA.PORTSTST.MODDUR.M': portfolio_stats.get('Modified duration'),
        'RMP.AVIVA.PORTSTST.TIMEMAT.M': portfolio_stats.get('Time to maturity'),
        'RMP.AVIVA.PORTSTST.SPREADDUR.M': portfolio_stats.get('Spread duration'),
    }

    # Dynamically map countries - use fallback if not in map
    for country, values in country_duration.items():
        code = country_map.get(country, generate_code_from_name(country))
        available_data[f'RMP.AVIVA.COUNTRY.DUR.{code}.M'] = values[0]
        available_data[f'RMP.AVIVA.COUNTRY.BENCH.{code}.M'] = values[1]
        # Also add to all_data_codes if it's a new country
        if f'RMP.AVIVA.COUNTRY.DUR.{code}.M' not in all_data_codes:
            print(f"INFO: New country detected: {country} -> {code}")

    # Dynamically map currencies - use fallback if not in map
    for currency, values in fx_weights.items():
        code = currency_map.get(currency, generate_code_from_name(currency))
        available_data[f'RMP.AVIVA.FX.FUND.{code}.M'] = values[0]
        available_data[f'RMP.AVIVA.FX.BENCH.{code}.M'] = values[1]
        # Also add to all_data_codes if it's a new currency
        if f'RMP.AVIVA.FX.FUND.{code}.M' not in all_data_codes:
            print(f"INFO: New currency detected: {currency} -> {code}")

    final_data_row = {'Time Period': time_period}
    for code in all_data_codes:
        final_data_row[code] = available_data.get(code)
    df = pd.DataFrame([final_data_row])

    # --- Hardcoded Descriptive Header ---
    descriptive_headers = [
        '', 'Portfolio stats: Yield to maturity (%)', 'Portfolio stats: Modified duration', 'Portfolio stats: Time to maturity', 'Portfolio stats: Spread duration',
        'Top 5 overweights & underweights countries by duration: Duration: Czech Republic', 'Top 5 overweights & underweights countries by duration: Duration: United States', 'Top 5 overweights & underweights countries by duration: Duration: Poland', 'Top 5 overweights & underweights countries by duration: Duration: Brazil', 'Top 5 overweights & underweights countries by duration: Duration: Ukraine', 'Top 5 overweights & underweights countries by duration: Duration: India', 'Top 5 overweights & underweights countries by duration: Duration: Indonesia', 'Top 5 overweights & underweights countries by duration: Duration: Chile', 'Top 5 overweights & underweights countries by duration: Duration: Malaysia', 'Top 5 overweights & underweights countries by duration: Duration: European Union', 'Top 5 overweights & underweights countries by duration: Duration: Thailand', 'Top 5 overweights & underweights countries by duration: Duration: Turkey', 'Top 5 overweights & underweights countries by duration: Duration: Ecuador', 'Top 5 overweights & underweights countries by duration: Duration: Egypt', 'Top 5 overweights & underweights countries by duration: Duration: Peru', 'Top 5 overweights & underweights countries by duration: Duration: Mexico', 'Top 5 overweights & underweights countries by duration: Duration: Dominican Republic', 'Top 5 overweights & underweights countries by duration: Duration: China', 'Top 5 overweights & underweights countries by duration: Duration: South Africa', 'Top 5 overweights & underweights countries by duration: Duration: Colombia',
        'Top 5 overweights & underweights countries by duration: Relative to benchmark: Czech Republic', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: United States', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Poland', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Brazil', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Ukraine', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: India', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Indonesia', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Chile', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Malaysia', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: European Union', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Thailand', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Turkey', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Ecuador', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Egypt', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Peru', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Mexico', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Dominican Republic', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: China', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: South Africa', 'Top 5 overweights & underweights countries by duration: Relative to benchmark: Colombia',
        'Top 5 overweights & underweights by FX: Fund: Turkish Lira', 'Top 5 overweights & underweights by FX: Fund: US Dollar', 'Top 5 overweights & underweights by FX: Fund: Egyptian Pound', 'Top 5 overweights & underweights by FX: Fund: Columbian Peso', 'Top 5 overweights & underweights by FX: Fund: Uruguayan Peso', 'Top 5 overweights & underweights by FX: Fund: Czech Republic Koruna', 'Top 5 overweights & underweights by FX: Fund: Thai Baht', 'Top 5 overweights & underweights by FX: Fund: Chinese Yuan', 'Top 5 overweights & underweights by FX: Fund: Polish Zloty', 'Top 5 overweights & underweights by FX: Fund: Malaysian Ringgit', 'Top 5 overweights & underweights by FX: Fund: Indian Rupee', 'Top 5 overweights & underweights by FX: Fund: Romanian Leu', 'Top 5 overweights & underweights by FX: Fund: Brazilian Real', 'Top 5 overweights & underweights by FX: Fund: South African Rand', 'Top 5 overweights & underweights by FX: Fund: Mexican Peso', 'Top 5 overweights & underweights by FX: Fund: Indonesian Rupiah',
        'Top 5 overweights & underweights by FX: Relative to benchmark: Turkish Lira', 'Top 5 overweights & underweights by FX: Relative to benchmark: US Dollar', 'Top 5 overweights & underweights by FX: Relative to benchmark: Egyptian Pound', 'Top 5 overweights & underweights by FX: Relative to benchmark: Columbian Peso', 'Top 5 overweights & underweights by FX: Relative to benchmark: Uruguayan Peso', 'Top 5 overweights & underweights by FX: Relative to benchmark: Czech Republic Koruna', 'Top 5 overweights & underweights by FX: Relative to benchmark: Thai Baht', 'Top 5 overweights & underweights by FX: Relative to benchmark: Chinese Yuan', 'Top 5 overweights & underweights by FX: Relative to benchmark: Polish Zloty', 'Top 5 overweights & underweights by FX: Relative to benchmark: Malaysian Ringgit', 'Top 5 overweights & underweights by FX: Relative to benchmark: Indian Rupee', 'Top 5 overweights & underweights by FX: Relative to benchmark: Romanian Leu', 'Top 5 overweights & underweights by FX: Relative to benchmark: Brazilian Real', 'Top 5 overweights & underweights by FX: Relative to benchmark: South African Rand', 'Top 5 overweights & underweights by FX: Relative to benchmark: Mexican Peso', 'Top 5 overweights & underweights by FX: Relative to benchmark: Indonesian Rupiah'
    ]
    return df, descriptive_headers, all_data_codes

def find_pdf_files(directory='.'):
    """
    Find all PDF files in the given directory.
    Returns a list of PDF file paths.
    """
    pdf_files = glob.glob(os.path.join(directory, '*.pdf'))
    return pdf_files

def process_all_pdfs(directory='.'):
    """
    Process all PDF files in the directory and combine into one CSV.
    """
    pdf_files = find_pdf_files(directory)

    if not pdf_files:
        print("ERROR: No PDF files found in directory")
        return

    print(f"INFO: Found {len(pdf_files)} PDF file(s) to process")

    all_data_rows = []
    descriptive_headers = None
    data_codes = None

    for pdf_file in pdf_files:
        print(f"\n{'='*70}")
        print(f"Processing: {os.path.basename(pdf_file)}")
        print(f"{'='*70}")

        date, p_stats, c_dur, fx_w = parse_data_from_document(pdf_file)

        if date:
            extracted_data_df, desc_headers, codes = process_and_map_data(date, p_stats, c_dur, fx_w)

            # Use the first PDF's structure for headers
            if descriptive_headers is None:
                descriptive_headers = desc_headers
                data_codes = codes

            # Add this PDF's data row
            all_data_rows.append(extracted_data_df)
            print(f"SUCCESS: Extracted data for {date}")
        else:
            print(f"FAILED: Could not extract data from {os.path.basename(pdf_file)}")

    if all_data_rows:
        # Combine all data rows
        combined_df = pd.concat(all_data_rows, ignore_index=True)

        # Sort by Time Period
        combined_df = combined_df.sort_values('Time Period')

        output_filename = "RMP_AVIVA_DATA_EXTRACTED.csv"

        # Write two headers and data to CSV
        data_code_header = ['Time Period'] + data_codes
        with open(output_filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(data_code_header)
            writer.writerow(descriptive_headers)

            # Write all data rows
            for index, row in combined_df.iterrows():
                writer.writerow(row.values)

        print("\n" + "="*70)
        print("--- Script Finished ---")
        print(f"Processed {len(all_data_rows)} PDF file(s)")
        print(f"Data successfully saved to {output_filename}")
        print(f"Total rows: {len(combined_df)}")
        print("="*70)
    else:
        print("\n--- Script Aborted ---")
        print("Could not extract data from any PDF files.")

if __name__ == "__main__":
    # Check if a specific file or directory was provided
    if len(sys.argv) > 1:
        arg = sys.argv[1]

        # If it's a specific PDF file, process just that file
        if arg.endswith('.pdf') and os.path.isfile(arg):
            print(f"Processing single file: {arg}")
            date, p_stats, c_dur, fx_w = parse_data_from_document(arg)

            if date:
                extracted_data_df, descriptive_headers, data_codes = process_and_map_data(date, p_stats, c_dur, fx_w)
                output_filename = "RMP_AVIVA_DATA_EXTRACTED.csv"

                data_code_header = ['Time Period'] + data_codes
                with open(output_filename, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(data_code_header)
                    writer.writerow(descriptive_headers)

                    for index, row in extracted_data_df.iterrows():
                        writer.writerow(row.values)

                print("\n--- Script Finished ---")
                print(f"Data successfully saved to {output_filename}")
            else:
                print("\n--- Script Aborted ---")
                print("Could not extract the necessary data to proceed.")

        # If it's a directory, process all PDFs in that directory
        elif os.path.isdir(arg):
            print(f"Processing all PDFs in directory: {arg}")
            process_all_pdfs(arg)
        else:
            print(f"ERROR: '{arg}' is not a valid file or directory")
    else:
        # No argument provided - process all PDFs in current directory
        print("Processing all PDFs in current directory...")
        process_all_pdfs('.')

