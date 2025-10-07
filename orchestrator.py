"""
Orchestrator script that runs the complete data extraction pipeline:
1. Runs main.py to scrape and download PDFs from Aviva website
2. Runs map.py to extract data from all downloaded PDFs
"""

import subprocess
import sys
import os

def run_script(script_name, description):
    """
    Run a Python script and handle its output.
    Returns True if successful, False otherwise.
    """
    print(f"\n{'='*70}")
    print(f"STEP: {description}")
    print(f"{'='*70}\n")

    try:
        result = subprocess.run(
            [sys.executable, script_name],
            cwd=os.path.dirname(os.path.abspath(__file__)),
            capture_output=False,
            text=True
        )

        if result.returncode == 0:
            print(f"\n✓ {script_name} completed successfully")
            return True
        else:
            print(f"\n✗ {script_name} failed with return code {result.returncode}")
            return False

    except Exception as e:
        print(f"\n✗ Error running {script_name}: {str(e)}")
        return False

def main():
    """
    Main orchestrator function that runs the complete pipeline.
    """
    print("\n" + "="*70)
    print("AVIVA DATA EXTRACTION PIPELINE")
    print("="*70)

    # Step 1: Run main.py to download PDFs
    success_scrape = run_script('main.py', 'Scraping Aviva website and downloading PDFs')

    if not success_scrape:
        print("\n" + "="*70)
        print("PIPELINE ABORTED: Scraping failed")
        print("="*70)
        sys.exit(1)

    # Step 2: Run map.py to extract data from PDFs
    success_extract = run_script('map.py', 'Extracting data from downloaded PDFs')

    if not success_extract:
        print("\n" + "="*70)
        print("PIPELINE ABORTED: Data extraction failed")
        print("="*70)
        sys.exit(1)

    # Success
    print("\n" + "="*70)
    print("PIPELINE COMPLETED SUCCESSFULLY")
    print("="*70)
    print("\nAll tasks completed:")
    print("1. PDFs downloaded from Aviva website")
    print("2. Data extracted and saved to RMP_AVIVA_DATA_EXTRACTED.csv")
    print("="*70 + "\n")

if __name__ == "__main__":
    main()
