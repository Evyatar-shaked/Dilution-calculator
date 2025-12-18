"""
Command-line interface for batch processing dilutions from Excel files.
Run with: python batch_process.py <input_file.xlsx> [--output-unit mL]
"""

import sys
import argparse
from dilution_core import process_excel_dilutions


def main():
    parser = argparse.ArgumentParser(
        description='Process batch dilution calculations from an Excel file.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python batch_process.py my_dilutions.xlsx
  python batch_process.py my_dilutions.xlsx --output-unit µL
  python batch_process.py my_dilutions.xlsx --output-unit mL --output my_results.xlsx

Excel Format:
  Use column headers with units: C1 (M), C2 (mM), V2 (mL)
  Cells should contain only numeric values.
        """
    )
    
    parser.add_argument(
        'input_file',
        help='Path to input Excel file (.xlsx)'
    )
    
    parser.add_argument(
        '--output-unit', '-u',
        default='mL',
        choices=['L', 'mL', 'µL', 'uL', 'nL'],
        help='Unit for output volumes V1 and Dilution (default: mL)'
    )
    
    parser.add_argument(
        '--output', '-o',
        default=None,
        help='Path to output Excel file (default: <input>_results.xlsx)'
    )
    
    args = parser.parse_args()
    
    try:
        print(f"Processing: {args.input_file}")
        print(f"Output unit: {args.output_unit}")
        
        # Process the dilutions
        df = process_excel_dilutions(
            input_file=args.input_file,
            output_file=args.output,
            output_unit=args.output_unit
        )
        
        # Display results
        print("\n" + "="*70)
        print("RESULTS:")
        print("="*70)
        print(df.to_string(index=False))
        print("="*70)
        
        # Check for errors
        errors = df[df['Error'].notna()]
        if not errors.empty:
            print("\n⚠️  WARNING: Some rows have errors:")
            print(errors[['Error']].to_string(index=True))
        else:
            print("\n✓ All dilutions calculated successfully!")
        
    except FileNotFoundError:
        print(f"ERROR: File not found: {args.input_file}")
        print(f"Make sure the file exists and the path is correct.")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
