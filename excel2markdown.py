#!/usr/bin/env python3
import pandas as pd
import sys
import os
import argparse
from pathlib import Path

def excel_to_markdown(excel_file, output_file=None, sheet_name=None, index=False):
    """
    Convert an Excel file to Markdown format.
    
    Parameters:
    -----------
    excel_file : str
        Path to the Excel file
    output_file : str, optional
        Path to the output Markdown file. If None, output is printed to console.
    sheet_name : str or list, optional
        Sheet name(s) to convert. If None, all sheets are converted.
    index : bool, optional
        Whether to include index column in the output.
    """
    try:
        # Get all sheet names if not specified
        if sheet_name is None:
            xls = pd.ExcelFile(excel_file)
            sheet_names = xls.sheet_names
        else:
            sheet_names = [sheet_name] if isinstance(sheet_name, str) else sheet_name
        
        # Initialize output string
        md_content = ""
        
        # Process each sheet
        for sheet in sheet_names:
            # Read the Excel file
            df = pd.read_excel(excel_file, sheet_name=sheet)
            
            # Add sheet name as header if multiple sheets
            if len(sheet_names) > 1:
                md_content += f"## {sheet}\n\n"
            
            # Convert dataframe to markdown
            md_table = df.to_markdown(index=index)
            md_content += md_table + "\n\n"
        
        # Write to file or print to console
        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(md_content)
            print(f"Converted Excel file to Markdown: {output_file}")
        else:
            print(md_content)
            
    except Exception as e:
        print(f"Error converting Excel file: {e}", file=sys.stderr)
        return False
    
    return True

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Convert Excel files to Markdown format.')
    parser.add_argument('excel_file', help='Path to the Excel file')
    parser.add_argument('-o', '--output', help='Path to the output Markdown file')
    parser.add_argument('-s', '--sheet', help='Sheet name to convert. If not specified, all sheets are converted.')
    parser.add_argument('-i', '--index', action='store_true', help='Include index column in the output')
    
    args = parser.parse_args()
    
    # Check if input file exists
    if not os.path.exists(args.excel_file):
        print(f"Error: File '{args.excel_file}' not found.", file=sys.stderr)
        return 1
    
    # Generate output file name if not specified
    if not args.output:
        input_path = Path(args.excel_file)
        args.output = input_path.with_suffix('.md')
    
    # Convert Excel to Markdown
    success = excel_to_markdown(
        args.excel_file, 
        args.output, 
        args.sheet, 
        args.index
    )
    
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main()) 