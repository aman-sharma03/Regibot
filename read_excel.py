#!/usr/bin/env python3
"""
Simple Excel Reader - Quick tool to read and display Excel files

Usage: python read_excel.py <filename.xlsx>
"""

import pandas as pd
import sys
import os


def read_excel_file(filename):
    """Read and display an Excel file"""
    
    if not os.path.exists(filename):
        print(f"❌ File '{filename}' not found!")
        return
    
    try:
        # Check if file has multiple sheets
        excel_file = pd.ExcelFile(filename)
        sheets = excel_file.sheet_names
        
        print(f"📊 Reading Excel file: {filename}")
        print(f"📋 Available sheets: {sheets}")
        print("=" * 60)
        
        # Read each sheet
        for i, sheet_name in enumerate(sheets):
            print(f"\n📄 Sheet {i+1}: '{sheet_name}'")
            print("-" * 40)
            
            df = pd.read_excel(filename, sheet_name=sheet_name)
            
            # Display basic info
            print(f"Shape: {df.shape} (rows, columns)")
            print(f"Columns: {list(df.columns)}")
            
            # Show first few rows
            print(f"\nFirst 5 rows:")
            print(df.head())
            
            # Show data types
            print(f"\nData types:")
            for col, dtype in df.dtypes.items():
                print(f"  {col}: {dtype}")
            
            # Check for missing values
            missing = df.isnull().sum()
            if missing.any():
                print(f"\nMissing values:")
                for col, count in missing[missing > 0].items():
                    print(f"  {col}: {count}")
            else:
                print("\n✅ No missing values")
            
            # Basic statistics for numeric columns
            numeric_cols = df.select_dtypes(include=['number']).columns
            if len(numeric_cols) > 0:
                print(f"\nNumeric summary:")
                print(df[numeric_cols].describe())
            
            if i < len(sheets) - 1:  # Add separator between sheets
                print("\n" + "="*60)
        
        print(f"\n✅ Successfully read {len(sheets)} sheet(s) from '{filename}'")
        
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")


def main():
    """Main function"""
    if len(sys.argv) != 2:
        print("Usage: python read_excel.py <filename.xlsx>")
        print("\nExample: python read_excel.py sample_data.xlsx")
        
        # Show available Excel files in current directory
        excel_files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls'))]
        if excel_files:
            print(f"\nAvailable Excel files in current directory:")
            for file in excel_files:
                print(f"  - {file}")
        return
    
    filename = sys.argv[1]
    read_excel_file(filename)


if __name__ == "__main__":
    main()