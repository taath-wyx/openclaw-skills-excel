"""
Basic usage examples for Excel Handler
"""

import sys
from pathlib import Path
import pandas as pd

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.excel_handler import ExcelHandler


def example_1_write_dataframe():
    """Example 1: Write a simple DataFrame to Excel"""
    print("=" * 50)
    print("Example 1: Write DataFrame to Excel")
    print("=" * 50)
    
    # Create sample data
    data = {
        'Name': ['Alice', 'Bob', 'Charlie', 'David'],
        'Age': [25, 30, 35, 40],
        'Department': ['Sales', 'IT', 'HR', 'Finance'],
        'Salary': [50000, 70000, 65000, 75000]
    }
    df = pd.DataFrame(data)
    
    # Write to Excel
    handler = ExcelHandler()
    output_file = Path(__file__).parent / 'output_example1.xlsx'
    handler.write_excel(df, output_file, sheet_name='Employees')
    
    print(f"✓ DataFrame written to {output_file}")
    print(f"\nData preview:\n{df}")
    print()


def example_2_read_excel():
    """Example 2: Read Excel file into DataFrame"""
    print("=" * 50)
    print("Example 2: Read Excel to DataFrame")
    print("=" * 50)
    
    input_file = Path(__file__).parent / 'output_example1.xlsx'
    
    if input_file.exists():
        handler = ExcelHandler()
        df = handler.read_excel(input_file, sheet_name='Employees')
        
        print(f"✓ Data read from {input_file}")
        print(f"\nData shape: {df.shape}")
        print(f"\nData preview:\n{df}")
    else:
        print(f"! File not found. Please run example_1_write_dataframe first.")
    print()


def example_3_multiple_sheets():
    """Example 3: Write multiple sheets to one file"""
    print("=" * 50)
    print("Example 3: Multiple Sheets")
    print("=" * 50)
    
    # Create multiple DataFrames
    employees_data = {
        'ID': [1, 2, 3],
        'Name': ['Alice', 'Bob', 'Charlie'],
        'Department': ['Sales', 'IT', 'HR']
    }
    
    sales_data = {
        'Month': ['Jan', 'Feb', 'Mar'],
        'Revenue': [50000, 55000, 60000],
        'Target': [50000, 50000, 60000]
    }
    
    dfs = {
        'Employees': pd.DataFrame(employees_data),
        'Sales': pd.DataFrame(sales_data)
    }
    
    # Write to Excel
    handler = ExcelHandler()
    output_file = Path(__file__).parent / 'output_example3.xlsx'
    handler.write_excel(dfs, output_file)
    
    print(f"✓ Multiple sheets written to {output_file}")
    print(f"✓ Sheet names: {list(dfs.keys())}")
    print()


def example_4_read_all_sheets():
    """Example 4: Read all sheets from Excel file"""
    print("=" * 50)
    print("Example 4: Read All Sheets")
    print("=" * 50)
    
    input_file = Path(__file__).parent / 'output_example3.xlsx'
    
    if input_file.exists():
        handler = ExcelHandler()
        
        # Get sheet names
        sheet_names = handler.get_sheet_names(input_file)
        print(f"✓ Available sheets: {sheet_names}")
        
        # Read all sheets
        dfs = handler.read_excel(input_file, sheet_name=None)
        
        for sheet_name, df in dfs.items():
            print(f"\n--- Sheet: {sheet_name} ---")
            print(df)
    else:
        print(f"! File not found. Please run example_3_multiple_sheets first.")
    print()


def example_5_append_data():
    """Example 5: Append data to existing Excel file"""
    print("=" * 50)
    print("Example 5: Append Data")
    print("=" * 50)
    
    input_file = Path(__file__).parent / 'output_example1.xlsx'
    
    if input_file.exists():
        # Create new data to append
        new_data = {
            'Name': ['Eve', 'Frank'],
            'Age': [28, 33],
            'Department': ['Marketing', 'Operations'],
            'Salary': [55000, 70000]
        }
        new_df = pd.DataFrame(new_data)
        
        handler = ExcelHandler()
        handler.append_data(new_df, input_file, sheet_name='Employees')
        
        print(f"✓ Data appended to {input_file}")
        
        # Read and display updated data
        df = handler.read_excel(input_file, sheet_name='Employees')
        print(f"\nUpdated data (total {len(df)} rows):\n{df}")
    else:
        print(f"! File not found. Please run example_1_write_dataframe first.")
    print()


def example_6_filter_data():
    """Example 6: Filter Excel data"""
    print("=" * 50)
    print("Example 6: Filter Data")
    print("=" * 50)
    
    input_file = Path(__file__).parent / 'output_example1.xlsx'
    
    if input_file.exists():
        handler = ExcelHandler()
        
        # Filter by Department
        filters = {'Department': 'IT'}
        df_filtered = handler.filter_data(input_file, sheet_name='Employees', filters=filters)
        
        print(f"✓ Filtered data (Department='IT'):\n{df_filtered}")
    else:
        print(f"! File not found. Please run example_1_write_dataframe first.")
    print()


def example_7_style_worksheet():
    """Example 7: Apply styling to worksheet"""
    print("=" * 50)
    print("Example 7: Style Worksheet")
    print("=" * 50)
    
    input_file = Path(__file__).parent / 'output_example1.xlsx'
    
    if input_file.exists():
        handler = ExcelHandler()
        handler.style_worksheet(input_file, sheet_name='Employees', 
                               header_style=True, auto_width=True)
        
        print(f"✓ Styling applied to {input_file}")
        print("  - Header row formatted with blue background and white text")
        print("  - Column widths auto-fitted")
    else:
        print(f"! File not found. Please run example_1_write_dataframe first.")
    print()


if __name__ == '__main__':
    print("\n" + "=" * 50)
    print("EXCEL HANDLER - USAGE EXAMPLES")
    print("=" * 50 + "\n")
    
    example_1_write_dataframe()
    example_2_read_excel()
    example_3_multiple_sheets()
    example_4_read_all_sheets()
    example_5_append_data()
    example_6_filter_data()
    example_7_style_worksheet()
    
    print("=" * 50)
    print("All examples completed!")
    print("=" * 50)
