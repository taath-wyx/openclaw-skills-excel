"""
Template and Format Conversion Examples
"""

import sys
from pathlib import Path
import pandas as pd
import json
import tempfile

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.excel_handler import ExcelHandler
from src.format_converter import DataFormatConverter
from src.template_engine import (
    ExcelTemplate, TemplateManager,
    create_employee_template,
    create_sales_template,
    create_inventory_template
)


def example_1_convert_csv_to_excel():
    """Example 1: Convert CSV to Excel using template"""
    print("=" * 60)
    print("Example 1: Convert CSV to Excel with Template")
    print("=" * 60)
    
    # Create sample CSV file
    csv_file = Path(__file__).parent / 'sample_employees.csv'
    csv_data = """id,name,dept,position,salary,start_date,email
1,Alice Johnson,Sales,Manager,70000,2020-01-15,alice@example.com
2,Bob Smith,IT,Developer,85000,2020-06-01,bob@example.com
3,Charlie Brown,HR,Specialist,65000,2021-03-10,charlie@example.com
4,David Wilson,Finance,Analyst,72000,2021-09-20,david@example.com"""
    
    csv_file.write_text(csv_data)
    print(f"✓ Created sample CSV: {csv_file}")
    
    # Read CSV data
    df = DataFormatConverter.read_csv(csv_file)
    print(f"✓ Read CSV with {len(df)} rows")
    
    # Use Employee template
    template = create_employee_template()
    print(f"✓ Using template: {template.name}")
    
    # Transform data according to template
    transformed_df = template.transform(df, validate=True)
    print(f"✓ Data transformed and validated")
    
    # Write to Excel using template
    handler = ExcelHandler()
    output_file = Path(__file__).parent / 'output_csv_to_excel.xlsx'
    handler.write_excel(transformed_df, output_file, sheet_name=template.sheet_name)
    handler.style_worksheet(output_file, sheet_name=template.sheet_name,
                           header_style=True, auto_width=True)
    
    print(f"✓ Excel file created: {output_file}")
    print(f"\nTransformed data:\n{transformed_df.head()}\n")


def example_2_convert_json_to_excel():
    """Example 2: Convert JSON to Excel"""
    print("=" * 60)
    print("Example 2: Convert JSON to Excel with Template")
    print("=" * 60)
    
    # Create sample JSON file
    json_file = Path(__file__).parent / 'sample_sales.json'
    json_data = [
        {
            "order_id": 1001,
            "customer_name": "Acme Corp",
            "product_name": "Widget A",
            "qty": 100,
            "unit_price": 25.50,
            "total": 2550.00,
            "order_date": "2024-01-15",
            "status": "Completed"
        },
        {
            "order_id": 1002,
            "customer_name": "TechCorp",
            "product_name": "Widget B",
            "qty": 50,
            "unit_price": 35.75,
            "total": 1787.50,
            "order_date": "2024-01-20",
            "status": "Pending"
        },
        {
            "order_id": 1003,
            "customer_name": "Global Inc",
            "product_name": "Widget C",
            "qty": 200,
            "unit_price": 15.00,
            "total": 3000.00,
            "order_date": "2024-02-01",
            "status": "Completed"
        }
    ]
    
    json_file.write_text(json.dumps(json_data, indent=2))
    print(f"✓ Created sample JSON: {json_file}")
    
    # Read JSON data
    df = DataFormatConverter.read_json(json_file)
    print(f"✓ Read JSON with {len(df)} records")
    
    # Use Sales template
    template = create_sales_template()
    print(f"✓ Using template: {template.name}")
    
    # Transform data
    transformed_df = template.transform(df, validate=True)
    print(f"✓ Data transformed and validated")
    
    # Write to Excel
    handler = ExcelHandler()
    output_file = Path(__file__).parent / 'output_json_to_excel.xlsx'
    handler.write_excel(transformed_df, output_file, sheet_name=template.sheet_name)
    handler.style_worksheet(output_file, sheet_name=template.sheet_name,
                           header_style=True, auto_width=True)
    
    print(f"✓ Excel file created: {output_file}")
    print(f"\nTransformed data:\n{transformed_df.head()}\n")


def example_3_custom_template():
    """Example 3: Create and use custom template"""
    print("=" * 60)
    print("Example 3: Custom Template Creation")
    print("=" * 60)
    
    # Create custom template for Product Catalog
    template = ExcelTemplate("Product Catalog")
    template.set_sheet_name("Products")
    template.add_column("Product ID", source_column="id", data_type="integer", required=True)
    template.add_column("Product Name", source_column="name", data_type="string", required=True)
    template.add_column("Category", source_column="category", data_type="string")
    template.add_column("Price", source_column="price", data_type="currency", required=True)
    template.add_column("Stock", source_column="stock", data_type="integer")
    template.add_column("Supplier", source_column="supplier", data_type="string")
    template.add_column("Description", source_column="desc", data_type="string")
    
    # Add validation rules
    template.add_validation("Price", lambda x: x is None or (isinstance(x, (int, float)) and x > 0),
                           "Price must be positive")
    template.add_validation("Stock", lambda x: x is None or (isinstance(x, int) and x >= 0),
                           "Stock must be non-negative")
    
    print(f"✓ Created custom template: {template.name}")
    print(f"✓ Columns: {', '.join(template.headers)}")
    
    # Create sample data
    sample_data = {
        'id': [1, 2, 3],
        'name': ['Premium Widget', 'Standard Widget', 'Economy Widget'],
        'category': ['Widgets', 'Widgets', 'Widgets'],
        'price': [99.99, 49.99, 29.99],
        'stock': [150, 500, 1000],
        'supplier': ['Supplier A', 'Supplier B', 'Supplier A'],
        'desc': ['High-end', 'Mid-range', 'Budget-friendly']
    }
    df = pd.DataFrame(sample_data)
    
    # Transform using template
    transformed_df = template.transform(df, validate=True)
    print(f"✓ Data transformed: {len(transformed_df)} rows")
    
    # Write to Excel
    handler = ExcelHandler()
    output_file = Path(__file__).parent / 'output_custom_template.xlsx'
    handler.write_excel(transformed_df, output_file, sheet_name=template.sheet_name)
    handler.style_worksheet(output_file, sheet_name=template.sheet_name,
                           header_style=True, auto_width=True)
    
    print(f"✓ Excel file created: {output_file}")
    print(f"\nTransformed data:\n{transformed_df}\n")


def example_4_multi_sheet_templates():
    """Example 4: Multiple sheets with different templates"""
    print("=" * 60)
    print("Example 4: Multiple Sheets with Different Templates")
    print("=" * 60)
    
    # Create sample data for different sheets
    employees_data = {
        'id': [1, 2, 3],
        'name': ['Alice', 'Bob', 'Charlie'],
        'dept': ['Sales', 'IT', 'HR'],
        'position': ['Manager', 'Developer', 'Specialist'],
        'salary': [70000, 85000, 65000],
        'start_date': ['2020-01-15', '2020-06-01', '2021-03-10'],
        'email': ['alice@example.com', 'bob@example.com', 'charlie@example.com']
    }
    
    sales_data = {
        'order_id': [1001, 1002],
        'customer_name': ['Acme', 'TechCorp'],
        'product_name': ['Widget A', 'Widget B'],
        'qty': [100, 50],
        'unit_price': [25.50, 35.75],
        'total': [2550, 1787.50],
        'order_date': ['2024-01-15', '2024-01-20'],
        'status': ['Completed', 'Pending']
    }
    
    # Transform with templates
    emp_template = create_employee_template()
    sales_template = create_sales_template()
    
    emp_df = emp_template.transform(pd.DataFrame(employees_data), validate=True)
    sales_df = sales_template.transform(pd.DataFrame(sales_data), validate=True)
    
    print(f"✓ Employees sheet: {len(emp_df)} rows")
    print(f"✓ Sales sheet: {len(sales_df)} rows")
    
    # Write multiple sheets
    data_sheets = {
        emp_template.sheet_name: emp_df,
        sales_template.sheet_name: sales_df
    }
    
    handler = ExcelHandler()
    output_file = Path(__file__).parent / 'output_multi_templates.xlsx'
    handler.write_excel(data_sheets, output_file)
    
    # Style each sheet
    handler.style_worksheet(output_file, sheet_name="Employees", header_style=True, auto_width=True)
    handler.style_worksheet(output_file, sheet_name="Sales", header_style=True, auto_width=True)
    
    print(f"✓ Excel file with {len(data_sheets)} sheets created: {output_file}\n")


def example_5_template_manager():
    """Example 5: Template Manager for managing multiple templates"""
    print("=" * 60)
    print("Example 5: Template Manager")
    print("=" * 60)
    
    # Create template manager
    manager = TemplateManager()
    
    # Add predefined templates
    emp_template = create_employee_template()
    sales_template = create_sales_template()
    inv_template = create_inventory_template()
    
    manager.templates['Employee'] = emp_template
    manager.templates['Sales'] = sales_template
    manager.templates['Inventory'] = inv_template
    
    print(f"✓ Templates loaded: {manager.list_templates()}")
    
    # Save template configuration
    temp_dir = Path(tempfile.gettempdir())
    template_file = temp_dir / 'employee_template.json'
    manager.save_template('Employee', str(template_file))
    print(f"✓ Template saved to: {template_file}")
    
    # Load template
    manager2 = TemplateManager()
    loaded = manager2.load_template(str(template_file))
    print(f"✓ Template loaded: {loaded.name}")
    print(f"✓ Columns: {', '.join(loaded.headers)}\n")


def example_6_auto_format_detection():
    """Example 6: Auto-detect file format and convert"""
    print("=" * 60)
    print("Example 6: Auto Format Detection")
    print("=" * 60)
    
    # Create sample files
    csv_file = Path(__file__).parent / 'auto_detect.csv'
    json_file = Path(__file__).parent / 'auto_detect.json'
    
    # Create CSV
    csv_data = """product_id,product_name,stock
1,Widget A,100
2,Widget B,200
3,Widget C,150"""
    csv_file.write_text(csv_data)
    
    # Create JSON
    json_data = [
        {"product_id": 1, "product_name": "Widget A", "stock": 100},
        {"product_id": 2, "product_name": "Widget B", "stock": 200}
    ]
    json_file.write_text(json.dumps(json_data))
    
    # Auto-detect and read
    print(f"✓ CSV format: {DataFormatConverter.detect_format(csv_file)}")
    print(f"✓ JSON format: {DataFormatConverter.detect_format(json_file)}")
    
    df_csv = DataFormatConverter.read_file(csv_file)
    df_json = DataFormatConverter.read_file(json_file)
    
    print(f"✓ CSV data read: {len(df_csv)} rows")
    print(f"✓ JSON data read: {len(df_json)} rows\n")


if __name__ == '__main__':
    print("\n" + "=" * 60)
    print("TEMPLATE & FORMAT CONVERSION EXAMPLES")
    print("=" * 60 + "\n")
    
    try:
        example_1_convert_csv_to_excel()
        example_2_convert_json_to_excel()
        example_3_custom_template()
        example_4_multi_sheet_templates()
        example_5_template_manager()
        example_6_auto_format_detection()
        
        print("=" * 60)
        print("✓ All examples completed successfully!")
        print("=" * 60)
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
