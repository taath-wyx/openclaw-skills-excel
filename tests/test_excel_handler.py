"""
Unit tests for Excel Handler
"""

import sys
import unittest
from pathlib import Path
import tempfile
import shutil
import pandas as pd

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.excel_handler import ExcelHandler


class TestExcelHandler(unittest.TestCase):
    """Test cases for ExcelHandler"""
    
    @classmethod
    def setUpClass(cls):
        """Set up test fixtures"""
        cls.temp_dir = tempfile.mkdtemp()
        cls.test_file = Path(cls.temp_dir) / 'test.xlsx'
        
        # Create sample data
        cls.sample_data = pd.DataFrame({
            'Name': ['Alice', 'Bob', 'Charlie'],
            'Age': [25, 30, 35],
            'City': ['New York', 'London', 'Paris']
        })
    
    @classmethod
    def tearDownClass(cls):
        """Clean up test fixtures"""
        shutil.rmtree(cls.temp_dir)
    
    def test_write_dataframe(self):
        """Test writing DataFrame to Excel"""
        handler = ExcelHandler()
        handler.write_excel(self.sample_data, self.test_file, sheet_name='Data')
        
        self.assertTrue(self.test_file.exists())
    
    def test_read_dataframe(self):
        """Test reading DataFrame from Excel"""
        handler = ExcelHandler()
        df = handler.read_excel(self.test_file, sheet_name='Data')
        
        self.assertEqual(len(df), 3)
        self.assertIn('Name', df.columns)
        self.assertEqual(df.iloc[0]['Name'], 'Alice')
    
    def test_write_multiple_sheets(self):
        """Test writing multiple sheets"""
        handler = ExcelHandler()
        multi_file = Path(self.temp_dir) / 'multi_sheet.xlsx'
        
        data = {
            'Sheet1': self.sample_data,
            'Sheet2': self.sample_data.copy()
        }
        
        handler.write_excel(data, multi_file)
        self.assertTrue(multi_file.exists())
        
        # Verify sheets
        sheet_names = handler.get_sheet_names(multi_file)
        self.assertIn('Sheet1', sheet_names)
        self.assertIn('Sheet2', sheet_names)
    
    def test_get_sheet_names(self):
        """Test getting sheet names"""
        handler = ExcelHandler()
        sheet_names = handler.get_sheet_names(self.test_file)
        
        self.assertIsInstance(sheet_names, list)
        self.assertIn('Data', sheet_names)
    
    def test_append_data(self):
        """Test appending data to Excel"""
        handler = ExcelHandler()
        append_file = Path(self.temp_dir) / 'append.xlsx'
        
        # Write initial data
        handler.write_excel(self.sample_data, append_file, sheet_name='Data')
        
        # Append new data
        new_data = pd.DataFrame({
            'Name': ['David'],
            'Age': [40],
            'City': ['Tokyo']
        })
        
        handler.append_data(new_data, append_file, sheet_name='Data')
        
        # Read and verify
        df = handler.read_excel(append_file, sheet_name='Data')
        self.assertEqual(len(df), 4)
    
    def test_filter_data(self):
        """Test filtering data"""
        handler = ExcelHandler()
        filter_file = Path(self.temp_dir) / 'filter.xlsx'
        
        handler.write_excel(self.sample_data, filter_file, sheet_name='Data')
        
        # Filter by Age > 25
        df = handler.read_excel(filter_file, sheet_name='Data')
        df_filtered = df[df['Age'] > 25]
        
        self.assertEqual(len(df_filtered), 2)
    
    def test_file_not_found(self):
        """Test handling of non-existent file"""
        handler = ExcelHandler()
        non_existent = Path(self.temp_dir) / 'non_existent.xlsx'
        
        with self.assertRaises(FileNotFoundError):
            handler.read_excel(non_existent)
    
    def test_invalid_file_path(self):
        """Test handling of invalid file path"""
        handler = ExcelHandler()
        
        with self.assertRaises(ValueError):
            handler.write_excel(self.sample_data, file_path=None)


if __name__ == '__main__':
    unittest.main()
