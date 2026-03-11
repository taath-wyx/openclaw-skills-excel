"""
OpenClaw格式转换器 - 精简版
支持CSV/JSON到Excel的转换
"""

from pathlib import Path
from typing import Union, Optional
import pandas as pd


class DataFormatConverter:
    """OpenClaw专用的格式转换器"""
    
    @staticmethod
    def read_csv(file_path: Union[str, Path], **kwargs) -> pd.DataFrame:
        """读取CSV文件"""
        try:
            return pd.read_csv(file_path, **kwargs)
        except Exception as e:
            raise ValueError(f"CSV读取失败: {str(e)}")
    
    @staticmethod
    def read_json(file_path: Union[str, Path], **kwargs) -> pd.DataFrame:
        """读取JSON文件"""
        try:
            return pd.read_json(file_path, **kwargs)
        except Exception as e:
            raise ValueError(f"JSON读取失败: {str(e)}")
    
    @staticmethod
    def read_file(file_path: Union[str, Path]) -> pd.DataFrame:
        """
        智能读取文件（自动检测格式）
        
        Args:
            file_path: 文件路径
            
        Returns:
            DataFrame
        """
        path = Path(file_path)
        suffix = path.suffix.lower()
        
        if suffix == '.csv':
            return DataFormatConverter.read_csv(file_path)
        elif suffix == '.json':
            return DataFormatConverter.read_json(file_path)
        elif suffix in ['.xlsx', '.xls']:
            return pd.read_excel(file_path)
        else:
            raise ValueError(f"不支持的文件格式: {suffix}")
    
    @staticmethod
    def convert_to_excel(input_file: Union[str, Path], 
                        output_file: Union[str, Path],
                        sheet_name: str = "Sheet1") -> None:
        """
        转换文件到Excel格式
        
        Args:
            input_file: 输入文件路径
            output_file: 输出Excel文件路径
            sheet_name: 工作表名称
        """
        df = DataFormatConverter.read_file(input_file)
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)