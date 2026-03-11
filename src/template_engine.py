"""
OpenClaw模板引擎 - 精简版
核心模板功能，去除复杂验证和高级特性
"""

from typing import Dict, List, Optional
import pandas as pd


class ExcelTemplate:
    """OpenClaw专用模板类"""
    
    def __init__(self, name: str):
        """初始化模板"""
        self.name = name
        self.columns = {}
        self.mappings = {}
        self.sheet_name = "Sheet1"
    
    def add_column(self, excel_column: str, source_column: Optional[str] = None,
                   data_type: str = "string") -> 'ExcelTemplate':
        """
        添加模板列
        
        Args:
            excel_column: Excel中的列名
            source_column: 源数据列名（如果不同）
            data_type: 数据类型
        """
        self.columns[excel_column] = {
            'data_type': data_type,
            'source_column': source_column or excel_column
        }
        self.mappings[source_column or excel_column] = excel_column
        return self
    
    def set_sheet_name(self, sheet_name: str) -> 'ExcelTemplate':
        """设置工作表名称"""
        self.sheet_name = sheet_name
        return self
    
    def transform(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        应用模板转换
        
        Args:
            df: 输入DataFrame
            
        Returns:
            转换后的DataFrame
        """
        # 简单列映射
        result_df = pd.DataFrame()
        
        for source_col, target_col in self.mappings.items():
            if source_col in df.columns:
                result_df[target_col] = df[source_col]
        
        return result_df


class TemplateManager:
    """OpenClaw模板管理器"""
    
    def __init__(self):
        self.templates = {}
    
    def create_template(self, name: str) -> ExcelTemplate:
        """创建新模板"""
        template = ExcelTemplate(name)
        self.templates[name] = template
        return template
    
    def get_template(self, name: str) -> ExcelTemplate:
        """获取模板"""
        return self.templates[name]


# OpenClaw预定义模板
def create_employee_template() -> ExcelTemplate:
    """员工数据模板"""
    template = ExcelTemplate("Employee")
    template.set_sheet_name("员工信息")
    template.add_column("ID", "id", "integer")
    template.add_column("姓名", "name", "string")
    template.add_column("部门", "dept", "string")
    template.add_column("职位", "position", "string")
    return template


def create_sales_template() -> ExcelTemplate:
    """销售数据模板"""
    template = ExcelTemplate("Sales")
    template.set_sheet_name("销售数据")
    template.add_column("订单号", "order_id", "integer")
    template.add_column("客户", "customer", "string")
    template.add_column("产品", "product", "string")
    template.add_column("数量", "quantity", "integer")
    template.add_column("单价", "unit_price", "float")
    return template