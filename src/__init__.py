"""
OpenClaw Excel Skills - 核心模块
专为OpenClaw设计的轻量级Excel处理技能
"""

__version__ = "2.1.0"
__author__ = "taath-wyx"
__description__ = "OpenClaw专用Excel文件读写技能"

# 核心导出 - 只保留OpenClaw必需的功能
from .excel_handler import ExcelHandler
from .format_converter import DataFormatConverter  
from .template_engine import ExcelTemplate, TemplateManager, create_employee_template, create_sales_template

__all__ = [
    "ExcelHandler",           # Excel文件读写
    "DataFormatConverter",    # 格式转换
    "ExcelTemplate",         # 模板引擎
    "TemplateManager",       # 模板管理
    "create_employee_template",  # 员工模板
    "create_sales_template",     # 销售模板
]