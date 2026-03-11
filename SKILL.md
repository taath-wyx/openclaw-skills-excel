---
name: openclaw-skills-excel
description: 一个为 OpenClaw 开发的 Excel 文件读写技能，提供方便的 API 来读取和写入 Excel 文件，支持多种格式导入和模板转换。
read_when:
  - 处理 Excel 文件
  - 数据格式转换
  - 模板化数据处理
metadata: 
  openclaw:
    emoji: 📊
    requires:
      bins: ["python", "pip"]
      env: []
---

# Excel Handler for OpenClaw

一个为 OpenClaw 开发的 Excel 文件读写技能，提供方便的 API 来读取和写入 Excel 文件，支持多种格式导入和模板转换。

## 功能特性

✨ **核心功能**
- 📖 读取 Excel 文件到 Pandas DataFrame
- 💾 写入 DataFrame 到 Excel 文件
- 📋 支持多工作表操作
- 🎨 工作表样式设置（颜色、字体、列宽自适应）
- 📊 数据过滤和追加
- 🔧 工作表管理（删除、重命名）

🌟 **高级功能**
- 📁 多格式支持：CSV、JSON、XML 自动转换
- 🎯 模板引擎：定义和应用数据模板
- ✅ 数据验证：自动检查数据完整性和正确性
- 🔀 列映射：灵活的字段映射和转换
- 📦 预定义模板：员工、销售、库存等
- 💾 模板管理：保存/加载模版配置

## 安装依赖

```bash
pip install -r requirements.txt
```

## 快速开始

### 基本用法 - Excel 读写

#### 1. 写入 Excel 文件

```python
from src.excel_handler import ExcelHandler
import pandas as pd

# 创建数据
data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'Department': ['Sales', 'IT', 'HR']
}
df = pd.DataFrame(data)

# 写入文件
handler = ExcelHandler()
handler.write_excel(df, 'output.xlsx', sheet_name='Employees')
```

#### 2. 读取 Excel 文件

```python
from src.excel_handler import ExcelHandler

handler = ExcelHandler()
# 读取指定工作表
df = handler.read_excel('data.xlsx', sheet_name='Employees')
print(df)

# 读取所有工作表
all_sheets = handler.read_excel('data.xlsx', sheet_name=None)
```

### 高级用法 - 模板转换

#### 3. CSV 转 Excel（使用模板）

```python
from src.format_converter import DataFormatConverter
from src.template_engine import create_employee_template
from src.excel_handler import ExcelHandler

# 读取 CSV 文件
df = DataFormatConverter.read_csv('employees.csv')

# 使用预定义员工模板
template = create_employee_template()

# 转换数据（自动验证）
transformed_df = template.transform(df, validate=True)

# 写入 Excel
handler = ExcelHandler()
handler.write_excel(transformed_df, 'output.xlsx', 
                   sheet_name=template.sheet_name)
handler.style_worksheet('output.xlsx')
```

## 使用场景

- 数据导入导出
- 报表生成
- 数据格式转换
- 模板化数据处理
- 批量数据操作

## 注意事项

- 需要安装 Python 和 pip
- 依赖 pandas 和 openpyxl 库
- 支持 Excel 2007+ 格式 (.xlsx)
- 大文件建议分块处理