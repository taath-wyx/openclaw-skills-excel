# OpenClaw Excel Skills - 精简版

专为OpenClaw优化的Excel文件处理技能，简洁高效，易于安装和使用。

## 🎯 核心功能

- **Excel文件读写** - 基于pandas，支持.xlsx格式
- **智能模板系统** - 自动数据映射和验证
- **格式转换** - CSV/JSON → Excel
- **自动样式** - 表头美化，列宽自适应

## ⚡ 快速安装

```bash
# 安装依赖（OpenClaw会自动处理）
pip install pandas openpyxl

# 技能安装
openclaw skill install openclaw-skills-excel
```

## 📖 基础用法

```python
# 读取Excel
from src.excel_handler import ExcelHandler
handler = ExcelHandler()
df = handler.read_excel('data.xlsx')

# 写入Excel  
handler.write_excel(df, 'output.xlsx')

# 使用模板
template = create_employee_template()
df_clean = template.transform(df)
```

## 🔧 OpenClaw集成

在OpenClaw中可以直接使用：

```bash
# 读取Excel文件
openclaw excel read data.xlsx

# 转换格式
openclaw excel convert data.csv output.xlsx

# 应用模板
openclaw excel template employees.csv
```

## 📁 文件结构

```
src/
├── excel_handler.py      # Excel读写核心
├── format_converter.py   # 格式转换
└── template_engine.py    # 模板引擎

examples/
├── basic_usage.py        # 基础示例
└── template_examples.py  # 模板示例
```

## 🎨 特色功能

- **零配置** - 开箱即用
- **智能检测** - 自动识别文件格式
- **模板驱动** - 标准化数据处理
- **错误处理** - 完善的异常捕获

## 🔍 使用场景

- 数据报表生成
- 格式转换自动化  
- 批量Excel处理
- 数据标准化

---
**作者**: taath-wyx  
**许可证**: MIT  
**版本**: 2.0.0