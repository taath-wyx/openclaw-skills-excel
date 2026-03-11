---
name: openclaw-skills-excel
version: 2.1.0
description: OpenClaw专用Excel文件读写技能 - 轻量级、高效、易安装
author: taath-wyx
license: MIT

# OpenClaw Excel Skills

专为OpenClaw优化的轻量级Excel处理技能，支持Excel文件读写、格式转换和模板化数据处理。

## 🚀 核心功能

- **Excel文件读写** - 基于pandas，支持.xlsx格式
- **智能模板系统** - 自动数据映射和标准化
- **格式转换** - CSV/JSON → Excel 一键转换
- **自动样式** - 表头美化，列宽自适应

## ⚡ OpenClaw集成

```bash
# 安装技能
openclaw skill install openclaw-skills-excel

# 使用技能
openclaw excel read data.xlsx
openclaw excel convert data.csv output.xlsx
```

## 📖 Python API

```python
from src.excel_handler import ExcelHandler
from src.template_engine import create_employee_template

# 读取Excel
handler = ExcelHandler()
df = handler.read_excel('data.xlsx')

# 使用模板
template = create_employee_template()
df_clean = template.transform(df)

# 写入Excel
handler.write_excel(df_clean, 'output.xlsx')
```

## 🔧 使用场景

- 数据报表自动生成
- 格式转换自动化
- 批量Excel处理
- 数据标准化

## 📁 项目结构

```
src/
├── __init__.py           # 包初始化
├── excel_handler.py      # Excel读写核心
├── format_converter.py   # 格式转换
└── template_engine.py    # 模板引擎

examples/
├── basic_usage.py        # 基础示例
└── template_examples.py  # 模板示例
```

## 🎯 特色功能

- **零配置** - 开箱即用
- **轻量级** - 仅依赖pandas和openpyxl
- **智能检测** - 自动识别文件格式
- **模板驱动** - 标准化数据处理
- **错误处理** - 完善的异常捕获

## 📝 依赖要求

- pandas >= 2.0.0
- openpyxl >= 3.1.5

## 🔍 版本历史

- v2.1.0: 精简优化，专为OpenClaw定制
- v2.0.0: 初始发布，完整功能

---
**许可证**: MIT  
**作者**: taath-wyx