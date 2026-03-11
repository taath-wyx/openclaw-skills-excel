# OpenClaw Excel Skills - 安装指南

## 🚀 一键安装

```bash
# 安装技能（OpenClaw会自动处理依赖）
openclaw skill install openclaw-skills-excel

# 验证安装
openclaw skill list | grep excel
```

## 📖 快速开始

### 1. 读取Excel文件
```python
# 在OpenClaw中直接使用
from src.excel_handler import ExcelHandler

handler = ExcelHandler()
df = handler.read_excel('data.xlsx')
print(f"读取了 {len(df)} 行数据")
```

### 2. 格式转换
```python
# CSV转Excel
from src.format_converter import DataFormatConverter

converter = DataFormatConverter()
converter.convert_to_excel('data.csv', 'output.xlsx')
```

### 3. 使用模板
```python
# 应用员工模板
from src.template_engine import create_employee_template

template = create_employee_template()
df_clean = template.transform(df)
```

## 🔧 OpenClaw命令行

```bash
# 读取Excel
openclaw excel read data.xlsx

# 转换格式
openclaw excel convert data.csv output.xlsx

# 应用模板
openclaw excel template employees.csv
```

## 📊 使用示例

### 示例1：读取Excel数据
```python
# 读取Excel文件
df = handler.read_excel('销售数据.xlsx')

# 查看工作表
sheets = handler.get_sheet_names('销售数据.xlsx')
print(f"工作表: {sheets}")
```

### 示例2：数据过滤
```python
# 过滤数据
filtered_df = handler.filter_data('数据.xlsx', {'部门': '销售部'})
```

### 示例3：模板转换
```python
# 使用销售模板
template = create_sales_template()
df_transformed = template.transform(df)

# 写入新Excel
handler.write_excel(df_transformed, '标准化销售数据.xlsx')
```

## 🎯 核心优势

- **轻量级** - 仅2个依赖：pandas + openpyxl
- **零配置** - 开箱即用
- **智能** - 自动格式检测
- **安全** - 完善的错误处理
- **标准** - 遵循OpenClaw规范

## 🔍 故障排除

### 常见问题

**Q: 安装失败？**
A: 确保网络连接正常，OpenClaw会自动安装依赖

**Q: Excel文件打不开？**
A: 检查文件路径是否正确，文件是否损坏

**Q: 模板转换失败？**
A: 确保源数据列名与模板匹配

### 调试模式
```python
import logging
logging.basicConfig(level=logging.DEBUG)

handler = ExcelHandler()
# 现在可以看到详细的调试信息
```

## 📈 高级用法

### 自定义模板
```python
from src.template_engine import ExcelTemplate

template = ExcelTemplate("自定义模板")
template.add_column("名称", "name", "string")
template.add_column("数量", "quantity", "integer")
template.set_sheet_name("自定义数据")

result = template.transform(df)
```

### 批量处理
```python
import glob

for file in glob.glob("*.csv"):
    converter.convert_to_excel(file, file.replace('.csv', '.xlsx'))
```

---
**版本**: 2.1.0  
**许可证**: MIT  
**作者**: taath-wyx