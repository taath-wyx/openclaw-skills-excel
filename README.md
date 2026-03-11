# OpenClaw Excel Skills

一个功能强大的Excel文件读写技能，专为OpenClaw设计。支持多种格式转换、智能模板引擎和自动化样式处理。

## ✨ 主要特性

- 📊 **Excel文件读写** - 基于pandas和openpyxl
- 🔄 **多格式支持** - CSV、JSON、XML自动转换
- 🎯 **模板引擎** - 智能数据映射和验证
- 🎨 **样式自动化** - 自动列宽、表头样式
- 📈 **批量处理** - 多工作表、大数据量支持

## 🛠️ 安装

```bash
# 克隆仓库
git clone https://github.com/yourusername/openclaw-skills-excel.git

# 安装依赖
pip install -r requirements.txt
```

## 📖 快速开始

```python
from src.excel_handler import ExcelHandler
from src.template_engine import create_employee_template

# 读取Excel
handler = ExcelHandler()
df = handler.read_excel('data.xlsx')

# 使用模板
emp_template = create_employee_template()
transformed_df = emp_template.transform(df)

# 写入Excel
handler.write_excel(transformed_df, 'output.xlsx')
```

## 🎯 使用场景

- 数据报表生成
- 格式转换 (CSV/JSON/XML → Excel)
- 数据标准化处理
- 批量Excel操作

## 📁 项目结构

```
openclaw-skills-excel/
├── src/                    # 核心源码
│   ├── excel_handler.py    # Excel处理引擎
│   ├── format_converter.py # 格式转换器
│   └── template_engine.py  # 模板引擎
├── examples/               # 使用示例
├── tests/                  # 单元测试
├── requirements.txt        # 依赖管理
└── README.md              # 项目文档
```

## 🤝 贡献

欢迎提交Issue和Pull Request！

## 📄 许可证

MIT License - 详见 [LICENSE](LICENSE) 文件

## 👨‍💻 作者

taath-wyx

---

如果这个项目对您有帮助，请给个⭐️支持！