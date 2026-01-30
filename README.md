# PDF-Killer 📄🔍

一个强大的办公自动化工具箱，支持 PDF 关键词提取合并，以及 Excel 员工工时自动统计。

## ✨ 功能特性

### 📄 PDF 处理
- 🔍 **关键词搜索** - 在多个 PDF 文件中搜索指定文本
- 📑 **页面提取** - 自动提取包含关键词的页面
- 📦 **智能合并** - 将所有匹配页面合并为一个 PDF

### 📊 Excel 工时统计 (新功能!)
- ⏱️ **工时汇总** - 自动扫描文件夹内所有 Excel 考勤表
- 🤖 **智能识别** - 自动识别员工姓名、工号和工作时长
- 📈 **批量计算** - 一键统计所有员工的月度总工时
- 🧹 **数据清洗** - 自动过滤无关文件，兼容多种表头格式

## 📁 项目结构

```
PDF-Killer/
├── excel_files/          # 存放 Excel 考勤表数据 (需自行创建)
├── searchpdf.py          # PDF 搜索脚本
├── sum_all_hours.py      # Excel 全局工时统计脚本 (主推)
├── search_excel_sum.py   # Excel 单个员工搜索脚本
├── web_app/              # Web 版本 (PDF功能)
├── requirements.txt      # Python 依赖
└── README.md
```

## 🚀 快速开始

### 1. 安装依赖

确保你安装了 Python 3，然后运行：

```bash
pip3 install -r requirements.txt
```

### 2. 使用 Excel 工时统计功能 📊

**场景**：你需要计算一个月内所有员工的工时总和。

1. **准备数据**：将所有 Excel 考勤表 (`.xlsx`) 放入 `PDF-Killer/excel_files/` 文件夹中。
2. **运行脚本**：

```bash
# 在项目根目录下运行
python3 PDF-Killer/sum_all_hours.py
```

3. **查看结果**：
终端将直接打印每个文件的员工姓名、工号及该表的工时合计，并在最后显示总工时。

**注意**：
- 脚本会自动寻找 `Arbeitsstunden` 或 `Arbeitszeit` 列。
- 脚本会读取 B3 单元格提取 `Nr.XX Name` 格式的员工信息。
- 只计算第 10 行到第 40 行（共31天）的数据，避免重复计算底部汇总。

---

### 3. 使用 PDF 搜索功能 📄

**场景**：你想把所有包含 "合同" 字样的 PDF 页面提取出来合并成一个新文件。

1. 打开 `searchpdf.py`，修改底部的 `source_folder_path` 和 `search_keyword`。
2. 运行脚本：

```bash
python3 searchpdf.py
```
3. 结果文件会生成在你的桌面。

### 4. Web 应用 (仅限 PDF 功能)

```bash
cd web_app
python3 app.py
```
访问 `http://127.0.0.1:5001`。

## 🛠️ 技术栈

- **Python 3.x**
- **pandas & openpyxl** - Excel 数据处理
- **PyPDF2** - PDF 处理
- **Flask** - Web 框架

## 📝 注意事项

- **Excel 格式**：工时统计脚本默认适配特定格式的考勤表（表头在第9行，数据从第10行开始，员工信息在B3）。
- **PDF 格式**：仅支持文本型 PDF（图片扫描件无法搜索）。

## 📄 License

MIT License