# PDF-Killer 📄🔍

一个强大的 PDF 关键词搜索与提取工具，可以在多个 PDF 文件中快速搜索指定关键词，并将包含该关键词的页面合并到一个新的 PDF 文件中。

## ✨ 功能特性

- 🔍 **关键词搜索** - 在多个 PDF 文件中搜索指定文本
- 📑 **页面提取** - 自动提取包含关键词的页面
- 📦 **智能合并** - 将所有匹配页面合并为一个 PDF
- 🖥️ **命令行工具** - 本地批量处理 PDF 文件
- 🌐 **Web 应用** - 浏览器在线使用，无需安装

## 📁 项目结构

```
PDF-Killer/
├── searchpdf.py          # 命令行脚本
├── web_app/
│   ├── app.py            # Flask Web 应用
│   └── templates/
│       └── index.html    # 前端页面
├── requirements.txt      # Python 依赖
├── vercel.json          # Vercel 部署配置
└── README.md
```

## 🚀 快速开始

### 安装依赖

```bash
pip3 install -r requirements.txt
```

### 方式一：命令行使用

1. 打开 `searchpdf.py`
2. 修改配置区的参数：

```python
# 文件夹路径（包含要搜索的 PDF 文件）
source_folder_path = '/path/to/your/pdf/folder'

# 要搜索的关键词
search_keyword = '你的关键词'
```

3. 运行脚本：

```bash
python3 searchpdf.py
```

4. 搜索结果将保存到桌面：`汇总搜索结果_关键词.pdf`

### 方式二：Web 应用使用

1. 启动服务器：

```bash
cd web_app
python3 app.py
```

2. 打开浏览器访问：`http://127.0.0.1:5001`

3. 上传 PDF 文件，输入关键词，点击提取

## ☁️ 部署到 Vercel

本项目支持一键部署到 Vercel：

1. Fork 本仓库
2. 在 Vercel 中导入项目
3. 自动部署完成

## 🛠️ 技术栈

- **Python 3.x**
- **PyPDF2** - PDF 读取与操作
- **Flask** - Web 框架
- **Vercel** - 云部署平台

## 📝 注意事项

- 仅支持文本型 PDF（扫描件需先 OCR）
- 搜索时会忽略空格和换行，提高匹配率
- Web 版本建议单次上传文件不超过 16MB

## 📄 License

MIT License
