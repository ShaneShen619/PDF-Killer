from flask import Flask, render_template, request, send_file
import PyPDF2
import pandas as pd
from xhtml2pdf import pisa
import io
import os

app = Flask(__name__)
app.secret_key = 'super_secret_key'

@app.route('/')
def index():
    return render_template('index.html')

def convert_html_to_pdf(source_html):
    """将 HTML 字符串转换为 PDF 字节流"""
    output = io.BytesIO()
    pisa_status = pisa.CreatePDF(source_html, dest=output)
    if pisa_status.err:
        return None
    output.seek(0)
    return output

@app.route('/process', methods=['POST'])
def process_files():
    if 'pdfs' not in request.files:
        return "没有上传文件", 400
    
    files = request.files.getlist('pdfs')
    keyword = request.form.get('keyword', '').strip()
    
    if not keyword:
        return "关键词不能为空", 400

    output_buffer = io.BytesIO()
    writer = PyPDF2.PdfWriter()
    
    total_found_items = 0
    scanned_files_count = 0
    keyword_lower = keyword.lower()

    print(f"🔍 开始处理任务：关键词 '{keyword}'")

    for file in files:
        filename = file.filename.lower()
        if not (filename.endswith('.pdf') or filename.endswith('.xlsx') or filename.endswith('.xls')):
            continue
            
        scanned_files_count += 1
        print(f"Processing: {file.filename}")
        
        try:
            # === 处理 PDF 文件 ===
            if filename.endswith('.pdf'):
                reader = PyPDF2.PdfReader(file)
                for i, page in enumerate(reader.pages):
                    text = page.extract_text()
                    if text:
                        clean_text = text.replace(" ", "").replace("\n", "").lower()
                        if keyword_lower in clean_text:
                            writer.add_page(page)
                            total_found_items += 1
                            print(f"  -> Found in PDF page {i+1}")

            # === 处理 Excel 文件 ===
            elif filename.endswith('.xlsx') or filename.endswith('.xls'):
                # 读取所有工作表
                # sheet_name=None 表示读取所有 sheet，返回一个字典 {sheet_name: dataframe}
                xls_data = pd.read_excel(file, sheet_name=None)
                
                for sheet_name, df in xls_data.items():
                    # 将整个 DataFrame 转为字符串进行搜索
                    # 这是一个比较粗暴但有效的全文搜索方式
                    df_str = df.to_string().lower()
                    
                    if keyword_lower in df_str:
                        print(f"  -> Found in Excel Sheet: {sheet_name}")
                        
                        # 构建美观一点的 HTML 表格
                        # 加上 style 确保表格有边框，且支持中文显示（依赖系统字体，但在纯英文环境可能回退）
                        html_content = f"""
                        <html>
                        <head>
                            <style>
                                @page {{ size: A4 landscape; margin: 1cm; }}
                                body {{ font-family: sans-serif; }}
                                h2 {{ color: #333; }}
                                table {{ width: 100%; border-collapse: collapse; font-size: 10px; }}
                                th, td {{ border: 1px solid #999; padding: 4px; text-align: left; }}
                                th {{ background-color: #f2f2f2; }}
                            </style>
                        </head>
                        <body>
                            <h2>文件: {file.filename} - 工作表: {sheet_name}</h2>
                            <p>关键词: {keyword}</p>
                            {df.to_html(index=False, na_rep='')}
                        </body>
                        </html>
                        """
                        
                        # 转换 HTML -> PDF 流
                        pdf_stream = convert_html_to_pdf(html_content)
                        
                        if pdf_stream:
                            # 将生成的 PDF 流作为新页面加入主 writer
                            temp_reader = PyPDF2.PdfReader(pdf_stream)
                            for page in temp_reader.pages:
                                writer.add_page(page)
                            total_found_items += 1

        except Exception as e:
            print(f"❌ 处理文件 {file.filename} 时出错: {e}")

    # === 结束循环，返回结果 ===
    if total_found_items > 0:
        writer.write(output_buffer)
        output_buffer.seek(0)
        
        print(f"✅ 成功! 共找到 {total_found_items} 处匹配 (PDF页面或Excel工作表)。")
        
        return send_file(
            output_buffer,
            as_attachment=True,
            download_name=f"搜索结果_{keyword}.pdf",
            mimetype='application/pdf'
        )
    else:
        return f"""
        <div style="text-align:center; margin-top:50px; font-family:sans-serif;">
            <h1>⚠️ 未找到结果</h1>
            <p>在 {scanned_files_count} 个文件中未找到包含 "{keyword}" 的内容。</p>
            <p><a href="/">返回</a></p>
        </div>
        """

if __name__ == '__main__':
    print("🚀 服务器已启动！请在浏览器访问: http://127.0.0.1:5001")
    app.run(debug=True, port=5001)