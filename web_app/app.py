from flask import Flask, render_template, request, send_file
import PyPDF2
import io
import os

app = Flask(__name__)
app.secret_key = 'super_secret_key'

@app.route('/')
def index():
    return render_template('index.html')

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
        if not filename.endswith('.pdf'):
            continue
            
        scanned_files_count += 1
        print(f"Processing: {file.filename}")
        
        try:
            # === 处理 PDF 文件 ===
            reader = PyPDF2.PdfReader(file)
            for i, page in enumerate(reader.pages):
                text = page.extract_text()
                if text:
                    clean_text = text.replace(" ", "").replace("\n", "").lower()
                    if keyword_lower in clean_text:
                        writer.add_page(page)
                        total_found_items += 1
                        print(f"  -> Found in PDF page {i+1}")

        except Exception as e:
            print(f"❌ 处理文件 {file.filename} 时出错: {e}")

    # === 结束循环，返回结果 ===
    if total_found_items > 0:
        writer.write(output_buffer)
        output_buffer.seek(0)
        
        print(f"✅ 成功! 共找到 {total_found_items} 处匹配 (PDF页面)。")
        
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
