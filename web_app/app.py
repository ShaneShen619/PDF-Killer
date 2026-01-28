from flask import Flask, render_template, request, send_file, flash
import PyPDF2
import io
import os

app = Flask(__name__)
app.secret_key = 'super_secret_key' # 用于 flash 消息

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_pdfs():
    if 'pdfs' not in request.files:
        return "没有上传文件", 400
    
    files = request.files.getlist('pdfs')
    keyword = request.form.get('keyword', '').strip()
    
    if not keyword:
        return "关键词不能为空", 400

    if not files or files[0].filename == '':
        return "请选择至少一个 PDF 文件", 400

    # 内存中的输出缓冲区
    output_buffer = io.BytesIO()
    writer = PyPDF2.PdfWriter()
    
    total_found_pages = 0
    scanned_files_count = 0

    print(f"🔍 开始处理任务：关键词 '{keyword}'")

    for file in files:
        if not file.filename.lower().endswith('.pdf'):
            continue
            
        scanned_files_count += 1
        
        try:
            # 直接从内存读取，不需要保存到硬盘
            # file.stream 就像一个打开的文件对象
            reader = PyPDF2.PdfReader(file)
            
            for i, page in enumerate(reader.pages):
                text = page.extract_text()
                if text:
                    # 同样的清洗逻辑
                    clean_text = text.replace(" ", "").replace("\n", "")
                    if keyword in clean_text:
                        writer.add_page(page)
                        total_found_pages += 1
        
        except Exception as e:
            print(f"❌ 处理文件 {file.filename} 时出错: {e}")

    if total_found_pages > 0:
        writer.write(output_buffer)
        output_buffer.seek(0)
        
        print(f"✅ 成功! 提取了 {total_found_pages} 页。")
        
        return send_file(
            output_buffer,
            as_attachment=True,
            download_name=f"搜索结果_{keyword}.pdf",
            mimetype='application/pdf'
        )
    else:
        return """
        <div style="text-align:center; margin-top:50px; font-family:sans-serif;">
            <h1>⚠️ 未找到结果</h1>
            <p>在 {scanned_files_count} 个文件中未找到包含 "{keyword}" 的页面。</p>
            <a href="/">返回</a>
        </div>
        """

if __name__ == '__main__':
    print("🚀 服务器已启动！请在浏览器访问: http://127.0.0.1:5001")
    app.run(debug=True, port=5001)
