import PyPDF2
import os

def extract_pages_by_keyword(pdf_path, keyword):
    # 1. 检查文件是否存在
    if not os.path.exists(pdf_path):
        print("❌ 错误：找不到文件，请检查路径是否正确。")
        return

    # 2. 准备输出文件名 (例如：原文件名_提取_张三.pdf)
    dir_name = os.path.dirname(pdf_path)
    base_name = os.path.basename(pdf_path)
    file_name_no_ext = os.path.splitext(base_name)[0]
    output_path = os.path.join(dir_name, f"{file_name_no_ext}_提取_{keyword}.pdf")
    

    try:
        # 3. 读取 PDF
        reader = PyPDF2.PdfReader(pdf_path)
        writer = PyPDF2.PdfWriter()
        found_pages = []

        print(f"🔍 正在搜索 '{keyword}' (共 {len(reader.pages)} 页)...")

        # 4. 遍历每一页
        for i, page in enumerate(reader.pages):
            # 提取文字
            text = page.extract_text()
            
            if text:
                # 关键步骤：移除所有空格和换行，提高中文匹配率
                # 比如 PDF 里是 "张   三"，移除空格后变成 "张三"，就能匹配到了
                clean_text = text.replace(" ", "").replace("\n", "")
                
                if keyword in clean_text:
                    writer.add_page(page)
                    found_pages.append(i + 1) # 记录页码 (从1开始)
                    print(f"   ✅ 第 {i + 1} 页已匹配")

        # 5. 保存结果
        if found_pages:
            with open(output_path, "wb") as f:
                writer.write(f)
            print("-" * 30)
            print(f"🎉 成功！共提取了 {len(found_pages)} 页。")
            print(f"📄 页码: {found_pages}")
            print(f"💾 文件已保存在: {output_path}")
        else:
            print(f"⚠️ 未找到包含 '{keyword}' 的页面。")

    except Exception as e:
        print(f"❌发生错误: {e}")

# ==========================================
# 👇 这里是配置区，只需要修改这里
# ==========================================

# 1. 这里输入你的 PDF 路径 (Mac 上可以直接把文件拖进代码编辑器获取路径)
source_pdf_path = '/Users/shane/Desktop/未命名文件夹/111.pdf' 

# 2. 这里输入你要搜索的关键词
search_keyword = 'Shishuai'

# 运行函数
extract_pages_by_keyword(source_pdf_path, search_keyword)