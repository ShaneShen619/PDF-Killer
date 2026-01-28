import PyPDF2
import os

def search_and_merge_from_folder(folder_path, keyword):
    # 1. 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print("❌ 错误：找不到文件夹，请检查路径是否正确。")
        return

    # 准备输出文件路径 (修改为桌面)
    output_filename = f"汇总搜索结果_{keyword}.pdf"
    desktop_path = "/Users/shane/Desktop"
    output_path = os.path.join(desktop_path, output_filename)
    
    # 初始化写入器 (用于合并所有结果)
    writer = PyPDF2.PdfWriter()
    total_found_pages = 0
    scanned_files_count = 0

    # 2. 获取文件夹内所有 PDF 文件
    # 过滤出 .pdf 结尾的文件，并按文件名排序
    all_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    all_files.sort() # 排序，保证按顺序处理

    if not all_files:
        print(f"⚠️ 在 '{folder_path}' 中未找到任何 PDF 文件。")
        return

    print(f"📂 准备在 {len(all_files)} 个文件中搜索 '{keyword}'...\n")

    # 3. 遍历每个文件
    for filename in all_files:
        # 跳过之前的搜索结果文件，避免循环套娃
        if filename.startswith("汇总搜索结果_"):
            continue

        file_path = os.path.join(folder_path, filename)
        scanned_files_count += 1
        
        try:
            reader = PyPDF2.PdfReader(file_path)
            print(f"reading... 📄 {filename} (共 {len(reader.pages)} 页)")
            
            file_hit_count = 0
            
            # 4. 遍历该文件的每一页
            for i, page in enumerate(reader.pages):
                text = page.extract_text()
                if text:
                    # 关键步骤：移除空格和换行，提高匹配率
                    clean_text = text.replace(" ", "").replace("\n", "")
                    
                    if keyword in clean_text:
                        writer.add_page(page)
                        file_hit_count += 1
                        total_found_pages += 1
                        print(f"   ✅ 找到! (第 {i + 1} 页)")
            
            if file_hit_count == 0:
                pass # 这个文件没找到，就不打印额外信息了，保持清爽

        except Exception as e:
            print(f"   ❌ 读取出错: {e}")

    # 5. 保存最终结果
    if total_found_pages > 0:
        with open(output_path, "wb") as f:
            writer.write(f)
        print("\n" + "=" * 30)
        print(f"🎉 全部完成！")
        print(f"📊 扫描文件: {scanned_files_count} 个")
        print(f"📑 提取总页数: {total_found_pages} 页")
        print(f"💾 结果文件: {output_path}")
    else:
        print("\n" + "=" * 30)
        print(f"⚠️ 在所有文件中都未找到包含 '{keyword}' 的内容。")

# ==========================================
# 👇 配置区
# ==========================================

# 1. 这里输入你的【文件夹】路径 (注意是文件夹，不是具体文件)
source_folder_path = '/Users/shane/Desktop/未命名文件夹'

# 2. 这里输入你要搜索的关键词
search_keyword = 'Shishuai'

# 运行
if __name__ == "__main__":
    search_and_merge_from_folder(source_folder_path, search_keyword)
