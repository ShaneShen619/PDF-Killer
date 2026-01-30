import os
import pandas as pd
import warnings
import re

# 忽略 openpyxl 的样式警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def sum_all_employees_hours(folder_path):
    """
    遍历指定文件夹下的所有 Excel 文件。
    统计所有包含 'Arbeitsstunden' (在 H9 或 I9) 的工作表的工时总和。
    """
    
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print(f"❌ 错误：找不到文件夹 '{folder_path}'")
        return

    # 获取所有 Excel 文件 (排除以 ~$ 开头的临时文件)
    excel_files = [f for f in os.listdir(folder_path) 
                   if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    
    if not excel_files:
        print(f"⚠️ 在 '{folder_path}' 中未找到任何 Excel 文件。")
        return

    print(f"📂 准备扫描 {len(excel_files)} 个 Excel 文件...\n")
    print(f"🎯 识别规则：H9 或 I9 包含 'Arbeitsstunden'")
    print("-" * 60)
    print(f"{ '文件名':<30} | {'工作表 (员工)':<20} | {'工时':<10}")
    print("-" * 60)

    grand_total_hours = 0
    valid_sheets_count = 0

    for filename in excel_files:
        file_path = os.path.join(folder_path, filename)
        print(f"📄 正在处理文件: {filename} ...")
        
        try:
            # 读取 Excel 文件的所有 sheet 名称
            xls = pd.ExcelFile(file_path)
            
            for sheet_name in xls.sheet_names:
                try:
                    # 读取整个表，header=None 保证我们可以用数字索引访问
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                    
                    target_col_index = -1
                    start_row_index = 9 # 数据从第 10 行开始 (index 9)
                    
                    # 检查 H9 或 I9 是否有关键字
                    if len(df) > 8:
                        row_9 = df.iloc[8] # 获取第 9 行 (index 8)
                        
                        # 获取 H9 (index 7) 和 I9 (index 8) 的值，并转为字符串
                        val_h = str(row_9[7]).strip() if len(row_9) > 7 else ""
                        val_i = str(row_9[8]).strip() if len(row_9) > 8 else ""
                        
                        # 打印调试信息，看看到底读到了什么
                        # print(f"   [Debug] {sheet_name} -> H9: '{val_h}' | I9: '{val_i}'")

                        # 定义匹配函数：不区分大小写，支持更宽泛的关键词 'arbeits'
                        def is_match(text):
                            t = text.lower()
                            return "arbeits" in t

                        # 检查 H9
                        if is_match(val_h):
                            target_col_index = 7
                        # 检查 I9
                        elif is_match(val_i):
                            target_col_index = 8

                    if target_col_index != -1:
                        # 提取该列数据并求和
                        # 只读取到 Excel 第 41 行 (index 40)，防止读取到底部的汇总信息
                        # start_row_index = 9 (Excel 第 10 行)
                        # 结束 index = 41 (对应 Excel 第 42 行之前，即包含 Excel 第 41 行)
                        hours_series = df.iloc[start_row_index:41, target_col_index]
                        hours_numeric = pd.to_numeric(hours_series, errors='coerce')
                        
                        sheet_sum = hours_numeric.sum()
                        
                        # 累加到全局总和
                        grand_total_hours += sheet_sum
                        valid_sheets_count += 1
                        
                        # ---------------------------------------------------
                        # 步骤 3: 提取员工姓名 (精准锁定 B3)
                        # ---------------------------------------------------
                        # B3 对应: Row index 2, Column index 1
                        employee_name_display = str(sheet_name) # 默认用 Sheet 名兜底
                        
                        try:
                            if len(df) > 2: # 确保至少有 3 行
                                cell_val = df.iloc[2, 1] # [Row 2, Col 1] = B3
                                if pd.notna(cell_val):
                                    val_str = str(cell_val).strip()
                                    if len(val_str) > 0:
                                        # 简化逻辑：捕获 Nr(.)? 之后， [ 之前的所有内容
                                        # Nr\.? 表示 "." 可有可无
                                        match = re.search(r"Nr\.?[[\s\xa0]*(.+?)[\s\xa0]*\[", val_str)
                                        
                                        if match:
                                            # 直接获取中间的全部内容，并移除可能残留的 "."
                                            raw_name = match.group(1).strip()
                                            employee_name_display = raw_name.replace(".", "")
                                        else:
                                            # 如果正则没匹配上，还是显示原字符串（截断一下）
                                            employee_name_display = val_str

                        except Exception:
                            pass

                        # ---------------------------------------------------
                        # 步骤 4: 打印输出
                        # ---------------------------------------------------
                        # 简化文件名：只取前 35 个字符，通常能包含日期和门店名
                        simple_filename = filename.replace(".xlsx", "").replace(".xls", "")
                        if len(simple_filename) > 35:
                            simple_filename = simple_filename[:35]
                        
                        # 截断过长的名字
                        if len(employee_name_display) > 20:
                            employee_name_display = employee_name_display[:20]
                        
                        print(f"{simple_filename:<35} | {employee_name_display:<20} | {sheet_sum:.2f}")

                except Exception as e:
                    # 读取单个 sheet 出错不影响整体
                    pass

        except Exception as e:
            print(f"❌ 无法读取文件 {filename}: {e}")

    # 最终结果
    print("-" * 60)
    print(f"🎉 统计完成！")
    print(f"📊 有效工时表数量: {valid_sheets_count}")
    print(f"⏱️  所有员工总工时之和: {grand_total_hours:.2f} 小时")
    print("=" * 60)

# ==========================================
# 👇 用户配置区
# ==========================================

# 1. 文件夹路径
FOLDER_PATH = '/Users/shane/Desktop/未命名文件夹'

# 运行
if __name__ == "__main__":
    sum_all_employees_hours(FOLDER_PATH)
