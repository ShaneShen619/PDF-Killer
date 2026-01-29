import os
import pandas as pd
import warnings

# 忽略 openpyxl 的样式警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def search_and_sum_employee_hours(folder_path, employee_name):
    """
    遍历指定文件夹下的 Excel 文件。
    1. 搜索包含 employee_name 的工作表（Sheets）。
    2. 定位到 H9 或 I9 单元格寻找 'Arbeitsstunden'。
       - H 列索引 = 7
       - I 列索引 = 8
       - 第 9 行索引 = 8
    3. 如果找到，锁定该列并累加下方所有数值。
    """
    
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print(f"❌ 错误：找不到文件夹 '{folder_path}'")
        return

    # 获取所有 Excel 文件
    excel_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print(f"⚠️ 在 '{folder_path}' 中未找到任何 Excel 文件。")
        return

    print(f"📂 准备在 {len(excel_files)} 个 Excel 文件中搜索员工: {employee_name}")
    print(f"🎯 目标锁定：检查 H9 或 I9 是否有 'Arbeitsstunden'...\n")

    total_hours = 0
    found_employee_sheets = 0

    for filename in excel_files:
        file_path = os.path.join(folder_path, filename)
        
        try:
            # 读取 Excel 文件的所有 sheet 名称
            xls = pd.ExcelFile(file_path)
            
            for sheet_name in xls.sheet_names:
                # =======================================================
                # 步骤 1: 判定是否为该员工的表
                # =======================================================
                is_target_sheet = False
                
                # 策略 A: Sheet 名匹配
                if employee_name.lower() in sheet_name.lower():
                    is_target_sheet = True
                else:
                    # 策略 B: 内容匹配 (读取前 20 行快速扫描)
                    try:
                        df_preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20)
                        if df_preview.astype(str).apply(lambda x: x.str.contains(employee_name, case=False, na=False)).any().any():
                            is_target_sheet = True
                    except Exception:
                        pass 

                if is_target_sheet:
                    print(f"🔎 在文件 [{filename}] -> Sheet [{sheet_name}] 找到员工记录")
                    
                    # =======================================================
                    # 步骤 2: 检查 H9 (row 8, col 7) 或 I9 (row 8, col 8)
                    # =======================================================
                    try:
                        # 读取整个表，header=None 保证我们可以用数字索引访问
                        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                        
                        target_col_index = -1
                        start_row_index = 9 # 数据从第 10 行开始 (index 9)
                        
                        # 确保表格至少有 9 行
                        if len(df) > 8:
                            row_9 = df.iloc[8] # 获取第 9 行 (index 8)
                            
                            # 检查 H9 (索引 7)
                            # 确保该行至少有 8 列
                            if len(row_9) > 7 and isinstance(row_9[7], str) and "stunden" in row_9[7]:
                                target_col_index = 7
                                print(f"   📍 在 H9 找到表头 'Arbeitsstunden'")
                            
                            # 如果 H9 没找到，检查 I9 (索引 8)
                            elif len(row_9) > 8 and isinstance(row_9[8], str) and "stunden" in row_9[8]:
                                target_col_index = 8
                                print(f"   📍 在 I9 找到表头 'Arbeitsstunden'")

                        if target_col_index != -1:
                            # ===================================================
                            # 步骤 3: 提取该列数据并求和
                            # ===================================================
                            # df.iloc[起始行:, 列索引]
                            hours_series = df.iloc[start_row_index:, target_col_index]
                            
                            # 转换为数字 (非数字转 NaN)
                            hours_numeric = pd.to_numeric(hours_series, errors='coerce')
                            
                            # 求和
                            sheet_sum = hours_numeric.sum()
                            total_hours += sheet_sum
                            found_employee_sheets += 1
                            
                            print(f"   ✅ 本表工时合计: {sheet_sum:.2f}")
                        else:
                            print(f"   ⚠️  找到员工表，但在 H9 或 I9 未找到 'Arbeitsstunden'，跳过。")
                            
                    except Exception as e:
                        print(f"   ❌ 读取工作表数据出错: {e}")

        except Exception as e:
            print(f"❌ 无法读取文件 {filename}: {e}")

    # 最终结果
    print("\n" + "=" * 40)
    if found_employee_sheets > 0:
        print(f"🎉 统计完成！")
        print(f"👤 员工: {employee_name}")
        print(f"📄 包含数据的表格数: {found_employee_sheets}")
        print(f"⏱️  总工作时长: {total_hours:.2f} 小时")
    else:
        print(f"⚠️ 未找到员工 '{employee_name}' 的有效工时记录。")
    print("=" * 40)

# ==========================================
# 👇 用户配置区
# ==========================================

# 1. 文件夹路径
FOLDER_PATH = '/Users/shane/Desktop/未命名文件夹'

# 2. 员工姓名
EMPLOYEE_NAME = 'Zhikuan'

# 运行
if __name__ == "__main__":
    search_and_sum_employee_hours(FOLDER_PATH, EMPLOYEE_NAME)
