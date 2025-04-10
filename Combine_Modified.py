import pandas as pd

# 读取 Excel 文件中的所有 sheet
input_file = 'Unmodified1.xlsx'  # 替换为您的输入文件路径
output_file = 'Unmodified1_output.xlsx'  # 输出的新文件路径
excel_file = pd.read_excel(input_file, sheet_name=None, header=None)

# 初始化一个字典来存储每个 Table 对应的数据
table_data_dict = {}  # 键是 Table 编号（如 'Table 2'），值是 DataFrame 列表
current_table = None  # 当前处理的 Table 编号
current_data_list = []  # 当前 Table 的数据列表

# 辅助函数：清理文本
def clean_text(text):
    if pd.isna(text):
        return ''
    return str(text).strip().replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')

# 遍历每个 sheet
for sheet_name, df in excel_file.items():
    print(f"正在处理 Sheet: {sheet_name}...")
    current_row = 0
    while current_row < len(df):
        row = df.iloc[current_row]
        row_values = [clean_text(val) for val in row.values]
        print(f"{sheet_name} 第 {current_row + 1} 行: {row_values}")
        
        # 检测新的 Table 分隔符
        row_text = ' '.join(row_values)
        if "Table " in row_text and any(str(i) in row_text for i in range(1, 8)):  # 假设 Table 1 到 Table 7
            table_num = row_text.split("Table ")[1].split('.')[0]  # 提取 Table 编号（如 '2'）
            new_table = f"Table {table_num}"
            print(f"  -> 检测到新表格分隔符：{new_table}")
            
            # 保存当前 Table 的数据（如果有）
            if current_data_list and current_table:
                table_data_dict[current_table] = current_data_list
                print(f"  -> 已保存 {current_table} 数据")
            
            current_table = new_table  # 更新当前 Table 编号
            current_data_list = []  # 清空当前数据列表，开始新表格
            current_row += 1  # 跳过 Table 行
            
            # 从下一行开始搜索新的标题行
            header_row = None
            for idx in range(current_row, len(df)):
                row_values = [clean_text(val) for val in df.iloc[idx].values]
                if any("Duplex" in val or "Duplex Name" in val for val in row_values):
                    header_row = idx
                    print(f"  -> 找到新表格的标题 'Duplex' 或 'Duplex Name'，行号 {header_row}")
                    break
            
            if header_row is None:
                print(f"  -> 未检测到新表格的 'Duplex' 或 'Duplex Name'，从行 {current_row} 开始，跳过本段。")
                current_row += 1
                continue
            else:
                current_row = header_row  # 更新当前行到标题行
                continue

        # 查找包含 "Duplex" 或 "Duplex Name" 的标题行
        header_row = None
        for idx in range(current_row, len(df)):
            row_values = [clean_text(val) for val in df.iloc[idx].values]
            if any("Duplex" in val or "Duplex Name" in val for val in row_values):
                header_row = idx
                print(f"  -> 找到标题 'Duplex' 或 'Duplex Name'，行号 {header_row}")
                break

        if header_row is None:
            print(f"  -> 未检测到 'Duplex' 或 'Duplex Name'，从行 {current_row} 开始，跳过本段。")
            current_row += 1
            continue

        # 找到结束行（下一个 Table 或文件末尾）
        end_row = len(df)
        for idx in range(header_row + 1, len(df)):
            row_values = [clean_text(val) for val in df.iloc[idx].values]
            row_text = ' '.join(row_values)
            if "Table " in row_text and any(str(i) in row_text for i in range(1, 8)):
                end_row = idx
                break

        print(f"  -> 表格起始行 {header_row}，结束行 {end_row}")

        # 获取标题行的 Series 对象，找到 "Duplex"、"Sense Sequence 5'to 3'" 和 "Antisense Sequence 5'to 3'" 的列名
        header_series = df.iloc[header_row]
        cleaned_headers = {clean_text(val): idx for idx, val in header_series.items() if pd.notna(val)}
        
        # 查找列名（允许部分匹配）
        duplex_col = None
        sense_col = None
        antisense_col = None
        
        for header, col_idx in cleaned_headers.items():
            if "Duplex" in header or "Duplex Name" in header:
                duplex_col = col_idx
            elif "Sense Sequence" in header:
                sense_col = col_idx
            elif "Antisense Sequence" in header:
                antisense_col = col_idx

        # 检查是否找到所有列
        if duplex_col is None or sense_col is None or antisense_col is None:
            print(f"警告：在 '{sheet_name}' 中未找到所有所需列（Duplex、Sense Sequence 5'to 3'、Antisense Sequence 5'to 3'）")
            current_row = end_row
            continue
        print(f"  -> 找到所有所需列：Duplex, Sense Sequence, Antisense Sequence")

        # 提取标题行之下的数据，仅保留这三列，直到结束行
        data_df = df.loc[header_row + 1:end_row - 1, [duplex_col, sense_col, antisense_col]].copy()
        
        # 将列名重命名为 "Duplex"、"Sense Sequence 5'to 3'" 和 "Antisense Sequence 5'to 3'"
        data_df.columns = ["Duplex", "Sense Sequence 5'to 3'", "Antisense Sequence 5'to 3'"]
        
        # 去除空行和重复行
        data_df = data_df.dropna(how='any').drop_duplicates()
        
        # 添加到当前 Table 的数据列表
        if not data_df.empty:
            if current_table:
                current_data_list.append(data_df)
            else:
                print(f"警告：未检测到 Table 编号，跳过数据保存")
        
        current_row = end_row  # 移动到下一个表格的起始位置

# 保存所有 Table 的数据到不同的 sheet
if table_data_dict or current_data_list:
    with pd.ExcelWriter(output_file) as writer:
        # 处理字典中的数据
        for table_num, data_list in table_data_dict.items():
            if data_list:
                combined_df = pd.concat(data_list, ignore_index=True)
                sheet_name = f"Sheet{table_num.replace('Table ', '')}"  # 例如 "Sheet2"
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  -> 已保存 {table_num} 数据到 {sheet_name}")
        
        # 处理剩余的 current_data_list（如果有且非空）
        if current_data_list and current_table:
            combined_df = pd.concat(current_data_list, ignore_index=True)
            sheet_name = f"Sheet{current_table.replace('Table ', '')}"
            combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  -> 已保存 {current_table} 数据到 {sheet_name}")
    print(f"数据已成功合并并保存到 '{output_file}'，包含多个 sheet。")
else:
    print("错误：所有 sheet 中均未找到有效数据")