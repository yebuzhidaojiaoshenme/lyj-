import pandas as pd

# 文件路径
input_file = 'Abb&Unmodified1.xlsx'  # 替换为您的输入文件路径
output_file = 'Abb&Unmodified1_output.xlsx'  # 输出文件路径

# 读取 Excel 文件中的所有 sheet
excel_file = pd.read_excel(input_file, sheet_name=None, header=None)

# 初始化存储数据的字典
table_data_dict = {}  # 键是 sheet 名称，值是 DataFrame 列表

# 辅助函数：清理文本
def clean_text(text):
    if pd.isna(text):
        return ''
    return str(text).strip().replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')

# 遍历每个 sheet
for sheet_name, df in excel_file.items():
    print(f"正在处理 Sheet: {sheet_name}...")
    current_row = 0
    current_table = None
    while current_row < len(df):
        row = df.iloc[current_row]
        row_values = [clean_text(val) for val in row.values]
        print(f"{sheet_name} 第 {current_row + 1} 行: {row_values}")
        
        # 检测新的 Table 分隔符
        row_text = ' '.join(row_values)
        if "Table " in row_text and any(str(i) in row_text for i in range(1, 8)):
            table_num = row_text.split("Table ")[1].split('.')[0]
            current_table = f"Table {table_num}"
            print(f"  -> 检测到新表格分隔符：{current_table}")
            current_row += 1
            continue

        # 查找标题行
        header_row = None
        table_type = None
        for idx in range(current_row, len(df)):
            row_values = [clean_text(val) for val in df.iloc[idx].values]
            if any("Duplex" in val or "Duplex Name" in val for val in row_values):
                table_type = "Duplex_Table"
                header_row = idx
                print(f"  -> 找到 'Duplex' 或 'Duplex Name' 标题行，行号 {header_row}")
                break
            elif any("Abbreviation" in val for val in row_values) and any("Nucleotide(s)" in val for val in row_values):
                table_type = "Abbreviation_Table"
                header_row = idx
                print(f"  -> 找到 'Abbreviation' 和 'Nucleotide(s)' 标题行，行号 {header_row}")
                break

        if header_row is None:
            print(f"  -> 未检测到有效标题行，跳过本段。")
            current_row += 1
            continue

        # 找到表格结束行
        end_row = len(df)
        for idx in range(header_row + 1, len(df)):
            row_values = [clean_text(val) for val in df.iloc[idx].values]
            row_text = ' '.join(row_values)
            if "Table " in row_text and any(str(i) in row_text for i in range(1, 8)):
                end_row = idx
                break

        print(f"  -> 表格起始行 {header_row}，结束行 {end_row}")

        # 获取标题行并定位列索引
        header_series = df.iloc[header_row]
        cleaned_headers = {clean_text(val): idx for idx, val in header_series.items() if pd.notna(val)}
        
        # 根据表格类型提取数据
        if table_type == "Duplex_Table":
            duplex_col = sense_col = antisense_col = None
            for header, col_idx in cleaned_headers.items():
                if "Duplex" in header or "Duplex Name" in header:
                    duplex_col = col_idx
                elif "Sense Sequence" in header:
                    sense_col = col_idx
                elif "Antisense Sequence" in header:
                    antisense_col = col_idx
            if duplex_col is None or sense_col is None or antisense_col is None:
                print(f"警告：在 '{sheet_name}' 中未找到所有所需列（Duplex、Sense Sequence、Antisense Sequence）")
                current_row = end_row
                continue
            data_df = df.loc[header_row + 1:end_row - 1, [duplex_col, sense_col, antisense_col]].copy()
            data_df.columns = ["Duplex", "Sense Sequence 5'to 3'", "Antisense Sequence 5'to 3'"]
            # 根据 current_table 设置保存的 sheet 名称
            if current_table:
                save_sheet_name = f"Sheet{current_table.replace('Table ', '')}"
            else:
                save_sheet_name = "Sheet2"  # 默认保存到 Sheet2，如果未检测到 Table

        elif table_type == "Abbreviation_Table":
            abbrev_col = nucleotide_col = None
            for header, col_idx in cleaned_headers.items():
                if "Abbreviation" in header:
                    abbrev_col = col_idx
                elif "Nucleotide(s)" in header:
                    nucleotide_col = col_idx
            if abbrev_col is None or nucleotide_col is None:
                print(f"警告：在 '{sheet_name}' 中未找到所有所需列（Abbreviation、Nucleotide(s)）")
                current_row = end_row
                continue
            data_df = df.loc[header_row + 1:end_row - 1, [abbrev_col, nucleotide_col]].copy()
            data_df.columns = ["Abbreviation", "Nucleotide(s)"]
            save_sheet_name = "Sheet1"  # 将 Abbreviation 数据保存到 Sheet1，作为 Table1

        # 清理数据：去除空行和重复行
        data_df = data_df.dropna(how='any').drop_duplicates()
        
        # 保存数据
        if not data_df.empty:
            if save_sheet_name not in table_data_dict:
                table_data_dict[save_sheet_name] = []
            table_data_dict[save_sheet_name].append(data_df)
            print(f"  -> 已保存 {table_type} 数据到 {save_sheet_name}")
        
        current_row = end_row

# 将数据保存到输出文件
if table_data_dict:
    with pd.ExcelWriter(output_file) as writer:
        for sheet_name, data_list in table_data_dict.items():
            if data_list:
                combined_df = pd.concat(data_list, ignore_index=True).drop_duplicates()
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  -> 已保存数据到 {sheet_name}")
    print(f"数据已成功保存到 '{output_file}'，包含多个 sheet。")
else:
    print("错误：所有 sheet 中均未找到有效数据")