import pandas as pd
import re

# 文件路径
input_file = 'raw.xlsx'  # 输入文件路径
output_file = 'Output.xlsx'  # 输出文件路径

# 读取 Excel 文件中的所有 sheet
excel_file = pd.read_excel(input_file, sheet_name=None, header=None)

# 初始化存储数据的字典
table_data_dict = {}  # 键是 sheet 名称，值是 DataFrame 列表

# 辅助函数：清理文本
def clean_text(text):
    if pd.isna(text):
        return ''
    text = str(text)
    # 规范化空格，只保留单个空格
    text = re.sub(r'\s+', ' ', text.strip())
    return text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')

# 跨 sheet 保持表格状态
current_table = None  # 当前处理的 Table 编号（如 Table 1、Table 2 等）

# 遍历每个 sheet
for sheet_name, df in excel_file.items():
    print(f"正在处理 Sheet: {sheet_name}...")
    current_row = 0
    while current_row < len(df):
        row = df.iloc[current_row]
        row_values = [clean_text(val) for val in row.values]
        print(f"{sheet_name} 第 {current_row + 1} 行: {row_values}")
        
        # 检测新的 Table 分隔符，使用正则表达式匹配 "Table X"（X 为 number）
        row_text = ' '.join(row_values)
        table_match = re.search(r'Table\s+(\d+)', row_text)
        if table_match:
            table_num = table_match.group(1)
            current_table = f"Table {table_num}"
            print(f"  -> 检测到新表格分隔符：{current_table}")
            current_row += 1
            continue

        header_row = None  # 查找标题行
        table_type = None  # 根据不同的表格类型采用不同的提取方式
        for idx in range(current_row, len(df)):
            row_values = [clean_text(val) for val in df.iloc[idx].values]
            # 检测 MessageRemaining 表格（包含 Duplex 和 ST DEV）
            if ("Duplex" in row_values or "Duplex Name" in row_values) and "ST DEV" in row_values:
                table_type = "MessageRemaining_Table"
                header_row = idx
                print(f"  -> 找到包含 'Duplex' 和 'ST DEV' 标题行，行号 {header_row}")
                break
            # 检测 Screen 表格（包含 Duplex 和 Avg 字段）
            elif ("Duplex" in row_values or "Duplex Name" in row_values) and \
                 "Avg 10 nM" in row_values and "Avg 1 nM" in row_values and "Avg 0.1 nM" in row_values:
                table_type = "Screen_Table"
                header_row = idx
                print(f"  -> 找到包含 'Duplex' 和 'Avg 10 nM/1 nM/0.1 nM' 标题行，行号 {header_row}")
                break
            # 检测 Duplex 表格（包含 Duplex 和序列字段）
            elif any("Duplex" in val or "Duplex Name" in val for val in row_values):
                table_type = "Duplex_Table"
                header_row = idx
                print(f"  -> 找到 'Duplex' 或 'Duplex Name' 标题行，行号 {header_row}")
                break
            # 检测 Abbreviation 表格（包含 Abbreviation 和 Nucleotide(s)）
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
            if re.search(r'Table\s+(\d+)', row_text):
                end_row = idx
                break

        print(f"  -> 表格起始行 {header_row}，结束行 {end_row}")

        # 获取标题行并定位列索引
        header_series = pd.Series(
            [str(val).strip().replace('\n', ' ').replace('\r', ' ').replace('\t', ' ') if pd.notna(val) else '' 
             for val in df.iloc[header_row].values]
        )
        print(header_series)

        # 创建 cleaned_headers 字典，用于所有表格类型
        cleaned_headers = {clean_text(val): idx for idx, val in enumerate(header_series) if pd.notna(val)}

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
            if current_table:
                table_num = int(current_table.replace('Table ', ''))
                save_sheet_name = f"Sheet{table_num}"
            else:
                save_sheet_name = "Sheet2"

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
            save_sheet_name = "Sheet1"  # 默认为 sheet1

        elif table_type == "Screen_Table":
            sd_indices = header_series[header_series == "SD"].index
            if len(sd_indices) < 3:
                print(f"  -> 未找到 3 个 'SD' 列，跳过本 sheet。")
                current_row = end_row
                continue
            print(f"成功找到 3 个 'SD' 列")
            sd10_col = sd_indices[0]
            sd1_col = sd_indices[1]
            sd0_1_col = sd_indices[2]
            try:
                duplex_col = header_series[header_series == "Duplex"].index[0]
                avg10_col = header_series[header_series == "Avg 10 nM"].index[0]
                avg1_col = header_series[header_series == "Avg 1 nM"].index[0]
                avg0_1_col = header_series[header_series == "Avg 0.1 nM"].index[0]
            except IndexError:
                print(f"警告：在 '{sheet_name}' 中未找到所有所需列（Duplex、Avg 10 nM、Avg 1 nM、Avg 0.1 nM）")
                current_row = end_row
                continue
            data_df = df.loc[header_row + 1:end_row - 1, [duplex_col, avg10_col, sd10_col, avg1_col, sd1_col, avg0_1_col, sd0_1_col]].copy()
            data_df.columns = ["Duplex", "Avg 10 nM", "SD (10 nM)", "Avg 1 nM", "SD (1 nM)", "Avg 0.1 nM", "SD (0.1 nM)"]
            if current_table:
                table_num = int(current_table.replace('Table ', ''))
                save_sheet_name = f"Sheet{table_num}"
            else:
                save_sheet_name = "Screen_Data"

        elif table_type == "MessageRemaining_Table":
            # 查找 Duplex 和 ST DEV 列
            duplex_col = None
            stdev_cols = []
            message_cols = []
            for idx, header in enumerate(header_series):
                header_cleaned = clean_text(header)
                if "Duplex" in header_cleaned or "Duplex Name" in header_cleaned:
                    duplex_col = idx
                    print(f"Duplex所在列 {duplex_col}")
                elif "ST DEV" in header_cleaned:
                    stdev_cols.append(idx)
                elif "%of Message Remaining" in header_cleaned or "% of Message Remaining" in header_cleaned:
                    message_cols.append(idx)
            print(f"找到 {len(stdev_cols)} 个 ST DEV: {stdev_cols}")
            print(f"找到 {len(message_cols)} 个 % of Message Remaining: {message_cols}")
            if duplex_col is None or len(stdev_cols) < 3 or len(message_cols) < 3:
                print(f"警告：在 '{sheet_name}' 中未找到所有所需列（Duplex、至少 3 个 ST DEV 和 % of Message Remaining）")
                current_row = end_row
                continue

            # 向上查找上一行提取 nM 信息
            nm_row = header_row - 1
            if nm_row >= 0:
                nm_values = [clean_text(val) for val in df.iloc[nm_row].values]
                nm_headers = [val for val in nm_values if any(str(x) + " nM" in val for x in [10, 1, 0.1])]
                if len(nm_headers) < 3 or not all(str(x) + " nM" in nm_headers for x in [10, 1, 0.1]):
                    print(f"警告：在 '{sheet_name}' 中未找到有效的 nM 标题（10 nM、1 nM、0.1 nM）")
                    current_row = end_row
                    continue
                print(f"  -> 找到 nM 标题：{nm_headers}")
            else:
                nm_headers = ["10 nM", "1 nM", "0.1 nM"]  # 默认值，如果上一行不存在

            # 提取数据
            columns = [duplex_col] + message_cols[:3] + stdev_cols[:3]
            data_df = df.loc[header_row + 1:end_row - 1, columns].copy()
            column_names = ["Duplex"]
            for i, nm in enumerate(nm_headers[:3]):
                column_names.append(f"% of Message Remaining ({nm})")
                column_names.append(f"ST DEV ({nm})")
            data_df.columns = column_names

            # 根据 Table 编号动态设置保存的 sheet
            if current_table:
                table_num = int(current_table.replace('Table ', ''))
                save_sheet_name = f"Sheet{table_num}"
            else:
                save_sheet_name = "MessageRemaining_Data"

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