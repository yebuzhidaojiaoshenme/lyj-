import pandas as pd

# 读取 Excel 文件中的所有 sheet
input_file = 'raw_Abb.xlsx'  # 替换为您的输入文件路径
output_file = 'Abb_Combine.xlsx'  # 输出的新文件路径
excel_file = pd.read_excel(input_file, sheet_name=None, header=None)

# 初始化一个空列表来存储每个 sheet 的数据
data_list = []

# 遍历每个 sheet
for sheet_name, df in excel_file.items():
    # 找到包含 "Abbreviation" 和 "Nucleotide(s)" 的标题行
    header_row = None
    for idx, row in df.iterrows():
        row_values = row.values
        print(row_values)
        if "Abbreviation" in row_values and "Nucleotide(s)" in row_values:
            header_row = idx
            break
    
    # 如果没有找到标题行，跳过这个 sheet（可选：可以添加错误提示）
    if header_row is None:
        print(f"警告：在 sheet '{sheet_name}' 中未找到标题 'Abbreviation' 和 'Nucleotide(s)'")
        continue
    
    # 获取标题行的 Series 对象，找到 "Abbreviation" 和 "Nucleotide(s)" 的列名
    header_series = df.iloc[header_row]
    abbrev_col = header_series[header_series == "Abbreviation"].index[0]
    nucleotide_col = header_series[header_series == "Nucleotide(s)"].index[0]
    
    # 提取标题行之下的数据，仅保留这两列
    data_df = df.loc[header_row + 1:, [abbrev_col, nucleotide_col]].copy()
    
    # 将列名重命名为 "Abbreviation" 和 "Nucleotide(s)"
    data_df.columns = ["Abbreviation", "Nucleotide(s)"]
    
    # 添加到数据列表
    data_list.append(data_df)

# 如果没有任何有效数据，打印提示并退出
if not data_list:
    print("错误：所有 sheet 中均未找到有效数据")
else:
    # 合并所有 DataFrame
    combined_df = pd.concat(data_list, ignore_index=True)
    
    # 将合并后的数据写入新的 Excel 文件
    combined_df.to_excel(output_file, index=False)
    print(f"数据已成功合并并保存到 '{output_file}'")
