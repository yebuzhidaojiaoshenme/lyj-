import pandas as pd

# 读取 Excel 文件中的所有 sheet（header=None 保证第一行也作为数据读入）
input_file = 'Modified2_Combine.xlsx'         # 替换为你的输入文件路径
output_file = 'Duplex_Combine.xlsx' # 输出文件路径，可自行修改
excel_file = pd.read_excel(input_file, sheet_name=None, header=None)

# 用于存放所有 sheet 中提取到的 Duplex 数据
data_list = []

# 遍历每个 sheet
for sheet_name, df in excel_file.items():
    print(f"正在处理 Sheet: {sheet_name}...")
    header_row = None  # 记录找到表头的行号

    # 先查找包含关键字段的那一行
    for idx, row in df.iterrows():
        row_values = [str(val).strip().replace('\n', ' ').replace('\r', ' ').replace('\t', ' ') if pd.notna(val) else '' for val in row.values]
        print(row_values)
        # 判断是否在同一行中同时找到了以下字段：
        # "Duplex", "Avg 10 nM", "Avg 1 nM", "Avg 0.1 nM"
        if ("Duplex" in row_values and 
            "Avg 10 nM" in row_values and
            "Avg 1 nM" in row_values and
            "Avg 0.1 nM" in row_values):
            header_row = idx
            break
    
    # 如果找不到表头行，则跳过该 sheet
    if header_row is None:
        print(f"  -> 未检测到包含 Duplex/avg... 字段的表头，跳过本 sheet。")
        continue

    # 从表头行中获取列索引
    # header_series = df.iloc[header_row]
    header_series = pd.Series(
    [str(val).strip().replace('\n', ' ').replace('\r', ' ').replace('\t', ' ') if pd.notna(val) else '' 
     for val in df.iloc[header_row].values]
)

    print(header_series)
    # 查找 "SD" 列（假设按顺序出现 3 次，用于 10 nM、1 nM、0.1 nM）
    sd_indices = header_series[header_series == "SD"].index
    if len(sd_indices) < 3:
        print(f"  -> 未找到 3 个 'SD' 列，跳过本 sheet。")
        continue
    else: print(f"成功找到3个 'SD' 列")
    sd10_col = sd_indices[0]
    sd1_col = sd_indices[1]
    sd0_1_col = sd_indices[2]


    # 找到必需列的索引
    try:
        duplex_col = header_series[header_series == "Duplex"].index[0]
        avg10_col = header_series[header_series == "Avg 10 nM"].index[0]
        avg1_col = header_series[header_series == "Avg 1 nM"].index[0]
        avg0_1_col = header_series[header_series == "Avg 0.1 nM"].index[0]
    except IndexError:
        print(f"  -> 关键列（Duplex/avg nM）索引解析失败，跳过本 sheet。")
        continue

  

    # 提取表头行以下的数据
    # 只保留需要的列：Duplex、avg 10 nM、SD(10 nM)、avg 1 nM、SD(1 nM)、avg 0.1 nM、SD(0.1 nM)
    needed_cols = [
        duplex_col, 
        avg10_col, sd10_col, 
        avg1_col, sd1_col, 
        avg0_1_col, sd0_1_col
    ]
    # 截取标题行以下的内容
    duplex_df = df.loc[header_row + 1:, needed_cols].copy()

    # 重命名列
    duplex_df.columns = [
        "Duplex",
        "avg 10 nM", "SD (10 nM)",
        "avg 1 nM",  "SD (1 nM)",
        "avg 0.1 nM","SD (0.1 nM)"
    ]

    # 将该 sheet 提取到的数据存入 data_list
    data_list.append(duplex_df)
    print(f"  -> 提取到 {len(duplex_df)} 行数据。")

# 如果 data_list 为空，说明没有提取到任何有效数据
if not data_list:
    print("\n错误：所有 sheet 中均未找到有效的 Duplex 类型数据。")
else:
    # 合并所有 DataFrame
    combined_df = pd.concat(data_list, ignore_index=True)
    # 写入新的 Excel 文件
    combined_df.to_excel(output_file, index=False)
    print(f"\n已成功合并并保存到 '{output_file}'，共 {len(combined_df)} 行。")
