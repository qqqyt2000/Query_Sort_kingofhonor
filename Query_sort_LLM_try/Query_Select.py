import pandas as pd
from datetime import timedelta

def process_excel(input_file, output_file):
    # 读取Excel文件
    df = pd.read_excel(input_file)
    df = df.dropna(subset=['C'])
    
    print(df['B'][1],type(df['B'][1]))
    # 转换B列为时间格式（hh:mm:ss）
    df['B'] = pd.to_timedelta('00:' + df['B'].astype(str))  # 在前面加上"00:"补全小时部分
    print(df['B'][1],type(df['B'][1]))

    # 按A列值聚类
    grouped = df.groupby('A')
    
    processed_data = []
    
    for abs_value, group in grouped:
        # 按B列去重，保留第一个出现的行
        deduped_group = group.drop_duplicates(subset='B', keep='first')
        
        # 检查子列表中是否有D列不是"【下发指令】"的行
        has_non_instruction = any(deduped_group['D'] != "【下发指令】")
        
        if not has_non_instruction:
            continue  # 跳过这个子列表
        
        # 找出所有非"【下发指令】"的行
        non_instruction_rows = deduped_group[deduped_group['D'] != "【下发指令】"].copy()
        
        # 存储符合条件的行（非指令且前30秒内有指令）
        valid_rows = []
        
        for idx, row in non_instruction_rows.iterrows():
            current_time = row['B']
            time_threshold = current_time - timedelta(seconds=60)
            
            # 查找同一组中在当前时间前30秒内的指令
            instructions_in_window = deduped_group[
                (deduped_group['D'] == "【下发指令】") & 
                (deduped_group['B'] >= time_threshold) & 
                (deduped_group['B'] < current_time)
            ]
            
            # 如果前30秒内有指令，则保留该行
            if not instructions_in_window.empty:
                valid_rows.append(row)
        
        # 如果有符合条件的行，加入最终结果
        if valid_rows:
            processed_data.extend(valid_rows)
    
    # 创建结果DataFrame
    if processed_data:
        result_df = pd.DataFrame(processed_data)
        # 转换时间列回字符串格式以便更好显示
        result_df['B'] = result_df['B'].astype(str).str.extract(r'(\d+:\d{2}:\d{2})')[0]
        # 保存到新Excel文件
        result_df.to_excel(output_file, index=False)
        print(f"处理完成，结果已保存到 {output_file}")
    else:
        print("没有符合条件的数据")

# 使用示例
input_file = '0717_process.xlsx'  # 替换为你的输入文件路径
output_file = '0717_select.xlsx'  # 替换为你想要的输出文件路径
process_excel(input_file, output_file)