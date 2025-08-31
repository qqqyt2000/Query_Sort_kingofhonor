import pandas as pd

def convert_frames_to_min_sec(frames):
    """将帧数转换为分钟:秒格式（MM:SS）"""
    total_seconds = int(int(frames) * 66 / 1000)
    minutes = int(total_seconds // 60)
    seconds = int(total_seconds % 60)
    time_format = '%s:%s' % (minutes,seconds)
    return time_format

def process_csv_to_excel(input_csv, output_excel):
    """
    处理CSV文件并生成新的Excel文件
    参数:
        input_csv: 输入的CSV文件路径
        output_excel: 输出的Excel文件路径
    """
    try:
        # 读取原始CSV文件
        df = pd.read_csv(input_csv)

        # 检查文件是否有足够的列
        required_columns = 7  # 需要至少7列才能获取G列(第7列)
        if df.shape[1] < required_columns:
            raise ValueError(f"错误：输入文件至少需要{required_columns}列，但只有{df.shape[1]}列")

        # 创建新DataFrame
        new_df = pd.DataFrame()

        # 1. 将前四列用"_"连接，放入新文件的第一列
        new_df['A'] = df.iloc[:, 0].astype(str) + '_' + df.iloc[:, 1].astype(str) + '_' + \
                               df.iloc[:, 2].astype(str) + '_' + df.iloc[:, 3].astype(str)

        # 2. 将原文件G列(帧号)转换为时间格式，放入新文件的第二列
        new_df['B'] = df.iloc[:, 6].apply(convert_frames_to_min_sec)

        # 3. 将原文件E列(第5列，索引4)放入新文件的第三列
        new_df['C'] = df.iloc[:, 4]

        # 4. 将原文件F列(第6列，索引5)放入新文件的第四列
        new_df['D'] = df.iloc[:, 5]

        # 将新DataFrame写入Excel文件
        new_df.to_excel(output_excel, index=False, engine='openpyxl')
        
        # 创建另一个Excel文件，保存F列中不是"【下发指令】"的值
        filtered_f = df[df.iloc[:, 5] != "【下发指令】"].iloc[:, [5]]
        output_file_filtered = output_excel.replace('.xlsx', '_filtered.xlsx')
        filtered_f.to_excel(output_file_filtered, index=False, header=['Filtered_F_Column'], engine='openpyxl')

        print(f"处理完成，结果已保存到 {output_excel}")
        print(f"过滤后的F列值已保存到 {output_file_filtered}")
        return True

    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        return False

# 使用示例
if __name__ == "__main__":
    input_file = '0729.csv'  # 替换为你的输入文件名
    output_file = '0729_process.xlsx'  # 替换为你想要的输出文件名
    
    process_csv_to_excel(input_file, output_file)