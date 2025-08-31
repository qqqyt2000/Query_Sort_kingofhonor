import pandas as pd
from collections import defaultdict
import jieba
import jieba.posseg as pseg
from openpyxl import Workbook
from gensim.models import Word2Vec
from sklearn.cluster import KMeans
import numpy as np
from itertools import groupby

def extract_and_cluster_phrases(input_file, output_file, sheet_name='Sheet1', text_column='对话', 
                              top_n=50, min_count=5, similarity_threshold=0.6):
    """
    提取动名词词组并归类语义相近的词组
    
    参数:
        input_file: 输入的Excel文件路径
        output_file: 输出的Excel文件路径
        sheet_name: 要读取的工作表名
        text_column: 包含文本的列名
        top_n: 返回最高频的N类词组
        min_count: 词组最低出现次数(过滤低频词组)
        similarity_threshold: 语义相似度阈值(0-1)
    """
    try:
        # 1. 读取Excel文件
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        if text_column not in df.columns:
            raise ValueError(f"列 '{text_column}' 不存在于Excel文件中")
        
        # 2. 提取所有动名词词组
        phrase_counter = defaultdict(int)
        all_sentences = []  # 用于训练词向量
        
        for text in df[text_column]:
            if pd.isna(text):
                continue
                
            text = str(text)
            all_sentences.append([word.word for word in pseg.cut(text)])
            
            words = list(pseg.cut(text))
            for i in range(len(words)-1):
                # 提取动词+名词组合
                if words[i].flag.startswith('v') and words[i+1].flag.startswith('n'):
                    phrase = f"{words[i].word}{words[i+1].word}"
                    phrase_counter[phrase] += 1
                
                # 提取动词+名词+名词组合
                if i < len(words)-2:
                    if (words[i].flag.startswith('v') and 
                        words[i+1].flag.startswith('n') and 
                        words[i+2].flag.startswith('n')):
                        phrase = f"{words[i].word}{words[i+1].word}{words[i+2].word}"
                        phrase_counter[phrase] += 1
        
        # 3. 过滤低频词组
        phrases = [phrase for phrase, count in phrase_counter.items() if count >= min_count]
        if not phrases:
            raise ValueError("没有找到足够高频的动名词词组")
        
        # 4. 训练词向量模型(简单版)
        # 更准确的做法是使用预训练的中文词向量模型
        model = Word2Vec(sentences=all_sentences, vector_size=100, window=5, min_count=1, workers=4)
        
        # 5. 获取词向量(对词组中的词向量取平均)
        phrase_vectors = []
        valid_phrases = []
        
        for phrase in phrases:
            words_in_phrase = []
            # 尝试分词(处理未登录词)
            for word in jieba.cut(phrase):
                if word in model.wv:
                    words_in_phrase.append(word)
            
            if words_in_phrase:
                # 计算词组的平均向量
                vector = np.mean([model.wv[word] for word in words_in_phrase], axis=0)
                phrase_vectors.append(vector)
                valid_phrases.append(phrase)
        
        if not valid_phrases:
            raise ValueError("无法为词组生成有效的词向量")
        
        # 6. 聚类相似的词组
        # 使用KMeans聚类，聚类数量自动确定为词组数量的1/4
        n_clusters = max(2, len(valid_phrases) // 4)
        kmeans = KMeans(n_clusters=n_clusters, random_state=42).fit(phrase_vectors)
        
        # 7. 组织聚类结果
        clustered_phrases = defaultdict(list)
        for phrase, label in zip(valid_phrases, kmeans.labels_):
            clustered_phrases[label].append((phrase, phrase_counter[phrase]))
        
        # 8. 为每个聚类选择代表词组(出现次数最多的)
        cluster_results = []
        for label, group in clustered_phrases.items():
            group_sorted = sorted(group, key=lambda x: x[1], reverse=True)
            representative = group_sorted[0][0]
            total_count = sum([cnt for _, cnt in group])
            members = [ph for ph, _ in group_sorted]
            cluster_results.append((representative, total_count, members))
        
        # 按总频次排序
        cluster_results.sort(key=lambda x: x[1], reverse=True)
        
        # 9. 写入Excel文件
        wb = Workbook()
        ws = wb.active
        ws.title = "语义归类词组"
        
        # 写入表头
        ws.append(["代表词组", "总频次", "同类词组"])
        
        # 写入数据(最多top_n类)
        for representative, total_count, members in cluster_results[:top_n]:
            # 限制同类词组显示数量，避免单元格过长
            members_str = ", ".join(members[:10]) + ("..." if len(members) > 10 else "")
            ws.append([representative, total_count, members_str])
        
        # 调整列宽
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(output_file)
        print(f"成功提取并归类动名词词组，保存到: {output_file}")
        print(f"共找到 {len(cluster_results)} 类词组，输出前 {top_n} 类")
        
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")

# 使用示例
if __name__ == "__main__":
    # 输入文件路径(修改为你的实际路径)
    input_excel = "0717_negative_analysis.xlsx"
    # 输出文件路径
    output_excel = "0717_语义归类词组.xlsx"
    
    # 调用函数
    extract_and_cluster_phrases(
        input_file=input_excel,
        output_file=output_excel,
        sheet_name='Sheet1',      # 工作表名
        text_column='B',       # 包含对话文本的列名
        top_n=10,               # 提取前20类高频词组
        min_count=3,             # 词组最低出现次数
        similarity_threshold=0.6  # 语义相似度阈值
    )