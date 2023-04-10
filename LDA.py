import pandas as pd
import jieba
import jieba.analyse
from gensim.models import Word2Vec
from gensim.models.word2vec import LineSentence
import numpy as np


def seg(data_source_path, result_deposit_path):

    # 读取数据
    df = pd.read_excel(result_deposit_path+'cleaned_comments.xlsx')


    # 读取Excel文件
    df = pd.read_excel('comments.xlsx')

    # 分词并进行停用词处理
    jieba.analyse.set_stop_words('哈工大停用词表.txt')  # 加载停用词表
    stop_words = ['三个', '停用词', '添加']  # 自定义停用词

    def process_text(text):
        words = jieba.lcut(text)
        words = [word for word in words if word not in stop_words]
        return ' '.join(words)

    df['分词结果'] = df['评论内容'].apply(process_text)

    # 词频统计
    word_count = {}
    for text in df['分词结果']:
        for word in text.split():
            word_count[word] = word_count.get(word, 0) + 1
    word_count = sorted(word_count.items(), key=lambda x: x[1], reverse=True)

    # 保存词频统计结果到Excel文件
    word_count_df = pd.DataFrame(word_count, columns=['词语', '出现次数'])
    word_count_df.to_excel('word_count.xlsx', index=False)

    # 训练词向量模型
    sentences = LineSentence(df['分词结果'])
    model = Word2Vec(sentences, sg=0, size=100, window=5, min_count=5, workers=4)

    # 生成词向量
    word_vectors = []
    for word in model.wv.vocab:
        word_vectors.append(np.append(word, model.wv[word]))
    word_vectors = np.array(word_vectors)

    # 保存词向量结果到Excel文件
    word_vectors_df = pd.DataFrame(word_vectors,
                                   columns=['词语'] + [f'向量维度{i + 1}' for i in range(model.vector_size)])
    word_vectors_df.to_excel('word_vectors.xlsx', index=False)


if __name__ == "__main__":
    data_source_path = './Data/'
    result_deposit_path = './Result/'
    seg(data_source_path, result_deposit_path)