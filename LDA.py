import pandas as pd
import gensim
from gensim import corpora, models
from nltk.corpus import stopwords
import pyLDAvis.gensim

def func1(result_deposit_path):
    # 读取 Excel 数据
    df = pd.read_excel(result_deposit_path+'segmented_comments.xlsx')
    # 处理空值
    df = df.fillna("")

    # 清洗文本数据
    texts = []
    for i in range(len(df)):
        text = df['Segmented Comment'][i].lower()
        text = gensim.utils.simple_preprocess(text, deacc=True, min_len=2)
        texts.append(text)

    # 构建词典和文档-词矩阵
    dictionary = corpora.Dictionary(texts)
    corpus = [dictionary.doc2bow(text) for text in texts]

    # LDA 主题模型训练
    lda_model = models.LdaModel(corpus=corpus, id2word=dictionary, num_topics=4)

    # 输出每个主题的关键词
    for topic in lda_model.print_topics():
        print(topic)

    # 使用 pyLDAvis 可视化结果
    vis = pyLDAvis.gensim.prepare(lda_model, corpus, dictionary)
    pyLDAvis.save_html(vis, result_deposit_path+'lda_visualization.html')

if __name__ == "__main__":
    data_source_path = './Data/'
    result_deposit_path = './Result/'
    func1(result_deposit_path)
