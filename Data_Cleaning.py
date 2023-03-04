import time

import pandas as pd
import numpy as np
import jieba
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from gensim import corpora, models

def data_clearing():
    # 读取Excel文件
    df = pd.read_excel('./document/comments.xlsx')

    # 空值处理
    df = df.dropna(subset=['Comment'])

    # 剔除纯数字评论，先将其转为空字符串，之后对空字符串统一处理。  用空字符串('')替换纯数字('123')
    df['Comment'] = df['Comment'].str.replace('^[0-9]*$', '', regex=True)

    # 剔除单一重复字符的评论  用空字符串('')替换('111','aaa','....')等
    df['Comment'] = df['Comment'].str.replace(r'^(.)\1*$', '', regex=True)

    # 将评论中的时间转为空字符 用空字符串('')替换('2020/11/20 20:00:00')等
    df['Comment'] = df['Comment'].str.replace(r'\d+/\d+/\d+ \d+:\d+:\d+', '', regex=True)

    # 获取评论文本数据所在的列
    df_comment = df['Comment']

    # 对开头连续重复的部分进行压缩 效果：‘aaabdc’—>‘adbc’，‘很好好好好’—‘很好’
    # 将开头连续重复的部分替换为空''
    prefix_series = df_comment.str.replace(r'(.)\1+$', '', regex=True)
    # 将结尾连续重复的部分替换为空''
    suffix_series = df_comment.str.replace(r'^(.)\1+', '', regex=True)
    for index in range(len(df_comment)):
        # 对开头连续重复的只保留重复内容的一个字符(如'aaabdc'->'abdc')
        if prefix_series[index] != df_comment[index]:
            char = df_comment[index][-1]
            df_comment[index] = prefix_series[index] + char
        # 对结尾连续重复的只保留重复内容的一个字符(如'bdcaaa'->'bdca')
        elif suffix_series[index] != df_comment[index]:
            char = df_comment[index][0]
            df_comment[index] = char + suffix_series[index]
    # 将空字符串转为’np.nan’,在使用dropna（）来进行删除
    # 将空字符串转为'np.nan',即NAN,用于下一步删除这些评论
    df['Comment'].replace(to_replace=r'^\s*$', value=np.nan, regex=True, inplace=True)

    # 删除comment中的空值，并重置索引
    df = df.dropna(subset=['Comment'])
    df.reset_index(drop=True, inplace=True)

    # 删除无关数据
    # 根据实际情况，可以使用模糊匹配或关键词过滤等方法来删除无关数据
    # 例如，如果数据中有“广告”、“垃圾”等关键词，可以使用以下方法删除：
    df = df[~df['Comment'].str.contains('广告|垃圾')]

    # 删除无效数据
    # 如果数据中有缺失值或者不合理的值，可以使用以下方法删除：
    df.dropna(inplace=True)

    # 保存清洗后的数据
    df.to_excel('./document/cleaned_comments.xlsx', index=False)


def participle():
    # 加载停用词
    stopwords = []
    with open('./document/stopwords.txt', 'r', encoding='utf-8') as f:
        stopwords = [line.strip() for line in f.readlines()]
    stopwords.append('\n')
    stopwords.append('"')
    stopwords.append(' ')
    # 读取Excel文件
    df = pd.read_excel('./document/cleaned_comments.xlsx')
    # 对评论文本进行分词
    df['cut_comment'] = df['Comment'].apply(lambda x: [word for word in jieba.cut(x) if word not in stopwords])

    # 生成词频统计表
    words_count = Counter([word for words in df['cut_comment'] for word in words])
    words_count_df = pd.DataFrame(list(words_count.items()), columns=['word', 'count'])
    words_count_df = words_count_df.sort_values(by='count', ascending=False)
    print(f"词频统计表共{len(words_count_df)}")
    words_count_df.to_csv('./document/words_count.txt', index=False, sep='\t')
    # 生成词云图
    wordcloud = WordCloud(font_path='msyh.ttc', background_color='white', max_words=200, max_font_size=60,
                          random_state=42).generate_from_frequencies(words_count)
    plt.figure(figsize=(12, 8))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    plt.savefig('./document/词云图.jpg')
    df.to_excel('./document/comments_tokenized.xlsx', index=False)

def lda_model():
    # 读入数据
    df = pd.read_excel('./document/comments_tokenized.xlsx')
    corpus = df['cut_comment'].apply(lambda x: x.split()).tolist()

    # 建立文本词袋
    dictionary = corpora.Dictionary(corpus)
    doc_term_matrix = [dictionary.doc2bow(doc) for doc in corpus]

    # 训练LDA模型
    # 设置主题数
    num_topics = 3
    # 建立LDA模型
    ldamodel = models.ldamodel.LdaModel(doc_term_matrix, num_topics=num_topics, id2word=dictionary, passes=50)

    # 获取主题及对应的词语和概率
    topics = ldamodel.print_topics(num_topics=num_topics, num_words=10)
    for topic in topics:
        print(topic)

    # 将结果写入excel中
    # 获取各主题及其对应的关键词和概率
    # topic_df = pd.DataFrame()
    # for idx, topic in topics:
    #     keywords = topic.split('+')
    #     keywords = [word.split('*') for word in keywords]
    #     keywords = [(word[1].replace('"', '').strip(), float(word[0])) for word in keywords]
    #     topic_df = topic_df.append({'Topic': idx, 'Keywords': keywords}, ignore_index=True)
    #
    # # 将结果写入excel中
    # topic_df.to_excel('result.xlsx', index=False)


if __name__ == "__main__":
    # data_clearing()
    # participle()
    # time.sleep(5)
    lda_model()
