import jieba.posseg as psg
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from wordcloud import WordCloud


def get_word():
    # 去重，去除完全重复的数据
    reviews = pd.read_excel(result_deposit_path + "segmented_comments.xlsx")
    content = reviews['content']

    # 分词
    worker = lambda s: [(x.word, x.flag) for x in psg.cut(s)]  # 自定义简单分词函数
    seg_word = content.apply(worker)
    # 将词语转为数据框形式，一列是词，一列是词语所在的句子ID，最后一列是词语在该句子的位置
    n_word = seg_word.apply(lambda x: len(x))  # 每一评论中词的个数

    n_content = [[x + 1] * y for x, y in zip(list(seg_word.index), list(n_word))]
    index_content = sum(n_content, [])  # 将嵌套的列表展开，作为词所在评论的id

    seg_word = sum(seg_word, [])
    word = [x[0] for x in seg_word]  # 词

    nature = [x[1] for x in seg_word]  # 词性

    content_type = [[x] * y for x, y in zip(list(reviews['content_type']),
                                            list(n_word))]
    content_type = sum(content_type, [])  # 评论类型

    result = pd.DataFrame({"index_content": index_content,
                           "word": word,
                           "nature": nature,
                           "content_type": content_type})

    # 删除标点符号
    result = result[result['nature'] != 'x']  # x表示标点符号

    # 删除停用词
    stop_path = open(data_source_path + "cn_stopwords.txt", 'r', encoding='UTF-8')
    stop = stop_path.readlines()
    stopwordsupdate = ['哈哈哈', "喝", "买", "非常", '东方', '树叶', '饮料', '没有', '很', ' ',
                       '都', '不', '太', '没', '说', '搬', '放'
                                                           '购买',
                       '特别', '一直', '夏天', 'hellip', 'helliphellip', '第二天', '茉莉花', '茉莉花茶',
                       '乌龙茶', '普洱茶',
                       '红茶', '绿茶', '农夫山泉', '越来越', '京东', '茶', '茶饮料', '淡淡的', '会',
                       '普洱', '起来',
                       '这款', '款']

    stop.extend(stopwordsupdate)
    stop = [x.replace('\n', '') for x in stop]
    word = list(set(word) - set(stop))
    result = result[result['word'].isin(word)]

    # 构造各词在对应评论的位置列
    n_word = list(result.groupby(by=['index_content'])['index_content'].count())
    index_word = [list(np.arange(0, y)) for y in n_word]
    index_word = sum(index_word, [])  # 表示词语在改评论的位置

    # 合并评论id，评论中词的id，词，词性，评论类型
    result['index_word'] = index_word

    # 提取含有名词类的评论
    ind = result[['n' in x for x in result['nature']]]['index_content'].unique()
    result = result[[x in ind for x in result['index_content']]]

    # 将结果写出
    result.to_csv(result_deposit_path + "word.csv", index=False, encoding='gbk')


def get_posorneg():
    word = pd.read_csv(result_deposit_path + "word.csv", encoding='GBK')
    # 读入正面、负面情感评价词
    pos_comment = pd.read_csv(data_source_path + "《知网》情感分析用词语集（beta版）\正面评价词语（中文）.txt", header=None,
                              sep="\t", encoding='utf-8', engine='python')
    neg_comment = pd.read_csv(data_source_path + "《知网》情感分析用词语集（beta版）\负面评价词语（中文）.txt", header=None,
                              sep="\t", encoding='utf-8', engine='python')
    pos_emotion = pd.read_csv(data_source_path + "《知网》情感分析用词语集（beta版）\正面情感词语（中文）.txt", header=None,
                              sep="\t", encoding='utf-8', engine='python')
    neg_emotion = pd.read_csv(data_source_path + "《知网》情感分析用词语集（beta版）\负面情感词语（中文）.txt", header=None,
                              sep="\t", encoding='utf-8', engine='python')

    # 合并情感词与评价词
    positive = set(pos_comment.iloc[:, 0]) | set(pos_emotion.iloc[:, 0])
    negative = set(neg_comment.iloc[:, 0]) | set(neg_emotion.iloc[:, 0])
    intersection = positive & negative  # 正负面情感词表中相同的词语
    positive = list(positive - intersection)
    negative = list(negative - intersection)
    positive = pd.DataFrame({"word": positive,
                             "weight": [1] * len(positive)})
    negative = pd.DataFrame({"word": negative,
                             "weight": [-1] * len(negative)})

    posneg = positive.append(negative)
    #  将分词结果与正负面情感词表合并，定位情感词
    data_posneg = posneg.merge(word, left_on='word', right_on='word',
                               how='right')
    data_posneg = data_posneg.sort_values(by=['index_content', 'index_word'])

    # -------------------------------------------------------------------------------------------------------
    # 根据情感词前时候有否定词或双层否定词对情感值进行修正
    # 载入否定词表
    notdict = pd.read_csv(data_source_path + "not.csv")

    # 处理否定修饰词
    data_posneg['amend_weight'] = data_posneg['weight']  # 构造新列，作为经过否定词修正后的情感值
    data_posneg['id'] = np.arange(0, len(data_posneg))
    only_inclination = data_posneg.dropna()  # 只保留有情感值的词语
    only_inclination.index = np.arange(0, len(only_inclination))
    index = only_inclination['id']

    for i in np.arange(0, len(only_inclination)):
        review = data_posneg[data_posneg['index_content'] ==
                             only_inclination['index_content'][i]]  # 提取第i个情感词所在的评论
        review.index = np.arange(0, len(review))
        affective = only_inclination['index_word'][i]  # 第i个情感值在该文档的位置
        if affective == 1:
            ne = sum([i in notdict['term'] for i in review['word'][affective - 1]])
            if ne == 1:
                data_posneg['amend_weight'][index[i]] = - \
                    data_posneg['weight'][index[i]]
        elif affective > 1:
            ne = sum([i in notdict['term'] for i in review['word'][[affective - 1,
                                                                    affective - 2]]])
            if ne == 1:
                data_posneg['amend_weight'][index[i]] = - \
                    data_posneg['weight'][index[i]]

    # 更新只保留情感值的数据
    only_inclination = only_inclination.dropna()

    # 计算每条评论的情感值
    emotional_value = only_inclination.groupby(['index_content'],
                                               as_index=False)['amend_weight'].sum()

    # 去除情感值为0的评论
    emotional_value = emotional_value[emotional_value['amend_weight'] != 0]

    # 给情感值大于0的赋予评论类型（content_type）为pos,小于0的为neg
    emotional_value['a_type'] = ''
    emotional_value['a_type'][emotional_value['amend_weight'] > 0] = 'pos'
    emotional_value['a_type'][emotional_value['amend_weight'] < 0] = 'neg'

    # -------------------------------------------------------------------------------------------------------

    # 查看情感分析结果
    result = emotional_value.merge(word,
                                   left_on='index_content',
                                   right_on='index_content',
                                   how='left')

    result = result[['index_content', 'content_type', 'a_type']].drop_duplicates()
    confusion_matrix = pd.crosstab(result['content_type'], result['a_type'],
                                   margins=True)  # 制作交叉表
    (confusion_matrix.iat[0, 0] + confusion_matrix.iat[1, 1]) / confusion_matrix.iat[2, 2]

    # 提取正负面评论信息
    ind_pos = list(emotional_value[emotional_value['a_type'] == 'pos']['index_content'])
    ind_neg = list(emotional_value[emotional_value['a_type'] == 'neg']['index_content'])
    posdata = word[[i in ind_pos for i in word['index_content']]]
    negdata = word[[i in ind_neg for i in word['index_content']]]

    # 正面情感词词云
    freq_pos = posdata.groupby(by=['word'])['word'].count()
    freq_pos = freq_pos.sort_values(ascending=False)
    font_path = data_source_path + 'simhei.ttf'  # 设置中文字体
    backgroud_Image = plt.imread(data_source_path + 'backgrand.jpg')
    wordcloud = WordCloud(font_path=font_path,
                          max_words=100,
                          background_color='white',
                          mask=backgroud_Image)
    wordcloud.generate_from_frequencies(freq_pos)
    wordcloud.to_file(result_deposit_path + "freq_pos.png")

    # 负面情感词词云
    freq_neg = negdata.groupby(by=['word'])['word'].count()
    freq_neg = freq_neg.sort_values(ascending=False)
    wordcloud.generate_from_frequencies(freq_neg)
    wordcloud.to_file(result_deposit_path + "freq_neg.png")

    # 将结果写出，每条评论作为一行
    posdata.to_csv(result_deposit_path + "posdata.csv", index=False, encoding='GBK')
    negdata.to_csv(result_deposit_path + "negdata.csv", index=False, encoding='GBK')


if __name__ == "__main__":
    data_source_path = './Data/'
    result_deposit_path = './Result/'
    get_word()
    get_posorneg()
