import re
from collections import Counter
import jieba
import matplotlib.pyplot as plt
import nltk
import numpy as np
import openpyxl
import pandas as pd
from PIL import Image
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
from nltk.tokenize import word_tokenize
from wordcloud import WordCloud


# 下载必要的 NLTK 数据
def download_NLTK():
    nltk.download('punkt')
    nltk.download('stopwords')
    nltk.download('wordnet')


# 数据清洗主函数
def data_cleaning(data_source_path, result_deposit_path):
    # 读取 Excel 文件
    df = pd.read_excel(data_source_path + "comments.xlsx")

    # 删除空值行
    df.dropna(subset=['Comment'], inplace=True)

    # 去除重复评论
    df.drop_duplicates(subset=['Comment'], keep='first', inplace=True)

    # 创建 DataFrame 副本
    df_clean = df.copy()

    # 清理评论文本
    df_clean['Cleaned Comment'] = df_clean['Comment'].apply(clean_text)

    # 删除空白行
    df_clean.dropna(subset=['Cleaned Comment'], inplace=True)

    # 去除字符串中的空格和其他字符
    df_clean['Cleaned Comment'] = df_clean['Cleaned Comment'].str.strip()
    df_clean = df_clean.replace('', np.nan)

    # 再次删除空白行
    df_clean.dropna(subset=['Cleaned Comment'], inplace=True)

    # 剔除纯数字或纯字母评论
    df_clean = df_clean[~df_clean['Cleaned Comment'].str.match(r'^\d+$|^[a-zA-Z]+$')]

    # 保存清洗后的数据到新的 Excel 文件
    df_clean.to_excel(result_deposit_path + 'cleaned_comments.xlsx', index=False)

    # 分析函数
    data_analyze(df_clean, result_deposit_path)


# 定义清理函数
def clean_text(text):
    # 将文本转换为小写
    text = text.lower()
    # 去除标点符号、数字和特殊字符
    text = re.sub(r'[^\w\s]', '', text)
    text = re.sub(r'\d+', '', text)
    text = re.sub(r'\s+', ' ', text)

    # 去除停用词
    stop_words = set(stopwords.words('english'))
    word_tokens = word_tokenize(text)
    filtered_text = [word for word in word_tokens if word not in stop_words]
    # 词形还原
    lemmatizer = WordNetLemmatizer()
    lemmatized_text = [lemmatizer.lemmatize(word) for word in filtered_text]
    # 去除多余的空格和换行符
    cleaned_text = ' '.join(lemmatized_text)
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()
    return cleaned_text


# 对清洗后的评论进行分析
def data_analyze(df_clean, result_deposit_path):
    # 统计评论数量、平均长度等基本信息
    print(f"Total number of comments: {len(df_clean)}")
    print(f"Average length of comments: {df_clean['Cleaned Comment'].str.len().mean():.2f}")
    # 统计最常见的词汇
    all_words = []
    for comment in df_clean['Cleaned Comment']:
        all_words += word_tokenize(comment)

    stop_words = stopwords.words('english')
    filtered_words = [word.lower() for word in all_words if word.lower() not in stop_words and word.isalpha()]

    word_freq = Counter(filtered_words)
    most_common_words = word_freq.most_common(10)

    print("Most common words:")
    for word, freq in most_common_words:
        print(f"{word}: {freq}")

    # 取前10个最常见的词汇
    top_words = [word[0] for word in most_common_words]
    top_word_freq = [word[1] for word in most_common_words]

    # 生成柱状图
    plt.rcParams['font.family'] = ['SimHei']  # 将字体设置为 SimHei
    plt.bar(top_words, top_word_freq)
    plt.title("Top 10 Most Frequent Words")
    plt.xlabel("Words")
    plt.ylabel("Frequency")
    plt.savefig(result_deposit_path + 'top_words_frequency.png')


def data_segmentation(data_source_path, result_deposit_path):
    df = pd.read_excel(result_deposit_path + "cleaned_comments.xlsx", sheet_name=0)
    # 加载停用词表
    stopwords_path = data_source_path + 'cn_stopwords.txt'
    stopwords = set()
    with open(stopwords_path, 'r', encoding='utf-8') as f:
        for line in f:
            stopwords.add(line.strip())

    # 添加自定义停用词
    stopwords.update(["喝", "买", "非常", '东方', '树叶', '饮料', '没有', '很', ' ', '好', '都', '不', '喜欢', '购买', '特别', '一直','夏天'])

    # 对评论文本进行分词并去除停用词
    def segment(text):
        text = re.sub(r'\s+', ' ', text)  # 去除多余的空格
        words = jieba.cut(text)
        words = [word for word in words if word not in stopwords]
        return ' '.join(words)

    # 对所有评论文本进行分词并去除停用词
    df['Segmented Comment'] = df['Cleaned Comment'].apply(segment)
    # 词频统计
    words = []
    for text in df['Segmented Comment']:
        words += text.split()
    word_counts = pd.Series(words).value_counts()
    # 输出词频最高的前10个词语
    print(word_counts.head(10))

    # 生成词云图并保存到本地
    backgroud_Image = np.array(Image.open(data_source_path + 'backgrand.jpg'))
    font_path = data_source_path + 'simhei.ttf'  # 设置中文字体
    wc = WordCloud(background_color="white", font_path=font_path, mask=backgroud_Image, max_words=2000, width=800,
                   height=400)
    wc.generate_from_frequencies(word_counts)
    wc.to_file(result_deposit_path + "word_cloud.png")

    # 保存分词后的结果
    df.to_excel(result_deposit_path + "segmented_comments.xlsx", sheet_name="Sheet1", index=False)

    # 保存词频统计结果
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Word Counts"
    ws.append(["Word", "Count"])
    for word, count in word_counts.items():
        ws.append([word, count])
    wb.save(result_deposit_path + "word_counts.xlsx")


if __name__ == "__main__":
    data_source_path = './Data/'
    result_deposit_path = './Result/'
    # data_cleaning(data_source_path, result_deposit_path)
    data_segmentation(data_source_path, result_deposit_path)
