import os
import pandas as pd
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from textblob import TextBlob
import nltk
from collections import Counter
import re

# 下载 NLTK 所需的资源
nltk.download('punkt')
nltk.download('stopwords')

# 加载和处理数据
def load_and_process_data(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' does not exist.")
        return None

    # 加载 CSV 文件
    data = pd.read_csv(file_path)

    # 打印列名以供调试
    print("Columns in the CSV file:", data.columns.tolist())

    # 检查必要的列是否存在
    required_columns = ['Subject', 'ReceivedTime']
    for col in required_columns:
        if col not in data.columns:
            print(f"Error: Required column '{col}' is missing in the CSV file. Exiting.")
            return None

    # 填充缺失值并清理数据
    data['ReceivedTime'] = pd.to_datetime(data['ReceivedTime'], errors='coerce')
    data['Subject_cleaned'] = data['Subject'].fillna('').apply(lambda x: re.sub(r'[^\w\s]', '', str(x).lower()))

    # 如果需要处理正文内容，但正文列不存在，可以跳过或添加一个空列
    if 'Body' not in data.columns:
        print("Warning: 'Body' column not found in the CSV file. Skipping body processing.")
        data['Body_cleaned'] = ''  # 添加一个空的 Body_cleaned 列
    else:
        data['Body_cleaned'] = data['Body'].fillna('').apply(lambda x: re.sub(r'[^\w\s]', '', str(x).lower()))

    return data

# 每日工作容量分析
def daily_work_capacity(data):
    daily_activity = data.groupby(data['ReceivedTime'].dt.date).size()
    print("Daily Work Capacity Overview:")
    print(daily_activity)

    # 绘制每日活动图表
    daily_activity.plot(kind='bar', figsize=(10, 6))
    plt.title("Daily Work Capacity Overview")
    plt.xlabel('Date')
    plt.ylabel('Number of Emails')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

# 生成词云
def generate_wordcloud(data):
    all_subjects = " ".join(data['Subject_cleaned'])
    wordcloud = WordCloud(width=800, height=400, background_color="white").generate(all_subjects)
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis("off")
    plt.title("Word Cloud of Email Subjects")
    plt.show()

# 情感分析
def sentiment_analysis(data):
    print("Sentiment Analysis of Subjects:")

    # 确保所有值为字符串类型，填充缺失值为空字符串
    data['Subject'] = data['Subject'].fillna("").astype(str)

    sentiment_scores = []
    for subject in data['Subject']:
        analysis = TextBlob(subject)
        sentiment_scores.append(analysis.sentiment.polarity)

    data['Sentiment'] = sentiment_scores
    print(data[['Subject', 'Sentiment']].head())

    # 绘制情感分布图
    plt.hist(data['Sentiment'], bins=20, color='skyblue', edgecolor='black')
    plt.title("Sentiment Distribution of Email Subjects")
    plt.xlabel("Sentiment Polarity")
    plt.ylabel("Frequency")
    plt.show()

# 提取关键词
def extract_top_keywords(data):
    print("Top Keywords from Email Subjects:")

    # 将所有主题拼接成一个字符串
    all_subjects = " ".join(data['Subject_cleaned'])

    # 使用简单的正则表达式分词，替代 NLTK 的 word_tokenize
    tokens = re.split(r'\W+', all_subjects)

    # 移除停用词
    stop_words = set(nltk.corpus.stopwords.words('english'))
    filtered_tokens = [word for word in tokens if word and word.lower() not in stop_words]

    # 统计词频
    word_freq = Counter(filtered_tokens)
    top_keywords = word_freq.most_common(10)
    print("Top 10 Keywords:", top_keywords)

    # 绘制关键词柱状图
    keywords, counts = zip(*top_keywords)
    plt.bar(keywords, counts, color='lightgreen')
    plt.title("Top Keywords in Email Subjects")
    plt.xlabel("Keywords")
    plt.ylabel("Frequency")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()


# 项目贡献分析
def project_contributions(data):
    print("Key Achievements & Project Contributions:")
    project_related_words = ['project', 'meeting', 'report', 'submission', 'license', 'order']

    project_emails = data[data['Subject_cleaned'].str.contains('|'.join(project_related_words), na=False)]
    print(f"Found {len(project_emails)} project-related emails")
    print(project_emails[['Subject', 'Attachments']].head())

# 主函数
def main():
    # 自动检测当前目录中的 CSV 文件
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]

    if not csv_files:
        print("No CSV files found in the current directory.")
        return

    # 如果只有一个文件，自动选择
    if len(csv_files) == 1:
        print(f"Only one CSV file found: {csv_files[0]}. Using it automatically.")
        file_path = csv_files[0]
    else:
        # 多个文件时，让用户选择
        print("Select a CSV file for analysis:")
        for idx, file in enumerate(csv_files):
            print(f"{idx + 1}. {file}")

        try:
            choice = int(input("Enter the number corresponding to the file: ")) - 1
            if choice < 0 or choice >= len(csv_files):
                print("Invalid choice. Exiting.")
                return
        except ValueError:
            print("Invalid input. Please enter a number.")
            return

        file_path = csv_files[choice]

    print(f"Selected file: {file_path}")

    # 加载并处理数据
    data = load_and_process_data(file_path)
    if data is None:
        return

    # 执行分析
    daily_work_capacity(data)
    generate_wordcloud(data)
    sentiment_analysis(data)
    extract_top_keywords(data)
    project_contributions(data)

# 运行主函数
if __name__ == "__main__":
    main()
