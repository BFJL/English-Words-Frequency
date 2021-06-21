
from nltk import word_tokenize, pos_tag
from nltk.corpus import wordnet
from nltk.stem import WordNetLemmatizer
from nltk.corpus import stopwords
import xlwt
import re

# 加载停用词词典
stopset = stopwords.words('english')
# 扩充停用词词典
with open('./stop_word.txt', 'r', encoding='utf-8') as f:
    for word in f.readlines():
        word = word.replace('\n', '')
        if word not in stopset:
            stopset.append(word)


# 获取单词的词性
def get_wordnet_pos(tag):
    if tag.startswith('J'):
        return wordnet.ADJ
    elif tag.startswith('V'):
        return wordnet.VERB
    elif tag.startswith('N'):
        return wordnet.NOUN
    elif tag.startswith('R'):
        return wordnet.ADV
    else:
        return None

# 读取文章预料， 文章与文章之间手动添加 [SEP] 分割符号
word_dict = {}
with open('./english.txt', 'r', encoding='utf-8') as f:
    sentences = f.read().split('[SEP]')

word_count_list = []  # 保存每一篇文章的词频
for sentence in sentences:
    word_count = {}  # 当前文章的词频
    pattern = re.compile(r"[^a-zA-Z']")
    sentence = re.sub(pattern, ' ', sentence)
    tokens = word_tokenize(sentence)
    tagged_sent = pos_tag(tokens)  # 获取单词词性
    wnl = WordNetLemmatizer()
    lemmas_sent = []
    for tag in tagged_sent:  # 单词 词性二元列表
        word_pos = get_wordnet_pos(tag[1]) or wordnet.NOUN
        word = wnl.lemmatize(tag[0].lower(), pos=word_pos)  # 词形还原
        # 判断是否在停用词表中
        if word not in stopset and len(word) > 2:
            word_count[word] = word_count.get(word, 0)
            word_count[word] += 1
    word_count_list.append(word_count)

# 统计全部文章的词频
total_word_count = {}
for word_count in word_count_list:
    for word, count in word_count.items():
        total_word_count[word] = total_word_count.get(word, 0) + count


def statistics(word_count_list, word):
    # 返回word在每一篇文章中的count
    every_text_count = []
    for word_count in word_count_list:
        count = word_count.get(word, 0)
        every_text_count.append(count)
    return every_text_count


# 计算单词的词频分布
lines = []
for word, count in total_word_count.items():
    line = [word, count]
    every_text_count = statistics(word_count_list, word)
    line.extend(every_text_count)
    lines.append(line)


# 创建excel工作表
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('sheet1')

# 设置表头
worksheet.write(0, 0, label='单词')
worksheet.write(0, 1, label='总计频数')
for i in range(len(sentences)):
    label = "第%s篇" % str(i+1)
    worksheet.write(0, i+2, label=label)

# 变量用来循环时控制写入单元格
val = 1
for line in lines:
    for i in range(len(line)):
        worksheet.write(val, i, line[i])
    val += 1
# 保存
workbook.save('word.xlsx')





