import os
from xlsxwriter import Workbook
from re import findall
from datetime import datetime
from sys import argv
from nltk import stem

arguments = argv[1:]
count_args = len(arguments)

if count_args < 1:
    print("You need to specify filename as argument: python test_stops.py c:\\filename.txt")
    exit(1)

input_filename = argv[1]

# open and read input text file exclude stopwords

input_file = open(input_filename, "r", encoding="utf-8")
input_text = input_file.read()
ru_stemmer = stem.snowball.RussianStemmer(True)
stop_words_list = ru_stemmer.stopwords
separate_text_words = findall(r'\w+', input_text)
separate_text_words_lower = []

for i in separate_text_words:
    if i.lower() not in stop_words_list:
        separate_text_words_lower.append(i)

# stem result word list

stemmer_list = []

for word in separate_text_words_lower:
    stemmer_list.append((word, ru_stemmer.stem(word)))

# form full list of stemmed words and count it

counted_stemmer_list = []
count_stem = 0
is_exist = False
full_word_list = ''

for word, stem_word in stemmer_list:
    for t in stemmer_list:
        if stem_word in t:
            count_stem += 1
            full_word_list += word + ', '

    for t in counted_stemmer_list:
        if stem_word in t:
            is_exist = True

    if not is_exist: counted_stemmer_list.append((full_word_list.strip(', '), stem_word, count_stem))
    count_stem = 0
    is_exist = False
    full_word_list = ''

counted_stemmer_list.sort(key=lambda x: x[2], reverse=True)

# write data in Excel file

cur_dir_path = os.getcwd()
workbook = Workbook(cur_dir_path + '\\result_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx')
worksheet = workbook.add_worksheet()

col1_name = 'List of word'
col2_name = 'Stem word'
col3_name = 'Frequency'

worksheet.write(0, 0, col1_name)
worksheet.write(0, 1, col2_name)
worksheet.write(0, 2, col3_name)
row = 1

for word, stem_word, count in counted_stemmer_list:
    worksheet.write(row, 0, word)
    worksheet.write(row, 1, stem_word)
    worksheet.write(row, 2, count)
    row += 1

workbook.close()
print("Result wrote in " + workbook.filename)
