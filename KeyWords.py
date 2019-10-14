import os
from xlsxwriter import Workbook
from re import findall
from datetime import datetime
from sys import argv
from nltk import stem

arguments = argv[1:]
args_count = len(arguments)

if args_count < 1:
    print("You need to specify filename as argument: python test_stops.py c:\\filename.txt")
    exit(1)

input_filename = argv[1]

# Открываем и читаем файл с исходным текстом

input_file = open(input_filename, "r", encoding="utf-8")
input_text = input_file.read()

# Формируем список из пар (слово, основ слова) исключая стоп слова

ru_stemmer = stem.snowball.RussianStemmer(True)
stop_words_list = ru_stemmer.stopwords
separate_words_list = findall(r'\w+', input_text)

separate_text_words_lower = []
stem_list = []

for word in separate_words_list:
    if word.lower() not in stop_words_list:
        stem_list.append((word, ru_stemmer.stem(word)))

# Выявляем частовстречающие слова и поизводим сопоставление (слова, основа слова, число вхождений)

counted_stem_list = []
stem_count = 0
is_exist = False
full_word_list_text = ''

for word, stem_word in stem_list:
    for stem in stem_list:
        if stem_word in stem:
            stem_count += 1
            full_word_list_text += word + ', '

    for stem in counted_stem_list:
        if stem_word in stem:
            is_exist = True

    if not is_exist:
        counted_stem_list.append((full_word_list_text.strip(', '), stem_word, stem_count))
    stem_count = 0
    is_exist = False
    full_word_list_text = ''

counted_stem_list.sort(key=lambda x: x[2], reverse=True)

# записываем результат в файл Excel

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

for word, stem_word, count in counted_stem_list:
    worksheet.write(row, 0, word)
    worksheet.write(row, 1, stem_word)
    worksheet.write(row, 2, count)
    row += 1

workbook.close()
print("Result wrote in " + workbook.filename)
