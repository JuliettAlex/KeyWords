from stop_words import get_stop_words
from xlsxwriter import Workbook
from re import findall
from operator import itemgetter
from datetime import datetime
from collections import Counter, OrderedDict
from sys import argv
import os

arguments = argv[1:]
count_args = len(arguments)

if count_args < 1:
    print("You need to specify filename as argument: python test_stops.py c:/filename.txt")
    exit(1)

input_filename = argv[1]

input_file = open(input_filename, "r", encoding="utf-8")
input_text = input_file.read()

stop_words_list = get_stop_words('ru')

separate_text_words = findall(r'\w+', input_text)

separate_text_words_lower = []

for i in separate_text_words:
    separate_text_words_lower.append(i.lower())

freq_text_words = OrderedDict(sorted(Counter(separate_text_words_lower).items(), key=itemgetter(1), reverse=True))

cur_dir_path = os.getcwd()
workbook = Workbook(cur_dir_path + '\\result_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx')
worksheet = workbook.add_worksheet()

col1_name = 'Word'
col2_name = 'Frequency'

worksheet.write(0, 0, col1_name)
worksheet.write(0, 1, col2_name)
row = 1

for key in freq_text_words:
    if key not in stop_words_list:
        worksheet.write(row, 0, key)
        worksheet.write(row, 1, freq_text_words[key])
        row += 1

workbook.close()
print("Result wrote in " + workbook.filename)
