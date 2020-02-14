import xlrd
import random
import subprocess

kanji_data = 'kanji_data.xlsx'

wb = xlrd.open_workbook(kanji_data)
sheet = wb.sheet_by_name('Sheet1')
row = sheet.nrows

question = 10

lists = []
for i in range(row):
    lists.append(i)
question_number = random.sample(lists,question)

kanji_lists = []
for i in question_number:
    kanji_lists.append(sheet.row_values(i))

kanji_view = 'kanji_view.txt'

with open(kanji_view, mode='w', encoding='utf-8') as f:
    for i in range(question):
        yomi = kanji_lists[i][1]
        kaki = kanji_lists[i][2]
        f.write(str(i+1)   + "." + " "*(4-len(str(i+1))) + kanji_lists[i][1] + " "*(60-len(yomi)*2-len(kaki)*2) + kanji_lists[i][2] + '\n'*5)

subprocess.Popen([r'C:\WINDOWS\system32\notepad.exe', r'kanji_view'])
