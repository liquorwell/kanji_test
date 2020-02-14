import sys
import tkinter
import xlrd
import random
import subprocess

kanji_data = 'kanji_data.xlsx'
wb = xlrd.open_workbook(kanji_data)
sheet = wb.sheet_by_name('Sheet1')
row = sheet.nrows


def btn_click():
    try:
        start = int(box1.get())
        finish = int(box2.get())
        question = int(box3.get())
        if start > finish:
            label9["text"] = "エラー : 出題範囲を正しい順番で入力してください！"
        elif finish > row:
            label9["text"] = "エラー : 出題範囲が問題番号を超えています！"
        elif 0 < (finish - start + 1) < question:
            label9["text"] = "エラー : 問題数が出題範囲より多いです！"
        else:
            lists = []
            for i in range(start,finish+1):
                lists.append(i)
            question_number = random.sample(lists,question)

            kanji_lists = []
            for i in question_number:
                kanji_lists.append(sheet.row_values(i-1))

            kanji_view = 'kanji_view.txt'

            with open(kanji_view, mode='w', encoding='utf-8') as f:
                for i in range(question):
                    yomi = kanji_lists[i][1]
                    kaki = kanji_lists[i][2]
                    f.write(str(i+1)   + "." + " "*(4-len(str(i+1))) + kanji_lists[i][1] + " "*(60-len(yomi)*2-len(kaki)*2) + kanji_lists[i][2] + '\n'*5)

            subprocess.Popen([r'C:\WINDOWS\system32\notepad.exe', r'kanji_view'])

    except ValueError:
        label9["text"] = "エラー : 数字を入力してください！"


root = tkinter.Tk()
root.title(u"漢字テスト生成マシーン")
root.geometry("400x300")

label1 = tkinter.Label(text=u"問題番号は1番から" + str(row) + "番です。")
label1.place(x=20, y=20)
label2 = tkinter.Label(text=u"出題範囲：")
label2.place(x=20, y=80)
label3 = tkinter.Label(text=u"番～")
label3.place(x=120, y=80)
label4 = tkinter.Label(text=u"番")
label4.place(x=190, y=80)
label5 = tkinter.Label(text=u"問題数：")
label5.place(x=20, y=110)
label6 = tkinter.Label(text=u"問")
label6.place(x=110, y=110)
label7 = tkinter.Label(text=u"※うまく印刷できない場合はフォントサイズや用紙サイズを変更してみてください。")
label7.place(x=20, y=190)
label8 = tkinter.Label(text=u"　( 推奨フォントサイズ:14pt　推奨用紙サイズ:A4 )")
label8.place(x=20, y=210)
label9 = tkinter.Label(text=u"　")
label9.place(x=20, y=150)

box1 = tkinter.Entry(width=4)
box1.place(x=90, y=80)
box2 = tkinter.Entry(width=4)
box2.place(x=160, y=80)
box3 = tkinter.Entry(width=4)
box3.place(x=80, y=110)
box3.insert(0, "10")

button = tkinter.Button(text=u"生成", command=btn_click)
button.place(x=175, y=250)

root.mainloop()
