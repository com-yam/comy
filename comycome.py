# -*- coding: utf-8 -*-
import sys
import os
import math
import re
#import xlwt

START = 5 # C5
do_error_check = True

if len(sys.argv) != 2:
    # print("使い方  =>  $ comy.py root_dir_path template_file.path")
    print("使い方  =>  $ comy.py root_dir_path")
    # print("\"root_dir_path\"にはSumaC12_1c_XXXフォルダが入ってるフォルダのパスを渡してください")
    print("\"root_dir_path\"にはSumaC12_1c_XXXフォルダのパスを渡してください")
    exit(0)
root_path = sys.argv[1]
if not os.path.isdir(root_path):
    print("\"" + root_path + "\"フォルダが見つかりません")
    exit(-1)
#template_path = sys.argv[2]
# if not os.path.isfile(template_path):
#     print("\"" + template_path + "\"ファイルが見つかりません")
#     exit(-1)


# 途中で作業ディレクトリを変えるので，最初に現在のディレクトリを記憶
start_cwd = os.getcwd()

# 一個一個のexample_aaa_bbb.chi（のうちのCの列）のデータとか
class Example:
    def __init__(self, filename): #filenameを受け取る
        self.c_column = []
        self.filename = filename
        for line in open(filename):
            linedata = line.strip().split(" ")#""で分割
            element = "" # Cの列がないときは空白を書き込む
            if len(linedata) >= 3: # Cの列が存在するときは
                element = linedata[2] # その値を書き込む
            self.c_column.append(element)
        splited_filename = os.path.splitext(filename)[0].split("_") # 拡張子外して"_"で分割
        if len(splited_filename) != 4 or not splited_filename[2].isdigit() or not splited_filename[3].isdigit():
            print("変な名前のファイル : " + filename)
            exit(-1)
        self.ex_num_str = splited_filename[2]
        self.ex_num = int(self.ex_num_str)
        self.data_num_str = splited_filename[3]
        self.data_num = int(self.data_num_str)

# フォルダ中の全exampleデータを読み込み
examples = []
os.chdir(root_path)
re_chi = re.compile("SumaC12_1c_[0-9]+_[0-9]+.chi")
for filename in os.listdir("."):
    if not re_chi.match(filename):
        continue
    examples.append(Example(filename))
# 最後にdata_numでソート
examples = sorted(examples, key=lambda e:e.data_num)
# 最大行数
max_row = max([len(e.c_column) for e in examples])

# エラーチェック
if do_error_check: #do_error_check = fals にしたら実行されない
    isError = False
    if len(examples) == 0:
        print("実験データが見つからない")
        exit(-1)

    pre_num = 0
    for i,e in enumerate(examples):
        if i == 0:
            if e.data_num != 1:
                print("example_aaa_bbbのbbbが1から始まってない : 最初=>" + e.filename)
                isError = True
        else:
            if examples[i-1].data_num + 1 != e.data_num:
                print("番号が続いてない : 前回=>" + examples[i-1].filename + "  今回=>" + e.filename)
                isError = True

    first_ex_num = examples[0].ex_num
    for e in examples[1:]:
        if e.ex_num != first_ex_num:
            print("実験番号が違う : " + e.filename)
            isError = True
    if isError:
        exit(-1)

# まとめる
os.chdir(start_cwd) #start_cwdに移動
out_name = "alldata_" + examples[0].ex_num_str + ".csv"
with open(out_name, mode='w') as out:
    # はじめに1～5行目までを出力　空白
    out.write("\n"*5)
    for i in range(START-1, max_row):
        line = "," # Aの列は空っぽ
        for ex in examples:
            if len(ex.c_column) > i:
                line += ex.c_column[i]
            line += ","
        out.write(line + "\n")

print("できました => " + out_name)

