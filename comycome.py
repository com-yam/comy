# -*- coding: utf-8 -*-
import sys
import os
import math
import re
#import xlwt

START = 5 # C5
do_error_check = True

# 一個一個のSumaC12_1c_XXX_YYYYY.chi（のうちのCの列）のデータとかを保存するやつ
class Example:
    def __init__(self, filename): #filenameを受け取る
        self.c_column = []
        self.filename = filename
        # ファイルが実際に存在するかチェック
        if not os.path.isfile(self.filename):
            print("ファイルが存在しませんでした : " + self.filename)
            exit(-1)
        # self.filenameからexample番号とdata番号を取り出す
        self.setNums()
        # ファイルからcの列を読み込む
        self.readingColumnC()

    # self.filenameからexample番号とdata番号を取り出す。
    def setNums(self):
        # foundにはカッコでくくったそれぞれの場所にある文字列がタプルで入る
        found = re.findall("SumaC12_1c_([0-9]+)_([0-9]+).chi", self.filename)[0]
        # foundにはexample番号とdata番号の二つが入るはずなのでチェック
        if len(found) != 2 or not found[0].isdigit() or not found[1].isdigit():
            print("変な名前のファイルが見つかりました : " + self.filename)
            exit(-1)
        # チェックを無事通ったら自分の中に格納する
        self.ex_num_str = found[0]
        self.ex_num = int(self.ex_num_str)
        self.data_num_str = found[1]
        self.data_num = int(self.data_num_str)
        return

    # self.filenameファイルからcの列のデータを読み込む
    def readingColumnC(self):
        for line in open(filename): # filenameのファイルを各行lineとして読み込み
            linedata = line.strip().split(" ") # 各行を空白で分割
            if len(linedata) < 3: # Cの列がないときは
                element = "" # 空白を入れる
            if len(linedata) >= 3: # Cの列が存在するときは
                element = linedata[2] # その値を入れる
            self.c_column.append(element) # 実際の格納処理
        return


#------------------------------------実際の処理部分-----------------------------------------------------
if len(sys.argv) != 2:
    # print("使い方  =>  $ comy.py root_dir_path template_file.path")
    print("使い方  =>  $ comy.py root_dir_path")
    # print("\"root_dir_path\"にはSumaC12_1c_XXXフォルダが入ってるフォルダのパスを渡してください")
    print("\"root_dir_path\"にはSumaC12_1c_XXXフォルダのパスを渡してください")
    exit(0)
root_path = os.path.abspath(sys.argv[1])
if not os.path.isdir(root_path):
    print("\"" + root_path + "\"フォルダが見つかりません")
    exit(-1)
#template_path = sys.argv[2]
# if not os.path.isfile(template_path):
#     print("\"" + template_path + "\"ファイルが見つかりません")
#     exit(-1)


# 途中で作業ディレクトリを変えるので，最初に現在のディレクトリを絶対パスで記憶
start_cwd = os.path.abspath(os.getcwd())

# コマンドライン引数で渡されたroot_pathに移動
os.chdir(root_path)
# そのフォルダ中にある各ファイル・フォルダ名について
for suma_path in os.listdir("."):
    # まずはSumaC12_1c_TEMPの形式になっているかを確認
    re_suma_folder = re.compile("^SumaC12_1c_([0-9]+)$")
    if not re_suma_folder.match(suma_path): continue # 一致しないので次のファイル・フォルダへ
    # 一致していた
    temp_value = re_suma_folder.findall(suma_path)[0]
    # imageフォルダの存在を確認
    if not os.path.isdir(suma_path + "/image"): continue # 存在しないので次のファイル・フォルダへ
    # 確認できた
    # SumaC12_1c_TEMP/imageフォルダに移動
    os.chdir(suma_path + "/image")
    # フォルダ中の全データを読み込む
    examples = []
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
    os.chdir("../") # 一つ上のフォルダ（imageフォルダがあるフォルダ）に移動
    out_name = os.path.abspath("alldata_" + temp_value + ".csv")
    # 書き込む
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
    os.chdir(root_path) # 元のroot_pathに戻る
    print("できました => " + os.path.relpath(out_name, start_cwd))

