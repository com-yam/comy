# -*- coding: utf-8 -*-
import sys
import os
import math
import re
import xlrd
import xlwt

START = 5 # C5
do_error_check = True

#--------------1個1個のSumaC12_1c_XXX_YYYYY.chi（のうちのCの列）のデータとかを保存するやつ-------------------
class CHI:
    def __init__(self, filename): #filenameを受け取る
        self.c_column = []
        self.filename = filename
        # ファイルが実際に存在するかチェック
        if not os.path.isfile(self.filename):
            print("ファイルが存在しませんでした : " + self.filename)
            exit(-1)
        # self.filenameからchi番号とdata番号を取り出す
        self.setNums()
        # ファイルからcの列を読み込む
        self.readingColumnC()

    # self.filenameからchi番号とdata番号を取り出す。
    def setNums(self):
        # foundにはカッコでくくったそれぞれの場所にある文字列がタプルで入る
        found = re.findall("SumaC12_1c_([0-9]+)_([0-9]+).chi", self.filename)[0]
        # foundにはchi番号とdata番号の二つが入るはずなのでチェック
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
        for line in open(self.filename): # filenameのファイルを各行lineとして読み込み
            linedata = line.strip().split(" ") # 各行を空白で分割
            if len(linedata) < 3: # Cの列がないときは
                element = "" # 空白を入れる
            if len(linedata) >= 3: # Cの列が存在するときは
                element = linedata[2] # その値を入れる
            self.c_column.append(element) # 実際の格納処理
        return


#-------------------------xlsファイルをメモリ上にコピーして書き込む準備するやつ----------------------------
def CopyBook(source_filename):
    source_book = xlrd.open_workbook(source_filename) # コピー元
    output_book = xlwt.Workbook() # コピー先（output_book.save("コピー先ファイルネーム")を呼ぶまで書き込みは行われない）
    # コピー
    for sheet_num in range(source_book.nsheets):
        source_sheet = source_book.sheet_by_index(sheet_num)
        new_sheet = output_book.add_sheet(source_sheet.name, cell_overwrite_ok=True)
        for row in range(source_sheet.nrows):
            for col in range(source_sheet.ncols):
                new_sheet.write(row, col, source_sheet.cell(row,col).value)
    # シートが1枚しかなかったらもう一枚追加
    if source_book.nsheets == 1: output_book.add_sheet("sheet2", cell_overwrite_ok=True)
    # コピーしたoutput_bookを返す
    return output_book



#-------------------------SumaC12_1c_TEMPフォルダに対して実行するやつ----------------------------
def SumaSuma(folder_path):
    # program_start_pathとtemplate_path変数はこの関数の外側で定義されている
    global program_start_path
    global template_path
    # 関数実行時にいたフォルダに戻ってこられるように絶対パスを記憶
    function_start_path = os.path.abspath(os.getcwd())

    # まずは渡されたフォルダ名がSumaC12_1c_TEMPの形式になっているかと，存在しているかを確認
    re_suma_folder = re.compile(".*SumaC12_1c_([0-9]+)/?$")
    if not re_suma_folder.match(folder_path) or not os.path.isdir(folder_path):
        # 一致しないか存在しないため終了
        return
    # 一致して存在していたのでTEMPの部分を取得
    temp_value = re_suma_folder.findall(folder_path)[0]
    # このフォルダ内に移動
    os.chdir(folder_path)
    # imageフォルダの存在を確認
    if not os.path.isdir("image"):
        # 存在しないので元のフォルダに戻って終了
        os.chdir(function_start_path)
        return
    # 確認できたのでimageフォルダに移動し全データを読み込む
    os.chdir("image")
    chis = []
    re_chi = re.compile("SumaC12_1c_[0-9]+_[0-9]+.chi")
    for filename in os.listdir("."):
        if not re_chi.match(filename):
            continue # 実験データでなければ次のファイルへ
        chis.append(CHI(filename))
    # 最後にdata_numでソート
    chis = sorted(chis, key=lambda e:e.data_num)
    # 最大行数
    max_row = max([len(e.c_column) for e in chis])

    # エラーチェック
    if do_error_check: # 上の方でdo_error_check = False にしたらエラーチェックはスキップされる
        isError = False
        if len(chis) == 0:
            print("実験データが見つからない")
            exit(-1)

        pre_num = 0
        for i,e in enumerate(chis):
            if i == 0:
                if e.data_num != 1:
                    print("SumaC12_1c_TEMP_XXX.chiのXXXが1から始まってない : 最初=>" + e.filename)
                    isError = True
            else:
                if chis[i-1].data_num + 1 != e.data_num:
                    print("番号が続いてない : 前回=>" + chis[i-1].filename + "  今回=>" + e.filename)
                    isError = True

        first_ex_num = chis[0].ex_num
        for e in chis[1:]:
            if e.ex_num != first_ex_num:
                print("実験番号が違う : " + e.filename)
                isError = True
        if isError:
            exit(-1)

    # imageフォルダがあるフォルダに戻ってくる
    os.chdir("..")
    # まとめたデータを書き込む
    # まずはテンプレートになるxlsファイルの内容をメモリ上にコピーし，2枚目（0から数えるので1番）のシートを持ってくる
    book = CopyBook(template_path)
    sheet2 = book.get_sheet(1)
    # sheet2の1Bに1c_TEMPを，3BにTEMPを入力
    sheet2.write(0,1, "1c_" + temp_value)
    sheet2.write(2,1, int(temp_value))
    # 次に各chiデータをメモリ上に書き込む
    for i,ex in enumerate(chis): # 各chiファイルについて（先頭からi=0,1,2,...番）
        col = i+1 # 列を一つ右にずらして（0番目のchiはBの列に，1番目はCの列に・・・）
        for row in range(START-1, len(ex.c_column)): # C5(START)～C末尾までを
            sheet2.write(row+1, col, float(ex.c_column[row])) # 1行下の位置（row+1）に書き込む（C5は6行目に，C6は7行目に・・・）
    # コピー先のファイル名を生成
    out_name = os.path.abspath("SumaC12_1c_" + temp_value + ".xls")
    # このファイルに書き込む
    book.save(out_name)
    print("できました => " + os.path.relpath(out_name, program_start_path)) # プログラム起動時のパスとの相対パスを表示（この関数の実行時のパスではない）
    os.chdir(function_start_path) # 関数実行時のパスに戻る



#***********************************実際の処理部分****************************************
if len(sys.argv) != 3:
    print("使い方  =>  $ comy.py root_dir_path template_file.path")
    print("\"root_dir_path\"には\"SumaC12_1c_XXXフォルダが入っているフォルダ\"のパスを渡してください")
    print("template_file.pathには元になるSumaC12_1c_xxx.xlsのパスを渡してください")
    exit(0)
root_path = os.path.abspath(sys.argv[1])
if not os.path.isdir(root_path):
    print("\"" + root_path + "\"フォルダが見つかりません")
    exit(-1)
template_path = os.path.abspath(sys.argv[2])
if not os.path.isfile(template_path):
    print("\"" + template_path + "\"ファイルが見つかりません")
    exit(-1)

# 途中で作業ディレクトリを変えるので，最初に起動時のパスを絶対パスで記憶
program_start_path = os.path.abspath(os.getcwd())

# コマンドライン引数で渡されたroot_pathに移動
os.chdir(root_path)
# そのフォルダ中にある各ファイル・フォルダ名について
for suma_path in os.listdir("."):
    SumaSuma(suma_path)


