# comy
使い方
$ comycome.py folder_path template_path.xls
folder_path : SumaC12_1c_TEMPフォルダが入っているフォルダ
template_path.xls : いわゆるSumaC12_1c_xxx.xls

# 環境：
・python3
・xlrd，xlwtライブラリが必要
   $ pip install xlrd xlwt  # これで入ります

# その他
comycome.pyはsample_root中のchiファイルが文字列で書いてあるせいで動きません
サンプルで動作させるときはcomycome_sampleVersion.pyを使ってください
例:
$ comycome_sampleVersion.py sample_root sample_root/SumaC12_1c_xxx.xls

