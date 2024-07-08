# LatexToWord_Test

Extract latex files into one MS word file. 

# motivation

Sometime I need to convert latex files into word file.
PDF -> MS WORD conversion is not perfect, and 
the paragraphs are sometimes broken and divided into several parts. 

This tool provides very simple conversion of latex files in one MS WORD file. 
#input macro is extended automatically. 

In the conversion, you can "ignore" dedicated lines starting with words in "config.txt".




# how to use 


windows専用ですが，latexファイルをWordに押し込むプログラム．
texファイルの中にあるinputコマンドを自動で展開する．
inputコマンドの行は，\input{ファイル名}　から始まらないといけない
(コメントされていたり行の途中にあるときは自動展開されない)

０．LatexToWordのプログラムが入ったフォルダを好きなところにおく

１．LatexToWordのフォルダのアドレスバーにcmdを打ち込んでEnterで，
　　コマンドプロンプトを開く
　　（あるいはコマンドプロンプト上でexeがあるフォルダまで来る）

2．コマンドを実行する
LatexToWord <ここに処理したいファイル名，パスを含んでも良い> ignore.txt

例：一つ上のフォルダに入っているsample_texフォルダにあるroot.texを処理する

LatexToWord ..\\sample_tex\\root.tex ignore.txt
(windowsではフォルダの区切りが/ではなくて\\である)


- ignore.txtについて
ignore.txtに記載のあるコマンドから始まる行は，WORDに書き出さない．

デフォルトでは下記が無視されるようになっている．

----- ignore.txt : includes words to skip lines -----
\usepackage
\newcommand

-------------------

