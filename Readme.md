# これは何
TSVファイルに含まれたレコードをExcelへ書き込むツールです

# 使い方
1. Buildします

1. config.jsonを実行ファイルと同ディレクトリに配置して下さい

1. 引数に読み込むtsvファイルとexcelファイルを与え実行します

#  ``config.json``について

config/config.jsonにサンプルがありますので参考にして下さい。

|プロパティ|役割|
|----|----|
|splitChar(必須)|TSVファイルの区切り文字を指定します (例：「\t」)|
|targetSheet(必須)|excelファイルの編集先シート番号を指定します|
|identifier(必須)|TSVファイルのヘッダ名とエクセルファイルの列の組み合わせを指定します。この列の値をキーにして更新行を検索します。|
|filledColumn|excel側に常に値が入っている列があれば指定して下さい|
|endOfColumn(必須)|excelのどの列まで探索するかを指定します|
|mapping|TSVファイルのヘッダ名とエクセルファイルの列の組み合わせを指定します。複数対複数も可能です。|

mappingにてTSVの複数の列をexcelのセルに書き込む場合は、``mapping[].SplitChar``を区切り文字として書き込みます。