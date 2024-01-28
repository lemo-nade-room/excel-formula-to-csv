# Excelの式をCSVに変換する

## 使い方

```shell
deno run --allow-read https://raw.githubusercontent.com/lemo-nade-room/excel-formula-to-csv/0.0.1/main.ts --row 25 --colmun 6 --sheet <sheetname> <excel_file.xlsx>  
```

## 例

コマンド

```shell
deno run --allow-read https://raw.githubusercontent.com/lemo-nade-room/excel-formula-to-csv/0.0.1/main.ts -r 5 -c 6 sample.xlsx
```

結果

```csv
"行番号","A","B","C","D","E","F"
"1","1","2","A1+B1","","",""
"2","","","","","",""
"3","","","","","",""
"4","","","","","",""
"5","","","LOG(C4,2)","","",""
```

