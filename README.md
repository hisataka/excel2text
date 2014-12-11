excel2text
======================
Excelのシートデータをテキストに変換するためのツール

使用方法
----------------------
execフォルダに実行用ファイルがあります。

実行方法1  

    java -jar Excel2Text Excelファイル名 対象シート名 出力CSVファイル名
→成功すると「success」と表示されます。

実行方法2（成否の終了コードの返却が必要であればこちらを使用して下さい）  

    excel2text.bat Excelファイル名 対象シート名 出力CSVファイル名
→成功すると終了コードとして0を返します。  
　失敗すると終了コードとして1を返します。

ざっくりとした仕様
----------------------
詳細はソースを参照してください。

■CSV出力仕様  
・文字コード：MS932  
・セパレータ：,  
・引用符："  
・引用符のエスケープ文字："  
・改行コード：CRLF  

■Excel取込み仕様  
POIを使用していますので、  
POIのタイプ判定準拠で処理します。  

・CELL_TYPE_BLANK：""  
・CELL_TYPE_STRING：getRichStringCellValue().getString()  
・CELL_TYPE_BOOOLEAN："TRUE" or "FALSE"  
・CELL_TYPE_NUMERIC：org.apache.poi.ss.usermodel.DateUtil.isCellDateFormattedで日付と判定された  
　　　　　　　　　　　→yyyy/MM/dd HH:mm:ss形式の文字列  
 　　　　　　　　　　org.apache.poi.ss.usermodel.DateUtil.isCellDateFormattedで日付と判定されない  
　　　　　　　　　　　→getNumericCellValue()を文字列へ変換（末尾に".0"があれば除外）  
・CELL_TYPE_FORMULA：再帰的にタイプ判定処理を実施 