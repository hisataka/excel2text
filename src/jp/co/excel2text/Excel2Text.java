package jp.co.excel2text;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;

import au.com.bytecode.opencsv.CSVWriter;
import jp.co.excel2text.util.ExcelUtil;
import jp.co.excel2text.util.ExcelUtilException;


public class Excel2Text {

    
    public static void main(String[] args){
        // 引数処理
        // mode inFileName sheetName outFileName
        if(args.length != 3) {
            System.out.println("usage: Excel2Text inFileName sheetname outFileName");
            return ;
        }
        String fileName = args[0];
        String sheetName =args[1];
        String outFileName = args[2];
        
        // Excelデータの取得
        ExcelUtil excelUtil;
        try {
            excelUtil = new ExcelUtil(fileName);
        } catch (ExcelUtilException e) {
            System.out.println("error : " + e.getMessage());
            return ;
        }
        
        // CSV出力を実施
        CSVWriter writer;
        try {
            // 文字コード：MS932
            // セパレータ：,
            // 引用符："
            // エスケープ："
            // 改行コード：CRLF
            FileOutputStream input = new FileOutputStream(outFileName);
            OutputStreamWriter outWriter = new OutputStreamWriter(input, "MS932");
            writer = new CSVWriter(outWriter, ',', '"', '"', "\r\n");
            List<List<String>> data = excelUtil.sheet2StringList(sheetName);
            for(List<String> row : data) {
                String[] s = (String[])row.toArray(new String[0]);
                writer.writeNext(s);
            }
            writer.flush();
        } catch (IOException e) {
            System.out.println("error : " + e.getMessage());
            return ;
        }
        System.out.println("success");
    }
}
