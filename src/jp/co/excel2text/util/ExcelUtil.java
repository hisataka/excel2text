package jp.co.excel2text.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * POI利用を前提としたExcel操作ユーティリティクラス
 * 
 */
public class ExcelUtil {
    // ワークブックオブジェクト
    private Workbook workbook;
    
    /**
     * コンストラクタ
     * 
     * @param fileName 対象となるExcelブックのファイル名付きのフルパス
     * @throws ExcelUtilException 
     */
    public ExcelUtil(String fileName) throws ExcelUtilException {
        try {
            this.workbook = WorkbookFactory.create(new FileInputStream(fileName));
        } catch (InvalidFormatException e) {
            throw new ExcelUtilException(0, "Excelファイルのフォーマットを読み込むことが出来ませんでした。");
        } catch (IOException e) {
            throw new ExcelUtilException(1, "Excelファイルの読み込み中に予期せぬIOエラーが発生しました。");
        }
    }

    /**
     * Excelのシート名を受け取り、該当シートの内容をStringのList構造で返す
     * 
     * @param sheetName 対象となるExcelのシート名
     * @return List<List<String>>に変換した値
     */
    public List<List<String>> sheet2StringList(String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        return sheet2StringList(sheet);
    }

    /**
     * Excelのシートオブジェクトを受け取り、該当シートの内容をStringのList構造で返す
     * 
     * @param sheetName 対象となるExcelのシート名
     * @return List<List<String>>に変換した値
     */
    public List<List<String>> sheet2StringList(Sheet sheet) {
        List<List<String>> ret = new ArrayList<List<String>>();
        for(int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i ++) {
            List<String> row = new ArrayList<String>();
            Row r = sheet.getRow(i);
            if(r == null || r.getFirstCellNum() < 0 || r.getLastCellNum() < 0) {
                continue;
            }
            for(int j = r.getFirstCellNum(); j <= r.getLastCellNum(); j ++) {
                Cell c = r.getCell(j);
                if(c == null) {
                    continue;
                }
                row.add(cell2String(c));
            }
            ret.add(row);
        }
        return ret;
    }
    
    /**
     * Excelのセルオブジェクトを受け取り、Stringで返す
     * 
     * @param cell 取得対象のCellオブジェクト
     * @return Stringに変換した値
     */
    public String cell2String(Cell cell) {
        String ret;
        switch(cell.getCellType()) {
        case Cell.CELL_TYPE_BLANK  :
            ret = "";
            break;
        case Cell.CELL_TYPE_STRING:
            ret = cell.getRichStringCellValue().getString();
            break;
        case Cell.CELL_TYPE_BOOLEAN:
            ret = (cell.getBooleanCellValue()) ? "TRUE" : "FALSE";
            break;
        case Cell.CELL_TYPE_NUMERIC:
            // 日付・整数・少数の判別を行う
            if (DateUtil.isCellDateFormatted(cell)) {       //日付
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
                ret = sdf.format(cell.getDateCellValue());
            } else {  // 数値
                ret = String.valueOf(cell.getNumericCellValue());
                if(ret.endsWith(".0")) {
                    ret = ret.substring(0, ret.length() - 2);
                }
            }
            break;
        case Cell.CELL_TYPE_FORMULA:
            Workbook wb = cell.getSheet().getWorkbook();
            CreationHelper crateHelper = wb.getCreationHelper();
            FormulaEvaluator evaluator = crateHelper.createFormulaEvaluator();
            ret = cell2String(evaluator.evaluateInCell(cell));
            break;
        default:
            ret = "";
            break;
        }
        return ret;
    }
}
