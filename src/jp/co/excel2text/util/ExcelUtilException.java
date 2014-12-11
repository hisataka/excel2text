package jp.co.excel2text.util;

/**
 * ExcelUtilを利用した際に発生し得る例外クラス
 *
 */
public class ExcelUtilException extends Exception {
    private static final long serialVersionUID = 1L;
    private int code;
    
    public ExcelUtilException(int code, String message) {
        super(message);
        this.code = code;
    }
    
    public int getCode() {
        return code;
    }
}
