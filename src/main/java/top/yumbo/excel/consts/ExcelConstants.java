package top.yumbo.excel.consts;

import org.apache.commons.math3.exception.OutOfRangeException;

import java.util.HashMap;
import java.util.Map;

public class ExcelConstants {
    public static final String SHEET1="Sheet1";
    public static final String XLSX="xlsx";
    public static final String DEFAULT_PASSWORD="123456";
    // 用于26进制
    public static final String[] digits = {
            "0", "1", "2", "3", "4", "5",
            "6", "7", "8", "9", "a", "b",
            "c", "d", "e", "f", "g", "h",
            "i", "j", "k", "l", "m", "n",
            "o", "p"
    };
    public static final Map<Integer, String> intMap = new HashMap<Integer, String>() {
        {
            for (int i = 0; i < 26; i++) {
                put(i + 1, "" + Character.toChars('A' + i)[0]);
            }
        }
    };


}
