package top.yumbo.test.excel.exportDemo.simpleDemo;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
public class ExcelSimpleExportDemo {
    public static void main(String[] args) {
        String str="Hello {0}，我是 {1},今年{2}岁";
        str = str.replace("{0}", "CSDN");
        str = str.replace("{1}", "小猪");
        str = str.replace("{2}", "12");
        System.out.println(str);
    }
}
