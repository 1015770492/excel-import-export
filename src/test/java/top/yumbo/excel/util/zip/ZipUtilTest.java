package top.yumbo.excel.util.zip;

import org.junit.Test;


public class ZipUtilTest {


    @Test
    public void unzip() throws Exception {
        long start = System.currentTimeMillis();
        String zipFilePath = "src/test/resources/excel/ImportForQuarter_big.xlsx";
        String desDirectory = "src/test/resources/excel/";
        ZipUtil.unzip(zipFilePath, desDirectory);
        long end = System.currentTimeMillis();
        System.out.println("解压耗时" + (end - start) + "ms");
    }
}
