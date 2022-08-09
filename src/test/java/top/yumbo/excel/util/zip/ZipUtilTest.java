package top.yumbo.excel.util.zip;

import org.junit.Test;


public class ZipUtilTest {


    @Test
    public void unzip() throws Exception {
        long start = System.currentTimeMillis();
        String zipFilePath = "src/test/resources/excel/AvailableAssetInformation.xlsx";
        String desDirectory = zipFilePath.split("\\.")[0];
        ZipUtil.unzip(zipFilePath, desDirectory);
        long end = System.currentTimeMillis();
        System.out.println("解压耗时" + (end - start) + "ms");
    }

    @Test
    public void unzip_small() throws Exception {
        long start = System.currentTimeMillis();
        String zipFilePath = "src/test/resources/excel/ImportForQuarter.xlsx";
        String desDirectory = zipFilePath.split("\\.")[0];
        ZipUtil.unzip(zipFilePath, desDirectory);
        long end = System.currentTimeMillis();
        System.out.println("解压耗时" + (end - start) + "ms");
    }

    @Test
    public void t(){
        int[][] ints = new int[3000][200];
        for (int i = 0; i < ints.length; i++) {
            for (int i1 = 0; i1 < ints[0].length; i1++) {
                System.out.print(ints[i][i1]+"-");
            }
            System.out.println();
        }
    }
}
