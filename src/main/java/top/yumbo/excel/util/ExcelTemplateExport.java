package top.yumbo.excel.util;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.InputStream;

/**
 * @author 诗水人间
 * @link 博客:{https://yumbo.blog.csdn.net/}
 * @link github:{https://github.com/1015770492}
 * @link 在线文档:{https://1015770492.github.io/excel-import-export/}
 * @date 2021/9/4 22:42
 * <p>
 * excel模板导出，给excel待填充位置加上字段名称，自动进行模板导出
 * 向模板文件中填入字段`{字段名称}` 即可将数据填入
 */
public class ExcelTemplateExport implements HSSFListener {

    private SSTRecord sstrec;

    public static void main(String[] args) throws Exception {
        String xlsx = "src/test/java/top/yumbo/test/excel/2.xlsx";
        // 文件inputStream
        FileInputStream is = new FileInputStream(xlsx);
        POIFSFileSystem poifs = new POIFSFileSystem(is);

        InputStream din = poifs.createDocumentInputStream("Workbook");
        HSSFRequest req = new HSSFRequest();
        // 为HSSFRequest增加listener
        req.addListenerForAllRecords(new ExcelTemplateExport());
        HSSFEventFactory factory = new HSSFEventFactory();
        // 处理inputstream
        factory.processEvents(req, din);
        // 关闭inputstream
        is.close();
        din.close();
        System.out.println("done.");

    }

    @Override
    public void processRecord(Record record) {
        switch (record.getSid()) {
            // 标记workbook或sheet开始，这里会进行判断
            case BOFRecord.sid:
                BOFRecord bof = (BOFRecord) record;
                if (bof.getType() == bof.TYPE_WORKBOOK) {
                    System.out.println("Encountered workbook");
                    // assigned to the class level member
                } else if (bof.getType() == bof.TYPE_WORKSHEET) {
                    System.out.println("Encountered sheet reference");
                }
                break;
            //处理sheet
            case BoundSheetRecord.sid:
                BoundSheetRecord bsr = (BoundSheetRecord) record;
                System.out.println("New sheet named: " + bsr.getSheetname());
                break;
            //处理行
            case RowRecord.sid:
                RowRecord rowrec = (RowRecord) record;
                System.out.println("行 "
                        + rowrec.getFirstCol() + " last column at " + rowrec.getLastCol());
                break;
            //处理数字单元格
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;
                System.out.println("单元格 " + numrec.getValue()
                        + " at row " + numrec.getRow() + " and column " + numrec.getColumn());
                break;
            // 包含一行中所有文本单元格
            case SSTRecord.sid:
                sstrec = (SSTRecord) record;
                for (int k = 0; k < sstrec.getNumUniqueStrings(); k++) {
                    System.out.println("String table value " + k + " = " + sstrec.getString(k));
                }
                break;
            //处理文本单元格
            case LabelSSTRecord.sid:
                LabelSSTRecord lrec = (LabelSSTRecord) record;
                System.out.println("String cell found with value "
                        + sstrec.getString(lrec.getSSTIndex()));
                break;
        }
    }
}
