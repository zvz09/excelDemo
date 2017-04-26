package jxl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.junit.Test;

/**
 * Created by lilipo on 2017/4/26.
 */
public class jxlDemo {


    @Test
    public void WriteExcel() throws IOException, WriteException {

        // 1、创建工作簿(WritableWorkbook)对象，打开excel文件，若文件不存在，则创建文件
        WritableWorkbook writeBook = Workbook.createWorkbook(new File("D://write.xls"));

        // 2、新建工作表(sheet)对象，并声明其属于第几页
        WritableSheet firstSheet = writeBook.createSheet("第一个工作簿", 1);// 第一个参数为工作簿的名称，第二个参数为页数
        WritableSheet secondSheet = writeBook.createSheet("第二个工作簿", 0);

        // 3、创建单元格(Label)对象，
        Label label1 = new Label(1, 2, "test1");// 第一个参数指定单元格的列数、第二个参数指定单元格的行数，第三个指定写的字符串内容
        firstSheet.addCell(label1);
        Label label2 = new Label(1, 2, "test2");
        secondSheet.addCell(label2);

        // 4、打开流，开始写文件
        writeBook.write();

        // 5、关闭流
        writeBook.close();
    }

    @Test
    public void ReadExcel() throws IOException, BiffException {
        // 1、构造excel文件输入流对象
        String sFilePath = "D://write.xls";
        InputStream is = new FileInputStream(sFilePath);
        // 2、声明工作簿对象
        Workbook rwb = Workbook.getWorkbook(is);
        // 3、获得工作簿的个数,对应于一个excel中的工作表个数
        rwb.getNumberOfSheets();

        Sheet oFirstSheet = rwb.getSheet(0);// 使用索引形式获取第一个工作表，也可以使用rwb.getSheet(sheetName);其中sheetName表示的是工作表的名称
        //System.out.println("工作表名称：" + oFirstSheet.getName());
        int rows = oFirstSheet.getRows();//获取工作表中的总行数
        int columns = oFirstSheet.getColumns();//获取工作表中的总列数
        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < columns; j++) {
                Cell oCell = oFirstSheet.getCell(j, i);//需要注意的是这里的getCell方法的参数，第一个是指定第几列，第二个参数才是指定第几行
                System.out.println(oCell.getContents() + "\r\n");
            }
        }

    }
}
