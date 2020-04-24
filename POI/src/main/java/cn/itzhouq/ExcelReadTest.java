package cn.itzhouq;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFCellUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

/**
 * 读操作
 */
public class ExcelReadTest {
    String PATH = "F:\\workspaces\\workspaces_idea\\POIAndEasyExcel\\itzhouq-poi";

    @Test
    public void testRead03() throws IOException {
        // 获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "/人员统计表03.xls");
        // 1. 创建一个工作簿，使用excel能操作的，代码都能操作
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        // 2. 得到表
        Sheet sheet = workbook.getSheetAt(0);
        // 3. 得到行
        Row row = sheet.getRow(0);
        // 4. 得到列
        Cell cell = row.getCell(0);
        System.out.println(cell.getStringCellValue()); // 今日新增人员
        Cell cell1 = row.getCell(1);
        System.out.println(cell1.getNumericCellValue()); // 666.0
        fileInputStream.close();
    }

    @Test
    public void testRead07() throws IOException {
        // 获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "/人员统计表07.xlsx");
        // 1. 创建一个工作簿，使用excel能操作的，代码都能操作
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        // 2. 得到表
        Sheet sheet = workbook.getSheetAt(0);
        // 3. 得到行
        Row row = sheet.getRow(0);
        // 4. 得到列
        Cell cell = row.getCell(0);
        System.out.println(cell.getStringCellValue()); // 今日新增人员
        Cell cell1 = row.getCell(1);
        System.out.println(cell1.getNumericCellValue()); // 666.0
        fileInputStream.close();
    }

    @Test
    public void testCellType() throws IOException {
        // 获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "/明细表.xls");
        // 创建一个工作簿
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        // 获取标题内容
        Row rowTitle = sheet.getRow(0);
        if (rowTitle != null) {
            // 重点
            int rowCount = rowTitle.getPhysicalNumberOfCells(); // 获取列的数量
            for (int cellNum = 0; cellNum < rowCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null) {
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        // 获取表中内容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null) {
                // 读取列
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("[" + (rowNum + 1) + "-" + (cellNum + 1) + "]");

                    Cell cell = rowData.getCell(cellNum);
                    // 匹配列的数据类型
                    if (cell != null) {
                        int cellType = cell.getCellType();
                        String cellValue = "";
                        switch (cellType) {
                            case HSSFCell.CELL_TYPE_STRING: // 字符串
                                System.out.print("【String】");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN: // 布尔
                                System.out.print("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK: // 空
                                System.out.print("【BLANK】");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC: // 数字
                                System.out.print("【UMERIC】");
                                if (HSSFDateUtil.isCellDateFormatted(cell)) { // 日期
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString();
                                } else {
                                    // 不是日期格式，防止数字过长
                                    System.out.print("【装换为字符串输出】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print("【数据类型错误】");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        fileInputStream.close();
    }
}
