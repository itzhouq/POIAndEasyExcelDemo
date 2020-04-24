package cn.itzhouq;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * 计算公式
 */
public class Formula {
    String PATH = "F:\\workspaces\\workspaces_idea\\POIAndEasyExcel\\itzhouq-poi";

    @Test
    public void testForMula() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "/公式.xls");
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);

        // 拿到计算公式
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);

        // 输出单元格的内容
        int cellType = cell.getCellType();
        switch (cellType) {
            case (Cell.CELL_TYPE_FORMULA): // 公式
                String formula = cell.getCellFormula();
                System.out.println(formula); // SUM(A2:A4)

                // 计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue); // 1188.0
                break;
        }
    }
}
