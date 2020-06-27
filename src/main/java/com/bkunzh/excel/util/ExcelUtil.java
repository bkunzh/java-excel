package com.bkunzh.excel.util;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

import java.text.SimpleDateFormat;

public class ExcelUtil {
    /**
     * 获取单元格值，包括字符串、数字、布尔、日期、公式等
     * @param cell
     * @param workbook
     * @return
     */
    public static Object getCellValue(Cell cell, Workbook workbook) {
        Object rs = null;
        FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        switch (cell.getCellTypeEnum()) {
            case STRING:
                rs = cell.getStringCellValue();
                break;
            case NUMERIC: // 数字或日期
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    rs = sdf.format(cell.getDateCellValue());
                } else {
                    rs = cell.getNumericCellValue();
                }
                break;
            case BOOLEAN:
                rs = cell.getBooleanCellValue();
                break;
            case FORMULA: // 公式
//                String cellFormula = cell.getCellFormula();
                CellValue cellValue = formulaEvaluator.evaluate(cell);
                rs = cellValue.formatAsString();
                break;
        }
        return rs;
    }
}
