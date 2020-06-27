package com.bkunzh.excel;

import com.bkunzh.excel.util.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Ignore;
import org.junit.Test;

import java.io.*;

public class ExcelReadTest {
    String DIR = "F:\\0study-effective\\java-excel\\excel";

    @Test
    public void hssf_xls() throws IOException {
//        InputStream in = new FileInputStream(new File(DIR, "t2003.xls"));
        InputStream in = new FileInputStream(new File(DIR, "t1.xls"));
        Workbook workbook = new HSSFWorkbook(in);
        Sheet sheet0 = workbook.getSheetAt(0);
        Row rowTitle = sheet0.getRow(0);
        Cell cell = rowTitle.getCell(0);
//        System.out.println(cell.getStringCellValue());
        System.out.println(ExcelUtil.getCellValue(cell, workbook));
        in.close();
    }

    @Test
    public void xssf_xlsx() throws IOException {
        InputStream in = new FileInputStream(new File(DIR, "t2007.xlsx"));
        Workbook workbook = new XSSFWorkbook(in);
        Sheet sheet0 = workbook.getSheetAt(0);
        Row rowTitle = sheet0.getRow(0);
        Cell cell = rowTitle.getCell(0);
        System.out.println(cell.getStringCellValue());
        in.close();
    }

    @Ignore
    @Test
    public void xssf_xls() throws IOException {
        // OLE2NotOfficeXmlFileException: HSSF instead of XSSF
        InputStream in = new FileInputStream(new File(DIR, "t2003.xls"));
        Workbook workbook = new XSSFWorkbook(in);
        Sheet sheet0 = workbook.getSheetAt(0);
        Row rowTitle = sheet0.getRow(0);
        Cell cell = rowTitle.getCell(0);
        System.out.println(cell.getStringCellValue());
        in.close();
    }

    // 单元格多种数据类型：包括字符串、数字、布尔、日期、公式等
    @Test
    public void cell_type() throws IOException {
        InputStream in = new FileInputStream(new File(DIR, "t2007.xlsx"));
        Workbook workbook = new XSSFWorkbook(in);
        Sheet sheet0 = workbook.getSheetAt(0);
        Row rowTitle = sheet0.getRow(0);
        for (int cellNum = 0; cellNum < rowTitle.getPhysicalNumberOfCells(); cellNum++) {
            System.out.printf("[" + cellNum + "]");
            System.out.printf(rowTitle.getCell(cellNum).getStringCellValue());
            System.out.printf("  ");
        }
        System.out.println();

        for (int rowNum = 1; rowNum < sheet0.getPhysicalNumberOfRows(); rowNum++) {
            Row rowData = sheet0.getRow(rowNum);
            for (int cellNum = 0; cellNum < rowTitle.getPhysicalNumberOfCells(); cellNum++) {
                System.out.printf("[" + cellNum + "]");
                Cell cell = rowData.getCell(cellNum);
                System.out.printf(ExcelUtil.getCellValue(cell, workbook) + "  ");
            }
            System.out.println();
        }

        in.close();
    }


}
