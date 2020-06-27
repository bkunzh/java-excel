package com.bkunzh.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelWriteTest {
    String DIR = "F:\\0study-effective\\java-excel\\excel\\write";

    @Test
    public void t2003() throws IOException {
        // 创建一个工作薄
        Workbook workbook = new HSSFWorkbook();
        // 工作表
        Sheet sheet = workbook.createSheet("bk");
        Row row0 = sheet.createRow(0);
        Cell cell0 = row0.createCell(0);
//        cell0.setCellValue(new Date());
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        cell0.setCellValue(sdf.format(new Date()));

        row0.createCell(1).setCellValue("aa");
        row0.createCell(2).setCellValue(12);
        row0.createCell(3).setCellValue(true);

        Row row1 = sheet.createRow(3);
        row1.createCell(2).setCellValue("uuu");

        OutputStream out = new FileOutputStream(new File(DIR, "t1.xls"));
        workbook.write(out);

        out.close();
    }

    @Test
    public void t2007() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("bk");
        Row row0 = sheet.createRow(0);
        Cell cell0 = row0.createCell(0);
//        cell0.setCellValue(new Date());
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        cell0.setCellValue(sdf.format(new Date()));
        row0.createCell(1).setCellValue("aa");
        row0.createCell(2).setCellValue(12);
        row0.createCell(3).setCellValue(true);

        Row row1 = sheet.createRow(3);
        row1.createCell(2).setCellValue("uuu");

        OutputStream out = new FileOutputStream(new File(DIR, "t1.xlsx"));
        workbook.write(out);

        out.close();
    }

    @Test
    public void xls_bigData() throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
//        final int N = 100 * 1000;
        final int N = 65536;
        for (int i = 0; i < N; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                row.createCell(j).setCellValue(i + "," + j);
            }
        }
        OutputStream out = new FileOutputStream(new File(DIR, "t2.xls"));
        workbook.write(out);
        out.close();
    }

    @Test
    public void xlsx_bigData() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
//        final int N = 100 * 1000;
        final int N = 65536;
        for (int i = 0; i < N; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                row.createCell(j).setCellValue(i + "," + j);
            }
        }
        OutputStream out = new FileOutputStream(new File(DIR, "t2.xlsx"));
        workbook.write(out);
        out.close();
    }

    @Test
    public void xlsx_bigData2() throws IOException {
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        final int N = 100 * 1000;
//        final int N = 65536;
        for (int i = 0; i < N; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                row.createCell(j).setCellValue(i + "," + j);
            }
        }
        OutputStream out = new FileOutputStream(new File(DIR, "t3.xlsx"));
        workbook.write(out);
        ((SXSSFWorkbook) workbook).dispose();
        out.close();
    }
}
