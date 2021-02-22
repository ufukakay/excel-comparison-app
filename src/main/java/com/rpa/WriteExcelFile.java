package com.rpa;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class WriteExcelFile {

    private Workbook workbook;
    private Sheet sheet;
    private Row headerRow;
    private Row dataRow;

    public void setWorkbook(String[] columns, String sheetName) throws IOException {

        workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file
        sheet = workbook.createSheet(sheetName); // Create a Sheet
        headerRow = sheet.createRow(0); // Create a Header Row

        // Create cells
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
        }
    }

    public void setDataRow(Employee employee, int rowNum) {

        dataRow = sheet.createRow(rowNum);
        dataRow.createCell(0).setCellValue(employee.getName());
        dataRow.createCell(1).setCellValue(employee.getSurname());
        dataRow.createCell(2).setCellValue(employee.getBirthDate());
        dataRow.createCell(3).setCellValue(employee.getBirthPlace());
        dataRow.createCell(4).setCellValue(employee.getMail());
        dataRow.createCell(5).setCellValue(employee.getPhone());
        dataRow.createCell(6).setCellValue(employee.getStatus());
        dataRow.createCell(7).setCellValue(employee.getWorkStatus());
        dataRow.createCell(8).setCellValue(employee.getUniversity());

    }

    public void writeExcel(String fileName) throws IOException {

        for (int i = 0; i < 9; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(fileName);
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

}



