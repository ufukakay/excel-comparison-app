package com.rpa;

import java.io.File;

import java.io.FileInputStream;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcelFile {

    File file;
    FileInputStream inputStream;
    Workbook Workbook;
    Sheet Sheet;

    public void setExcelFile(String filePath, String fileName, String sheetName) throws IOException {

        file = new File(filePath + "\\" + fileName);

        inputStream = new FileInputStream(file);

        Workbook = null;

        String fileExtensionName = fileName.substring(fileName.indexOf("."));


        if (fileExtensionName.equals(".xlsx")) {

            Workbook = new XSSFWorkbook(inputStream);
        } else if (fileExtensionName.equals(".xls")) {

            Workbook = new HSSFWorkbook(inputStream);

        }

        Sheet = Workbook.getSheet(sheetName);
    }


    public List<Row> getDataRow() {

        List<Row> rows = new ArrayList<Row>();

        int rowCount = Sheet.getLastRowNum() - Sheet.getFirstRowNum();
        for (int i = 0; i < rowCount + 1; i++) {

            rows.add(Sheet.getRow(i));

        }
        return rows;
    }


}