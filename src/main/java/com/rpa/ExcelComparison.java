package com.rpa;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;

import java.io.IOException;
import java.util.List;


public class ExcelComparison {

    String filePath = System.getProperty("user.dir") + "\\InputData";
    String fileName = "Excel1.xlsx";
    String fileNameTwo = "Excel2.xlsx";
    String sheetName = "sheet1";
    ReadExcelFile excelFile;
    WriteExcelFile writeExcelFile;

    private static final Logger log = LogManager.getLogger(ExcelComparison.class);

    public void dataComparison() throws IOException {

        excelFile = new ReadExcelFile();


        excelFile.setExcelFile(filePath, fileName, sheetName);
        List<Row> rowsOne = excelFile.getDataRow();

        excelFile.setExcelFile(filePath, fileNameTwo, sheetName);
        List<Row> rowsTwo = excelFile.getDataRow();

        writeExcelFile = new WriteExcelFile();
        writeExcelFile.setWorkbook(new String[]{"AD", "SOYAD", "DOĞUM TARİHİ", "DOĞUM YERİ", "MAİL", "TELEFON", "DURUM", "ÇALIŞMA DURUMU", "ÜNİVERSİTE"}, "Employee");

        int rowNum = 0;

        for (int i = 1; i < rowsOne.size(); i++) {

            String nameOne = rowsOne.get(i).getCell(0).toString();
            String surnameOne = rowsOne.get(i).getCell(1).toString();

            String nameTwo = rowsTwo.get(i).getCell(0).toString();
            String surnameTwo = rowsTwo.get(i).getCell(1).toString();

            String mailOne = "1. email boş";
            if (rowsOne.get(i).getCell(3) != null) {
                mailOne = rowsOne.get(i).getCell(3).getStringCellValue();

            }


            String mailTwo = "2. email boş";
            if (rowsTwo.get(i).getCell(4) != null) {
                mailTwo = rowsTwo.get(i).getCell(4).getStringCellValue();

            }

            String phoneOne = "1.telefon boş";
            if (rowsOne.get(i).getCell(4) != null) {
                phoneOne = NumberToTextConverter.toText(rowsOne.get(i).getCell(4).getNumericCellValue());

            }

            String phoneTwo = "2.telefon boş";
            if (rowsTwo.get(i).getCell(5) != null) {
                phoneTwo = NumberToTextConverter.toText(rowsTwo.get(i).getCell(5).getNumericCellValue());

            }

            String birthDate = "";
            if (rowsOne.get(i).getCell(2) != null) {
                birthDate = rowsOne.get(i).getCell(2).getStringCellValue();
            }

            String birthPlace = "";
            if (rowsTwo.get(i).getCell(3) != null) {
                birthPlace = rowsTwo.get(i).getCell(3).getStringCellValue();
            }

            String status = "";
            if (rowsOne.get(i).getCell(5) != null) {
                status = rowsOne.get(i).getCell(5).getStringCellValue();
            }

            String workStatus = "";
            if (rowsTwo.get(i).getCell(6) != null) {
                workStatus = rowsTwo.get(i).getCell(6).getStringCellValue();
            }

            String university = "";
            if (rowsTwo.get(i).getCell(7) != null) {
                university = rowsTwo.get(i).getCell(7).getStringCellValue();
            }

            log.info("------------------");
            log.info("1. Excel Tablosu: " + " " + nameOne + " " + surnameOne + " " + "-" + " " + mailOne + " " + "-" + " " + phoneOne);
            log.info("2. Excel Tablosu: " + " " + nameTwo + " " + surnameTwo + " " + "-" + " " + mailTwo + " " + "-" + " " + phoneTwo);


            if ((nameOne.equals(nameTwo)) && (surnameOne.equals(surnameTwo)) && ((mailOne.equals(mailTwo)) || (phoneOne.equals(phoneTwo)))) {
                rowNum++;

                if (!(mailOne.equals(mailTwo))) {
                    mailOne = "";
                }
                if (!(phoneOne.equals(phoneTwo))) {
                    phoneOne = "";
                }


                Employee employee = new Employee(nameOne, surnameOne, birthDate, birthPlace, mailOne, phoneOne, status, workStatus, university);
                writeExcelFile.setDataRow(employee, rowNum);

                log.info("Sonuç: Eşleşiyor, Dosyaya yazdırıldı.");
                log.info("------------------ \n");

            } else {

                String name = "";
                if (!(nameOne.equals(nameTwo))) {
                    name = "Ad eşleşmiyor veya boş |";
                }

                String surName = "";
                if (!(surnameOne.equals(surnameTwo))) {
                    surName = "Soyad eşleşmiyor veya boş |";
                }

                String mail = "";
                if (!(mailOne.equals(mailTwo))) {
                    mail = "Mail eşleşmiyor veya boş |";
                }

                String phone = "";
                if (!(phoneOne.equals(phoneTwo))) {
                    phone = "Telefon eşleşmiyor veya boş";
                }

                log.info(name + " " + surName + " " + mail + " " + phone);
                log.info("------------------ \n");
            }


        }

        String fileOutPath = System.getProperty("user.dir") + "\\Report\\Employees.xlsx";
        writeExcelFile.writeExcel(fileOutPath);
    }

}
