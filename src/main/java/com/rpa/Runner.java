package com.rpa;


import java.io.IOException;


public class Runner {

    public static void main(String[] args) throws IOException {

        ExcelComparison excelComparison = new ExcelComparison();
        excelComparison.dataComparison();
    }
}
