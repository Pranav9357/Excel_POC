package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.Iterator;

public class Main {

    private static String fileName = "/home/er/Excel_POC/Excel_Read_POC/src/main/resources/SG_Final_22.189.xlsx";

    public static void main(String[] args) throws Exception {
        FileInputStream file = new FileInputStream(fileName);
        Workbook workbook = new XSSFWorkbook(file);
        System.out.println("hello" + workbook.getNumberOfSheets());
        DataFormatter dataFormatter = new DataFormatter();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet shets = workbook.getSheetAt(i);
            System.out.println("Sheet name : " + workbook.getSheetName(i));
            Iterator<Row> sheets = shets.iterator();
            while(sheets.hasNext()) {
                Row row = sheets.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String cellValue = dataFormatter.formatCellValue(cell);
                    System.out.print(cellValue+"\t");
                }
                System.out.println();
            }
        }
        workbook.close();
        file.close();
    }
}