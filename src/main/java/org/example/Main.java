package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class Main {
    private static final String FILE_NAME = "SG_Final_22.189.xlsx";
    private static final String DECIMAL_TO_DECIMAL = "([0-9]+\\.[0-9]+) To ([0-9]+\\.[0-9]+)";
    private static final String STRING_DASH_DECIMAL_DASH_DECIMAL = "^[A-Z]+-[0-9]\\.[0-9]+-[0-9]\\.[0-9]+$";

    public static void main(String[] args) throws Exception {
        Workbook workbook = readExcelFileFromResourceFolder();
        DataFormatter dataFormatter = new DataFormatter();
        ArrayList<Integer> sheetPositions = getSheetPositions(workbook);
        System.out.println(sheetPositions.size());
        ArrayList<HashMap<String, HashMap<String, Integer>>> sheetData = new ArrayList<>();
        for (int i = 0; i < 1; i++) {
            Integer sheetPosition = sheetPositions.get(i);
            System.out.println("Sheet Position: " + sheetPosition);
            Sheet sheet = workbook.getSheetAt(sheetPosition);
            System.out.println("Sheet Name: " + workbook.getSheetName(sheetPosition));
            Iterator<Row> rowIterator = sheet.rowIterator();
            HashMap<String, HashMap<String, Integer>> sheetDataMap = new HashMap<>();
            ArrayList<String> headers = new ArrayList<>();
            while (rowIterator.hasNext()) {
                // check if it is 1st row, if yes then ignore first column of the row and set the column name as key and column number as value and add it to the map
                Row row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        // ignore the column 0
                        if (cell.getColumnIndex() == 0) {
                            continue;
                        }

                        String cellValue = dataFormatter.formatCellValue(cell);

                        // ignore the column if it is empty
                        if (cellValue.isEmpty()) {
                            continue;
                        } else {
                            headers.add(cellValue);
                        }
                    }
                    // else if check if the row is not empty,
                    // if yes then break the loop
                } else {
                    if (isRowEmpty(row)) {
                        break;
                    }
                    // get the value of the first column and check if it is not empty, if yes then create a set of map and add it to the list
                    Cell firstCell = row.getCell(0);
                    String firstCellValue = dataFormatter.formatCellValue(firstCell);
                    if (!firstCellValue.isEmpty()) {
                        // the value of the first column is not empty, so create a map and add it to the list
                        HashMap<String, Integer> rowMap = new HashMap<>();
                        for (int j = 0; j < headers.size(); j++) {
                            Cell cell = row.getCell(j + 1);
                            String cellValue = dataFormatter.formatCellValue(cell);
                            if (cellValue.isEmpty()) {
                                rowMap.put(headers.get(j), 0);
                            } else {
                                rowMap.put(headers.get(j), Integer.parseInt(cellValue));
                            }
                        }
                        sheetDataMap.put(firstCellValue, rowMap);
                    }
                }
            }
            System.out.println(headers);
            sheetData.add(sheetDataMap);
        }
        sheetData.forEach(System.out::println);
        // for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
        //     Sheet sheetAt = workbook.getSheetAt(i);
        //     for (Row row : sheetAt) {
        //         Iterator<Cell> cellIterator = row.cellIterator();
        //         while (cellIterator.hasNext()) {
        //             Cell cell = cellIterator.next();
        //             String cellValue = dataFormatter.formatCellValue(cell);
        //         }
        //     }
        // }
        workbook.close();
    }

    private static boolean isRowEmpty(Row row) {
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != CellType.BLANK)
                return false;
        }
        return true;
    }

    private static boolean isCellMerged(Cell cell, ArrayList<CellRangeAddress> mergedRegions) {
        for (CellRangeAddress mergedRegion : mergedRegions) {
            if (mergedRegion.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                return true;
            }
        }
        return false;
    }

    private static ArrayList<CellRangeAddress> getMergedRegions(Sheet sheet) {
        ArrayList<CellRangeAddress> mergedRegions = new ArrayList<>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            mergedRegions.add(sheet.getMergedRegion(i));
        }
        return mergedRegions;
    }

    private static Workbook readExcelFileFromResourceFolder() throws Exception {
        ClassLoader classLoader = Main.class.getClassLoader();
        File file = new File(Objects.requireNonNull(classLoader.getResource(Main.FILE_NAME)).getFile());
        FileInputStream fileInputStream = new FileInputStream(file);
        return new XSSFWorkbook(fileInputStream);
    }

    private static ArrayList<Integer> getSheetPositions(Workbook workbook) {
        int sheetCount = 0;
        ArrayList<Integer> sheetPositions = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheetAt = workbook.getSheetAt(i);
            String sheetName = sheetAt.getSheetName();
            if (sheetName.matches(DECIMAL_TO_DECIMAL) || sheetName.matches(STRING_DASH_DECIMAL_DASH_DECIMAL)) {
                // System.out.println(sheetName);
                sheetCount++;
                sheetPositions.add(i);
            }
        }
        int numberOfSheets = workbook.getNumberOfSheets();
        // System.out.println("Total sheet count: " + numberOfSheets);
        // System.out.println("Total sheet count with not required data: " + (numberOfSheets - sheetCount));
        // System.out.println("Total sheet count with required data: " + sheetCount);
        // System.out.println("Sheet positions with required data: " + sheetPositions);
        // for (Integer sheetPosition : sheetPositions) {
        //     System.out.println("Sheet position: " + sheetPosition);
        // }
        return sheetPositions;
    }

    private static void parseExcelSheet(Sheet sheet) {
        int rows = sheet.getPhysicalNumberOfRows();
        int cols = 0; // No of columns
        int tmp = 0;

        for (int i = 0; i < 10 || i < rows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                if (tmp > cols) cols = tmp;
            }
        }

        String[][] data = new String[rows][cols];
        System.out.println("rows: " + rows + " cols: " + cols);

    }
}
