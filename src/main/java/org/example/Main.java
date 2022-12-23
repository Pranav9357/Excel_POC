package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.*;

public class Main {
    private static final String FILE_NAME = "SG_Final_22.189.xlsx";
    private static final String DECIMAL_TO_DECIMAL = "([0-9]+\\.[0-9]+) To ([0-9]+\\.[0-9]+)";
    private static final String STRING_DASH_DECIMAL_DASH_DECIMAL = "^[A-Z]+-[0-9]\\.[0-9]+-[0-9]\\.[0-9]+$";

    private static final String CSV_FILE_PATH = "./SG_Final_22.189.result.csv";

    public static void main(String[] args) throws Exception {
        Workbook workbook = readExcelFileFromResourceFolder();
        DataFormatter dataFormatter = new DataFormatter();
        ArrayList<Integer> sheetPositions = getSheetPositions(workbook);
        System.out.println(sheetPositions.size());
        ArrayList<HashMap<String, HashMap<String, HashMap<String, Integer>>>> sheetData = new ArrayList<>();

        for (int i = 0; i < sheetPositions.size(); i++) {
            Integer sheetPosition = sheetPositions.get(i);
            Sheet sheet = workbook.getSheetAt(sheetPosition);
            Iterator<Row> rowIterator = sheet.rowIterator();
            HashMap<String, HashMap<String, Integer>> sheetDataMap = new HashMap<>();
            ArrayList<String> headers = new ArrayList<>();
            ArrayList<String> values = new ArrayList<>();
            label:
            while (rowIterator.hasNext()) {
                // check if it is 1st row, if yes then ignore first column of the row and set the column name as key and column number as value and add it to the map
                Row row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    Iterator<Cell> cellIterator = row.cellIterator();
                    getRowValues(dataFormatter, headers, cellIterator);
                } else {
                    if (isRowEmpty(row)) {
                        break label;
                    }
                    Cell firstCell = row.getCell(0);
                    String firstCellValue = dataFormatter.formatCellValue(firstCell);
                    if (!firstCellValue.isEmpty()) {
                        Iterator<Cell> cellIterator = row.cellIterator();
                        HashMap<String, Integer> rowMap = new HashMap<>();
                        getRowValues(dataFormatter, values, cellIterator);
                    }
                    HashMap<String, Integer> innerMap = new HashMap<>();
                    for (int j = 0; j < headers.size(); j++) {
                        innerMap.put(headers.get(j), Integer.parseInt(values.get(j)));
                    }
                    sheetDataMap.put(firstCellValue, innerMap);
                }
            }
            HashMap<String, HashMap<String, HashMap<String, Integer>>> sheetMap = new HashMap<>();
            sheetMap.put(workbook.getSheetName(sheetPosition), sheetDataMap);
            sheetData.add(sheetMap);

        }
//        sheetData.forEach(System.out::println);
        File file = new File(CSV_FILE_PATH);
        FileWriter fw = new FileWriter(file);
        BufferedWriter bw = new BufferedWriter(fw);
        List<List<String>> list = new ArrayList<>();
        bw.write("  Sheet. NO.  ---  color  ---  Clarity  ---  value     ");
        for (int l = 0; l < sheetPositions.size(); l++) {
            Set<String> sheetname = sheetData.get(l).keySet() ;
            System.out.println(sheetname);
            String[] Sheetname = sheetname.toArray(new String[sheetname.size()]);
            Set<String> color = sheetData.get(l).get(Sheetname[0]).keySet();
            System.out.println(color);
            String[] Color = color.toArray(new String[color.size()]);
            Set<String> C = sheetData.get(l).get(Sheetname[0]).get(Color[0]).keySet();
            System.out.println(C);
            String[] Clarity = C.toArray(new String[color.size()]);

            for (int j = 0; j < sheetData.get(l).get(Sheetname[0]).size(); j++) {
                for (int k = 0; k < color.size(); k++) {
                    List<String> list1 = new ArrayList<>();
                    list1.add(Sheetname[0]);
                    list1.add((Color[j]));
                    list1.add(Clarity[k]);
                    list1.add((sheetData.get(l).get(Sheetname[0]).get(Color[j]).get(Clarity[k])).toString());
                    list.add(list1);
                }
            }

            bw.newLine();
            for (int p = 0; p < list.size(); p++) {
                bw.newLine();
                for (int j = 0; j < list.get(p).size(); j++)
                    bw.write(list.get(p).get(j) + "  ---  ");
            }
        }
        bw.close();
        fw.close();
//        list.forEach((Consumer<? super List<String>>) System.out::println);
        workbook.close();
    }

    private static void getRowValues(DataFormatter dataFormatter, ArrayList<String> values, Iterator<Cell> cellIterator) {
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            if (cell.getColumnIndex() != 0) {
                String cellValue = dataFormatter.formatCellValue(cell);
                if (!cellValue.isEmpty()) {
                    values.add(cellValue);
                } else {
                    continue;
                }
            }
        }
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
