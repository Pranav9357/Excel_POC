package org.example;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
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

    public static void main(String[] args) throws Exception {



        File file = new File("dataOutput.csv");
        FileWriter fw = new FileWriter(file);
        BufferedWriter bw = new BufferedWriter(fw);

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
                    Cell firstCell = row.getCell    (0);
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

            //output in String array

            ArrayList<ArrayList<String>>  al = new ArrayList<>();

            Set<String> sheetname = sheetData.get(i).keySet();
            String[] Sheetname =sheetname.toArray(new String[sheetname.size()]);
            Set<String> color = sheetData.get(i).get(Sheetname[i]).keySet();
            String[] Color =color.toArray(new String[color.size()]);
            Set<String> C = sheetData.get(i).get(Sheetname[i]).get(Color[i]).keySet();
            String[] Clarity =C.toArray(new String[color.size()]);
            for (int l=0;l<1;l++){
                for (int j=0;j<sheetData.get(i).get(Sheetname[i]).size();j++){
                    for (int k=0;k<color.size();k++){
                        ArrayList<String> list = new ArrayList<>();
                        list.add(Sheetname[l]);
                        list.add((Color[j]));
                        list.add(Clarity[k]);
                        list.add((sheetData.get(i).get(Sheetname[i]).get(Color[i]).get(Clarity[k])).toString());

                        al.add(list);

                    }
                }

                System.out.println(al);
            }
            bw.write("Pointer,Color,Clarity,Values");
            bw.newLine();

            for (int n=0; n< al.size();n++){
                bw.newLine();
                for (int m=0;m< al.get(n).size();m++){

                    bw.write(al.get(n).get(m)+",");
                }
            }
            bw.close();
            fw.close();
            System.out.println("data entered");
        }


//        sheetData.forEach(System.out::println);

        }

    private static void getRowValues(DataFormatter dataFormatter, ArrayList<String> values, Iterator<Cell> cellIterator) {
        while (cellIterator.hasNext()){
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
