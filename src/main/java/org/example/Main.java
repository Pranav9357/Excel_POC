package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.*;
import java.util.regex.Pattern;

public class Main {
    private static final String FILE_NAME = "SG_Final_22.189.xlsx";
    private static final String DECIMAL_TO_DECIMAL = "([0-9]+\\.[0-9]+) To ([0-9]+\\.[0-9]+)";
    private static final String STRING_DASH_DECIMAL_DASH_DECIMAL = "^[A-Z]+-[0-9]\\.[0-9]+-[0-9]\\.[0-9]+$";

    public static void main(String[] args) {
        String fileName = args[0];
        String table = args[1];
        excelToCsv(fileName, table);
    }

    private static void parseData(List<XSSFSheet> list, String table) {
        List<Map<String, List<Map<String, String>>>> sheetTable = new ArrayList<>();
        for (Sheet sheet : list) {
            List<Map<String, String>> sheetItems = new ArrayList<>();
            List<String> header = new ArrayList<>();
            String pointer = "";

            if (table.equals("1")) {
                parseTableOne(header, pointer, sheet, sheetItems);
            } else if (table.equals("2")) {
                try {
                    parseTableTwo(header, pointer, sheet, sheetItems);
                } catch (NullPointerException e) {
                    e.printStackTrace();
                }
            } else if (table.equals("3")) {
                parseTableThree(header, pointer, sheet, sheetItems);
            }

            Map<String, List<Map<String, String>>> sheetMap = new HashMap<>();
            sheetMap.put(sheet.getSheetName(), sheetItems);
            sheetTable.add(sheetMap);
        }
        convertToCsv(sheetTable, table);
    }

    private static void convertToCsv(List<Map<String, List<Map<String, String>>>> sheetTable, String table) {
        List<List<String>> csvTable = new ArrayList<>();
        if (table.equals("1")) {
            List<String> header = Arrays.asList("Sheet", "Pointer", "Clarity", "Color", "Price", "Font");
            for (Map<String, List<Map<String, String>>> sheet : sheetTable) {
                for (Map.Entry<String, List<Map<String, String>>> entry : sheet.entrySet()) {
                    for (Map<String, String> item : entry.getValue()) {
                        List<String> row = new ArrayList<>();
                        row.add(item.get("Sheet"));
                        row.add(item.get("Pointer"));
                        row.add(item.get("Clarity"));
                        row.add(item.get("Color"));
                        row.add(item.get("Price"));
                        row.add(item.get("Font"));
                        csvTable.add(row);
                    }
                }
            }
            saveCsv(csvTable, header, table);
        } else if (table.equals("2")) {
            List<String> header = Arrays.asList("Sheet", "Pointer", "Clarity", "Cut", "Color", "Florescence", "Font", "Value", "Value_Color");
            for (Map<String, List<Map<String, String>>> sheet : sheetTable) {
                for (Map.Entry<String, List<Map<String, String>>> entry : sheet.entrySet()) {
                    for (Map<String, String> item : entry.getValue()) {
                        List<String> row = new ArrayList<>();
                        row.add(item.get("Sheet"));
                        row.add(item.get("Pointer"));
                        row.add(item.get("Clarity"));
                        row.add(item.get("Cut"));
                        row.add(item.get("Color"));
                        row.add(item.get("Florescence"));
                        row.add(item.get("Font"));
                        row.add(item.get("Value"));
                        row.add(item.get("Value_Color"));
                        csvTable.add(row);
                    }
                }
            }
            saveCsv(csvTable, header, table);
        } else if (table.equals("3")) {
            List<String> header = Arrays.asList("Pointer", "Clarity", "Cut", "Color", "Florescence", "Font", "Value", "Value_Color");
            for (Map<String, List<Map<String, String>>> sheet : sheetTable) {
                for (Map.Entry<String, List<Map<String, String>>> entry : sheet.entrySet()) {
                    for (Map<String, String> item : entry.getValue()) {
                        List<String> row = new ArrayList<>();
                        row.add(item.get("Pointer"));
                        row.add(item.get("Clarity"));
                        row.add(item.get("Cut"));
                        row.add(item.get("Color"));
                        row.add(item.get("Florescence"));
                        row.add(item.get("Font"));
                        row.add(item.get("Value"));
                        row.add(item.get("Value_Color"));
                        csvTable.add(row);
                    }
                }
            }
            saveCsv(csvTable, header, table);
        }
    }

    private static void saveCsv(List<List<String>> csvTable, List<String> header, String table) {
        try {
            FileWriter csvWriter = new FileWriter(String.format("table_%s.csv", table));
            csvWriter.append(String.join(",", header));
            csvWriter.append("\n");
            for (List<String> row : csvTable) {
                csvWriter.append(String.join(",", row));
                csvWriter.append("\n");
            }
            csvWriter.flush();
            csvWriter.close();
            System.out.println("Data inserted successfully");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void parseTableOne(List<String> header, String pointer, Sheet sheet, List<Map<String, String>> sheetItems) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell != null) {
                    header.add(cell.getStringCellValue());
                }
            }
            pointer = header.get(0);
            header = header.subList(1, header.size());
            break;
        }
        for (Row row : sheet) {
            if (row.getCell(0) == null) {
                break;
            }
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.BLANK) {
                    continue;
                }
                if (cell.getRowIndex() == 0 || cell.getColumnIndex() == 0) {
                    continue;
                }
                CellStyle style = cell.getCellStyle();
                Font font = sheet.getWorkbook().getFontAt(style.getFontIndexAsInt());
                String fontStyle = "";
                if (font.getBold()) {
                    fontStyle = "bold".toUpperCase();
                } else if (font.getItalic()) {
                    fontStyle = "italic".toUpperCase();
                } else {
                    fontStyle = "normal".toUpperCase();
                }
                // price value if cell is numeric or string
                String price = cell.getCellType() == CellType.NUMERIC ? String.valueOf(cell.getNumericCellValue()) : cell.getStringCellValue();
                // print row number
//                System.out.println("row number: " + cell.getRowIndex());
                Map<String, String> rowMap = new HashMap<>();
                rowMap.put("Sheet", sheet.getSheetName());
                rowMap.put("Pointer", pointer);
                rowMap.put("Clarity", sheet.getRow(0).getCell(cell.getColumnIndex()).getStringCellValue());
                rowMap.put("Color", row.getCell(0).getStringCellValue());
                rowMap.put("Price", price);
                rowMap.put("Font", fontStyle);
                sheetItems.add(rowMap);
            }
        }
    }

    private static String getPointerIndex(Cell cell, List<Integer> pointerHeaderIndex, Sheet sheet) {
        Row sheetPointerHeaderRow = sheet.getRow(pointerHeaderIndex.get(0));
//        System.out.println("sheetPointerHeaderRow: " + pointerHeaderIndex.get(0));
        int cellColumnIndex = cell.getColumnIndex();
        if (sheetPointerHeaderRow.getCell(cellColumnIndex).getCellType() == CellType.BLANK) {
            while (sheetPointerHeaderRow.getCell(cellColumnIndex).getCellType() == CellType.BLANK) {
                cellColumnIndex -= 1;
            }
        }
        System.out.println(sheetPointerHeaderRow.getCell(cellColumnIndex).toString());
        return sheetPointerHeaderRow.getCell(cellColumnIndex).toString();
    }

    private static String getClarityIndex(Cell cell, int clarityHeaderRowIndex, Sheet sheet) {
        Row sheetClarityHeaderRow = sheet.getRow(clarityHeaderRowIndex);
        int cellColumnIndex = cell.getColumnIndex();
        if (sheetClarityHeaderRow.getCell(cellColumnIndex).getCellType() == CellType.BLANK) {
            while (sheetClarityHeaderRow.getCell(cellColumnIndex).getCellType() == CellType.BLANK) {
                cellColumnIndex -= 1;
            }
        }
        return sheetClarityHeaderRow.getCell(cellColumnIndex).toString();
    }

    private static String getCutIndex(Cell cell, int cutHeaderRowIndex, Sheet sheet) {
        Row sheetCutHeaderRow = sheet.getRow(cutHeaderRowIndex);
        int cellColumnIndex = cell.getColumnIndex();
        return sheetCutHeaderRow.getCell(cellColumnIndex).toString();
    }

    private static String getFlorescenceIndex(Cell cell, List<Integer> florescenceHeaderIndex, Sheet sheet) {
        int cellRowIndex = cell.getRowIndex();
        String sheetFlorescenceValue = sheet.getRow(cellRowIndex).getCell(florescenceHeaderIndex.get(1)).toString();
        return sheetFlorescenceValue.toUpperCase();
    }

    private static String getColorIndex(Cell cell, List<Integer> colorHeaderIndex, Sheet sheet) {
        int cellRowIndex = cell.getRowIndex();
        boolean isSheetColorBlank = sheet.getRow(cellRowIndex).getCell(colorHeaderIndex.get(1)).getCellType() == CellType.BLANK;
        String sheetFlorescenceValue = sheet.getRow(cellRowIndex).getCell(colorHeaderIndex.get(1)+1).toString();
        String sheetColorValue = "";
        if (isSheetColorBlank && sheetFlorescenceValue.equals("None")) {
            while (sheet.getRow(cellRowIndex).getCell(colorHeaderIndex.get(1)).getCellType() == CellType.BLANK) {
                cellRowIndex += 1;
                sheetColorValue = sheet.getRow(cellRowIndex).getCell(colorHeaderIndex.get(1)).toString();
            }
        } else if (isSheetColorBlank) {
            while (sheet.getRow(cellRowIndex).getCell(colorHeaderIndex.get(1)).getCellType() == CellType.BLANK) {
                cellRowIndex -= 1;
                sheetColorValue = sheet.getRow(cellRowIndex).getCell(colorHeaderIndex.get(1)).toString();
            }
        } else {
            sheetColorValue = sheet.getRow(cellRowIndex).getCell(colorHeaderIndex.get(1)).toString();
        }
        return sheetColorValue;
    }

    private static String getCellColor(Cell cell, Sheet sheet) {
        CellStyle style = cell.getCellStyle();
        if (style.getFillBackgroundColorColor() != null) {
            System.out.println("Row: " + cell.getRowIndex() + " : " + cell.getColumnIndex());
            System.out.println("Color: " + XSSFColor.toXSSFColor(style.getFillForegroundColorColor()).getARGBHex());
            XSSFColor argbColor = XSSFColor.toXSSFColor(style.getFillForegroundColorColor());
            if (argbColor.getARGBHex().equals("FF000000")) {
                return "black".toUpperCase();
            } else if (argbColor.getARGBHex().equals("FF00B050")) {
                return "green".toUpperCase();
            } else if (argbColor.getARGBHex().equals("FFFF0000")) {
                return "red".toUpperCase();
            } else if (argbColor.getARGBHex().equals("FF0000FF")) {
                return "blue".toUpperCase();
            } else {
                return "white".toUpperCase();
            }
        } else {
            return "white".toUpperCase();
        }

    }

    private static String getFontStyle(Cell cell, Sheet sheet) {
        CellStyle style = cell.getCellStyle();
        Font font = sheet.getWorkbook().getFontAt(style.getFontIndexAsInt());
        String fontStyle = "";
        if (font.getBold()) {
            fontStyle = "bold".toUpperCase();
        } else if (font.getItalic()) {
            fontStyle = "italic".toUpperCase();
        } else {
            fontStyle = "normal".toUpperCase();
        }
        return fontStyle;
    }

    private static void parseTableTwo(List<String> header, String pointer, Sheet sheet, List<Map<String, String>> sheetItems) {
        // clarityHeader is set of clarity headers
        List<Integer> pointerHeaderIndex = new ArrayList<>();
        String pointerHeader = "";
        Set<String> clarityHeader = new HashSet<>();
        int clarityHeaderRowIndex = 0;
        Set<String> cutHeader = new HashSet<>();
        int cutHeaderRowIndex = 0;
        List<Integer> colorHeaderIndex = new ArrayList<>();
        List<Integer> florescenceHeaderIndex = new ArrayList<>();
        List<String> florescenceHeader = new ArrayList<>();
        florescenceHeader.add("None");
        florescenceHeader.add("Faint");
        florescenceHeader.add("Medium");
        florescenceHeader.add("Strong");
        for (Row row : sheet) {
            if (row.getCell(0) == null) {
                continue;
            }
            if (row.getCell(0).toString().equals("Range =>")) {
                for (Cell cell : row) {
                    if (cell.toString().isEmpty()) {
                        continue;
                    }
                    pointerHeaderIndex.add(cell.getRow().getRowNum());
                    System.out.println("pointer index" + pointerHeaderIndex);
                    pointerHeader = sheet.getRow(cell.getRow().getRowNum()).getCell(2).toString();
                }
            } else {
                continue;
            }
        }

        // Loop through the sheet and get the data
        for (Row row : sheet) {
            if (row.getCell(0) == null) {
                continue;
            }
            if (row.getCell(0).toString().equals("Clarity =>")) {
                for (Cell cell : row) {
                    if (cell.toString().isEmpty()) {
                        continue;
                    }
                    clarityHeader.add(cell.toString());
                    clarityHeaderRowIndex = cell.getRow().getRowNum();
                }
            } else if (row.getCell(0).toString().equals("Cut =>")) {
                for (Cell cell : row) {
                    if (cell.toString().isEmpty()) {
                        continue;
                    }
                    cutHeader.add(cell.toString());
                    cutHeaderRowIndex = cell.getRow().getRowNum();
                }
            } else if (row.getCell(0).toString().equals("Color")) {
                for (Cell cell : row) {
                    if (cell.toString().isEmpty()) {
                        continue;
                    }
                    if (cell.toString().equals("Color")) {
                        colorHeaderIndex.add(cell.getRow().getRowNum());
                        colorHeaderIndex.add(cell.getColumnIndex());
                        System.out.println("color index" + colorHeaderIndex);
                    } else if (cell.toString().equals("Florescence")) {
                        florescenceHeaderIndex.add(cell.getRow().getRowNum());
                        florescenceHeaderIndex.add(cell.getColumnIndex());
                    }
                }
            } else {
                continue;
            }
        }

        for (int i = colorHeaderIndex.get(0) + 1; i < pointerHeaderIndex.get(2); i++) {
            for (int j = 2; j < sheet.getRow(i).getLastCellNum(); j++) {
                System.out.println("Row: " + i + " : " + j);
                Cell cell = sheet.getRow(i).getCell(j);
                if (cell == null) {
                    continue;
                }
                if (cell.getCellType() == CellType.STRING) {
                    continue;
                }
                System.out.println("Cell: " + cell);
                String pointerIndex = getPointerIndex(cell, pointerHeaderIndex, sheet);
                String clarityIndex = getClarityIndex(cell, clarityHeaderRowIndex, sheet);
                String cutIndex = getCutIndex(cell, cutHeaderRowIndex, sheet);
                String florescenceIndex = getFlorescenceIndex(cell, florescenceHeaderIndex, sheet);
                String colorIndex = getColorIndex(cell, colorHeaderIndex, sheet);
                String cellColor = getCellColor(cell, sheet);
                String fontStyle = getFontStyle(cell, sheet);

                Map<String, String> rowDict = new HashMap<>();
                rowDict.put("Sheet", sheet.getSheetName());
                rowDict.put("Pointer", pointerIndex);
                rowDict.put("Clarity", clarityIndex);
                rowDict.put("Cut", cutIndex);
                rowDict.put("Color", colorIndex);
                rowDict.put("Florescence", florescenceIndex);
                rowDict.put("Font", fontStyle);
                rowDict.put("Value", cell.getCellType() == CellType.BLANK ? "NONE" : cell.toString());
                rowDict.put("Value_Color", cellColor);
//                System.out.println(rowDict);
                sheetItems.add(rowDict);
            }
        }
    }

    private static void parseTableThree(List<String> header, String pointer, Sheet sheet, List<Map<String, String>> sheetItems) {
        // clarityHeader is set of clarity headers
        List<Integer> pointerHeaderIndex = new ArrayList<>();
        int lastPointer = 0;
        String pointerHeader = "";
        Set<String> clarityHeader = new HashSet<>();
        int clarityHeaderRowIndex = 0;
        Set<String> cutHeader = new HashSet<>();
        int cutHeaderRowIndex = 0;
        List<Integer> colorHeaderIndex = new ArrayList<>();
        List<Integer> florescenceHeaderIndex = new ArrayList<>();
        List<String> florescenceHeader = new ArrayList<>();
        florescenceHeader.add("None");
        florescenceHeader.add("Faint");
        florescenceHeader.add("Medium");
        florescenceHeader.add("Strong");
        for (Row row : sheet) {
            if (row.getCell(0) == null) {
                continue;
            }
            System.out.println(row.getCell(0).toString().equals("Range =>"));
            if (row.getCell(0).toString().equals("Range =>")) {
                for (Cell cell : row) {
                    if (cell.toString().isEmpty()) {
                        continue;
                    }
                    pointerHeaderIndex.add(cell.getRow().getRowNum());
                    pointerHeader = sheet.getRow(cell.getRow().getRowNum()).getCell(2).toString();
                }
            } else if(row.getCell(0).toString().equals("J++")){
                for (Cell cell : row) {
                    if (cell.toString().isEmpty()) {
                        continue;
                    }
                    lastPointer = cell.getRow().getRowNum();
                    System.out.println("last pointer" + lastPointer);
                }
            }
        }

        // Loop through the sheet and get the data
        for (Row row : sheet) {
            if (row.getCell(0) == null) {
                continue;
            }
            if (row.getCell(0).toString().equals("Clarity =>")) {
                for (Cell cell : row) {
                    if (cell.toString().isEmpty()) {
                        continue;
                    }
                    clarityHeader.add(cell.toString());
                    clarityHeaderRowIndex = cell.getRow().getRowNum();
                }
            } else if (row.getCell(0).toString().equals("Cut =>")) {
                for (Cell cell : row) {
                    if (cell.toString().isEmpty()) {
                        continue;
                    }
                    cutHeader.add(cell.toString());
                    cutHeaderRowIndex = cell.getRow().getRowNum();
                }
            } else if (row.getCell(0).toString().equals("Color")) {
                for (Cell cell : row) {
                    if (cell.toString().isEmpty()) {
                        continue;
                    }
                    if (cell.toString().equals("Color")) {
                        colorHeaderIndex.add(cell.getRow().getRowNum());
                        colorHeaderIndex.add(cell.getColumnIndex());
                    } else if (cell.toString().equals("Florescence")) {
                        florescenceHeaderIndex.add(cell.getRow().getRowNum());
                        florescenceHeaderIndex.add(cell.getColumnIndex());
                    }
                }
            } else {
                continue;
            }
        }

        for (int i = colorHeaderIndex.get(2) + 1; i <= lastPointer; i++) {
            for (int j = 2; j < sheet.getRow(i).getLastCellNum(); j++) {
                System.out.println("Row: " + i + " : " + j);
                Cell cell = sheet.getRow(i).getCell(j);
                if (cell == null) {
                    continue;
                }
                if (cell.getCellType() == CellType.STRING) {
                    continue;
                }
                System.out.println("Cell: " + cell.toString());
                String pointerIndex = getPointerIndex(cell, pointerHeaderIndex, sheet);
                String clarityIndex = getClarityIndex(cell, clarityHeaderRowIndex, sheet);
                String cutIndex = getCutIndex(cell, cutHeaderRowIndex, sheet);
                String florescenceIndex = getFlorescenceIndex(cell, florescenceHeaderIndex, sheet);
                String colorIndex = getColorIndex(cell, colorHeaderIndex, sheet);
                String cellColor = getCellColor(cell, sheet);
                String fontStyle = getFontStyle(cell, sheet);

                Map<String, String> rowDict = new HashMap<>();
                rowDict.put("Pointer", pointerIndex);
                rowDict.put("Clarity", clarityIndex);
                rowDict.put("Cut", cutIndex);
                rowDict.put("Color", colorIndex);
                rowDict.put("Florescence", florescenceIndex);
                rowDict.put("Font", fontStyle);
                rowDict.put("Value", cell.getCellType() == CellType.BLANK ? "NONE" : cell.toString());
                rowDict.put("Value_Color", cellColor);
                System.out.println(rowDict);
                sheetItems.add(rowDict);
            }
        }
    }

    private static boolean isCellMerged(Cell cell, ArrayList<CellRangeAddress> mergedRegions) {
        for (CellRangeAddress mergedRegion : mergedRegions) {
            if (mergedRegion.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                return true;
            }
        }
        return false;
    }

    private static void excelToCsv(String filePath, String table) {
        try {
            ClassLoader classLoader = Main.class.getClassLoader();
            File file = new File(Objects.requireNonNull(classLoader.getResource(filePath)).getFile());
            FileInputStream fileInputStream = new FileInputStream(file);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            List<XSSFSheet> sheets = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                String sheetName = workbook.getSheetName(i);
                if (Pattern.matches(DECIMAL_TO_DECIMAL, sheetName) || Pattern.matches(STRING_DASH_DECIMAL_DASH_DECIMAL, sheetName)) {
                    sheets.add(workbook.getSheet(sheetName));
                }
            }
            System.out.println("Total number of sheets: " + sheets.size());
            // only keep 1 sheet for testing
//            sheets = sheets.subList(0, 5);
            parseData(sheets, table);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

