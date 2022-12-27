package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
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
//                parseTableTwo(header, pointer, sheet, sheetItems);
            } else if (table.equals("3")) {
//                parseTableThree(header, pointer, sheet, sheetItems);
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
            List<String> header = Arrays.asList("Pointer", "Clarity", "Color", "Price", "Font");
            for (Map<String, List<Map<String, String>>> sheet : sheetTable) {
                for (Map.Entry<String, List<Map<String, String>>> entry : sheet.entrySet()) {
                    for (Map<String, String> item : entry.getValue()) {
                        List<String> row = new ArrayList<>();
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
                rowMap.put("Pointer", pointer);
                rowMap.put("Clarity", sheet.getRow(0).getCell(cell.getColumnIndex()).getStringCellValue());
                rowMap.put("Color", row.getCell(0).getStringCellValue());
                rowMap.put("Price", price);
                rowMap.put("Font", fontStyle);
                sheetItems.add(rowMap);
            }
        }
    }


    private static void getRowValues(DataFormatter dataFormatter, ArrayList<String> values, Iterator<Cell> cellIterator) {
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            if (cell.getColumnIndex() != 0) {
                String cellValue = dataFormatter.formatCellValue(cell);
                if (!cellValue.isEmpty()) {
                    values.add(cellValue);
                    // get cell properties like font, color, etc
                    System.out.println(cell.getCellStyle());
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
            sheets = sheets.subList(0, 1);
            parseData(sheets, table);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

