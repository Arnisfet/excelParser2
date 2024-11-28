package org.example;

import com.aspose.cells.*;

import java.util.*;

public class SheetParser {
    private Map<String, Integer> headerColumns = new HashMap<>();
    private Map<String, Integer> parsedDocumentMap = new HashMap<>();
    private final Map<String, List<String>> rules = new HashMap<>();
    private Workbook header;
    private Map<Integer, Integer> resultMap = new HashMap<>();
    private String fileUrl;
    private int headerColumnX = -1;
    private int headerRowY = -1;

    private int DataWorkColumnX = -1;
    private int DataWorkRowY = -1;


    public SheetParser() {
        try  {
            header = new Workbook("Header.xlsx");
            Worksheet worksheet = header.getWorksheets().get(0);
            Range range = worksheet.getCells().createRange("A1:Q2");
            for (int row = range.getFirstRow(); row < range.getFirstRow() + range.getRowCount(); row++) {
                for (int col = range.getFirstColumn(); col < range.getFirstColumn() + range.getColumnCount(); col++) {
                    Cell cell;
                    if (!worksheet.getCells().get(row + 1, col).getStringValue().isEmpty())
                        cell = worksheet.getCells().get(row + 1, col);
                    else
                        cell = worksheet.getCells().get(row, col);
                    headerColumns.put(cell.getStringValue().toLowerCase().trim().replaceAll(" +", " "),col);
                }
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }

    void matchingRules() {
        // Define the rules
        rules.put("Банк, предоставивший выписку", List.of("Банк, предоставивший выписку"));
        rules.put("вид (шифр) или ВО", List.of("вид (шифр) или ВО", "вид (шифр)"));

        rules.put("номер документа", List.of("№ док.", "номер"));
        rules.put("Дата совершения операции (dd.mm.yyyy) или дата проводки", List.of("Дата операции", "дата"));


        rules.put("наименование/ФИО", List.of("Наименование  плательщика", "наименование/Ф.И.О."));
        rules.put("ИНН/КИО", List.of("ИНН/КИО"));
        rules.put("КПП", List.of("КПП"));
        rules.put("Номер счета", List.of("номер счета (специального банковского счета)"));

        rules.put("По дебету", List.of("по дебету"));
        rules.put("По кредиту", List.of("по кредиту"));

        rules.put("Назначение платежа", List.of("Назначение платежа"));

        rules.put("номер корреспондентского счета", List.of("номер корреспондентского счета"));
        rules.put("наименование", List.of("наименование"));
        rules.put("БИК", List.of("БИК"));

        Map<String, List<String>> normalizedRules = new HashMap<>();

        rules.forEach((key, list) -> {
            String normalizedKey = key.trim().replaceAll(" +", " ").toLowerCase();

            List<String> normalizedList = list.stream()
                    .map(value -> value.trim().replaceAll(" +", " ").toLowerCase())
                    .toList();

            normalizedRules.put(normalizedKey, normalizedList);
        });

        rules.clear();
        rules.putAll(normalizedRules);
    }

    void parse(String filePath) {
        try {
            Workbook currentWorkbook = new Workbook(filePath);
            WorksheetCollection worksheets = currentWorkbook.getWorksheets();
            fileUrl = "file:///" + filePath.replace("\\", "/");


            for (int i = 0; i < worksheets.getCount(); i++) {
                Worksheet sheet = worksheets.get(i);
                System.out.println("Parsing sheet: " + sheet.getName());
                parseDocHeader(sheet);
                matchRulesWithHeaders();
                matrixCreation(sheet);
            }
            System.out.println("Main parsing completed.");
        } catch (Exception e) {
            throw new RuntimeException("Error parsing Excel file", e);
        }
    }

    void matrixCreation(Worksheet sheet) {
        try {
            // We need to work with the cells of the sheet starting from DataWorkRowY and DataWorkColumnX
            Cells cells = sheet.getCells();

            List<Map<Integer, String>> matrix = new ArrayList<>();

            // Loop over the rows starting from DataWorkRowY
            for (int row = DataWorkRowY; row <= cells.getMaxDataRow(); row++) {
                Map<Integer, String> rowMap = new HashMap<>();

                // Loop over the columns starting from DataWorkColumnX
                for (int col = DataWorkColumnX; col <= cells.getMaxDataColumn(); col++) {
                    // Get the cell value
                    Cell cell = cells.get(row, col);
                    String cellValue = cell.getStringValue().trim();

                    // Get the corresponding header from parsedDocumentMap
                    if (!cellValue.isEmpty())
                        rowMap.put(col, cellValue);
                }

                // Add rowMap to the matrix
                matrix.add(rowMap);
            }

            // Output the created matrix for verification
            System.out.println("Matrix created with " + matrix.size() + " rows.");
//            for (Map<Integer, String> rowMap : matrix) {
//                for (Map.Entry<Integer, String> entry : rowMap.entrySet()) {
//                    System.out.print(entry.getKey() + ": " + entry.getValue() + " | ");
//                }
//                System.out.println();
//            }
            System.out.println();

        } catch (Exception e) {
            throw new RuntimeException("Error while creating matrix", e);
        }
    }

    // Тут просто матчим колонки из документа, который парсим, и колонки хедера, логика не оч читаемая, но ля че делать
    void matchRulesWithHeaders() {
        Set<String> columnSet = new HashSet<>(headerColumns.keySet());
       for (Map.Entry<String, List<String>> entry : rules.entrySet()) {
           String key = entry.getKey();
           List<String> values = entry.getValue();

           parsedDocumentMap.forEach((docKey, docValue) -> {
               if (values.contains(docKey)) {
                   resultMap.put(headerColumns.get(key), docValue);
                    columnSet.remove(key);
               }
           });
       }
       if (resultMap.isEmpty())
           System.out.println("Map has not been matched at all!!!");
        System.out.println("Number of matched columns: " + resultMap.size());
        System.out.println("Unmutched columns: " + columnSet);
    }

    void parseDocHeader(Worksheet sheet) {
        try {
            Cells cells = sheet.getCells();
            boolean headerFound = false;

            outerLoop:
            for (int row = 0; row <= cells.getMaxDataRow(); row++) {
                for (int col = 0; col <= cells.getMaxDataColumn(); col++) {
                    Cell cell = cells.get(row, col);
                    String cellValue = cell.getStringValue().trim().replaceAll(" +", " "); // Normalize
                    CellArea mergedArea = getMergedAreaFromCell(cells, cell);

                    if (cellValue.equals("№ п.п") || cellValue.equals("№ п/п")) {
                        // Found start pos
                        if (mergedArea != null) {
                            headerRowY = row + 1; // Adjust to the row below the header
                            headerColumnX = col;
                        } else {
                            headerRowY = row;
                            headerColumnX = col;
                        }
                        System.out.println("Found start position at coordinates: (" + headerRowY + ", " + headerColumnX + ")");
                        headerFound = true;

                        if (cells.get(headerRowY + 1, headerColumnX).getStringValue().equals("1") &&
                                cells.get(headerRowY + 1, headerColumnX + 1).getStringValue().equals("2"))  {
                            DataWorkColumnX = headerColumnX;
                            DataWorkRowY = headerRowY + 2;
                        }
                        else {
                            DataWorkColumnX = headerColumnX;
                            DataWorkRowY = headerRowY + 1;
                        }
                        break outerLoop; // Stop searching
                    }
                }
            }

            // Parse the header row if "номер документа" was found
            if (headerFound && headerRowY != -1) {
                parsedDocumentMap.clear(); // Clear the previous map for each sheet
// ну и спагетти код...
                for (int col = 0; col <= cells.getMaxDataColumn(); col++) {
                    Cell cell = cells.get(headerRowY, col);
                    CellArea mergedArea = getMergedAreaFromCell(cells, cell);
                    int colToWrite = col;
                    if (mergedArea != null) {
                        cell = cells.get(mergedArea.StartRow, mergedArea.StartColumn);
                        col = mergedArea.EndColumn;
                        colToWrite  = mergedArea.StartColumn;
                    }

                    String cellValue = cell.getStringValue().trim().replaceAll(" +", " ");
                    if (!cellValue.isEmpty()) {
                        parsedDocumentMap.put(cellValue.toLowerCase(), colToWrite); // Map header name to column index
                    }
                }

                // Output parsed headers for this sheet
                System.out.println("Parsed headers in row " + headerRowY + ":");
                for (Map.Entry<String, Integer> entry : parsedDocumentMap.entrySet()) {
                    System.out.println("Header: " + entry.getKey() + ", Column: " + entry.getValue());
                }
            } else {
                System.out.println("start pos not found in sheet: " + sheet.getName());
            }
        } catch (Exception e) {
            throw new RuntimeException("Error parsing document header", e);
        }
    }

    private CellArea getMergedAreaFromCell(Cells cells, Cell cell) {
        ArrayList<CellArea> mergedAreas = cells.getMergedCells(); // Get all merged cell areas
        int row = cell.getRow();
        int col = cell.getColumn();

        for (CellArea area : mergedAreas) {
            // Check if the cell is within this merged area
            if (row >= area.StartRow && row <= area.EndRow &&
                    col >= area.StartColumn && col <= area.EndColumn) {
                return area; // Return the matched merged area
            }
        }
        return null; // No matching merged area found
    }

    public void save() {
        try {
            Worksheet worksheet = header.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

//
        } catch (Exception e) {
            throw new RuntimeException("Error while saving the workbook", e);
        }
    }
}