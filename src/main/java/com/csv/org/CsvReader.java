package com.csv.org;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.util.*;
import java.util.regex.*;

public class CsvReader {
    private static final Pattern FORMULA_PATTERN = Pattern.compile("=([A-Z]\\d+)([\\+\\-\\*/])(\\d+|[A-Z]\\d+)");

    public static void main(String[] args) throws Exception {
        String inputFile = "input.csv"; // Input CSV file
        String outputFile = "output.xls"; // Output XLS file to store the results
        
        // Step 1: Read CSV into a 2D array
        List<List<String>> csvData = readCSV(inputFile);

        // Step 2: Evaluate all formulas in the CSV
        evaluateFormulas(csvData);

        // Step 3: Write the result to an Excel file
        writeXLS(outputFile, csvData);

        System.out.println("Formula evaluation complete. Results written to " + outputFile);
    }

    private static List<List<String>> readCSV(String inputFile) throws IOException {
        List<List<String>> csvData = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(new FileReader(inputFile))) {
            String line;
            while ((line = br.readLine()) != null) {
                String[] values = line.split(",");
                csvData.add(new ArrayList<>(Arrays.asList(values)));
            }
        }
        return csvData;
    }

    private static void evaluateFormulas(List<List<String>> csvData) {
        for (int row = 0; row < csvData.size(); row++) {
            for (int col = 0; col < csvData.get(row).size(); col++) {
                String cellValue = csvData.get(row).get(col);
                if (cellValue.startsWith("=")) {
                    try {
                        String evaluatedValue = evaluateFormula(cellValue, csvData);
                        csvData.get(row).set(col, evaluatedValue);
                    } catch (Exception e) {
                        System.err.println("Error evaluating formula in cell (" + (row + 1) + "," + (col + 1) + "): " + e.getMessage());
                        csvData.get(row).set(col, "ERROR");
                    }
                }
            }
        }
    }

    private static String evaluateFormula(String formula, List<List<String>> csvData) throws Exception {
        Matcher matcher = FORMULA_PATTERN.matcher(formula);
        if (!matcher.matches()) {
            throw new Exception("Invalid formula syntax: " + formula);
        }

        // Extract cell references and operator from the formula
        String operand1 = matcher.group(1);
        String operator = matcher.group(2);
        String operand2 = matcher.group(3);

        // Resolve the first operand
        int value1 = resolveCellValue(operand1, csvData);

        // Resolve the second operand (can be a number or a cell reference)
        int value2;
        if (isCellReference(operand2)) {
            value2 = resolveCellValue(operand2, csvData);
        } else {
            value2 = Integer.parseInt(operand2);
        }

        // Perform the operation
        int result;
        switch (operator) {
            case "+":
                result = value1 + value2;
                break;
            case "-":
                result = value1 - value2;
                break;
            case "*":
                result = value1 * value2;
                break;
            case "/":
                if (value2 == 0) throw new Exception("Division by zero");
                result = value1 / value2;
                break;
            default:
                throw new Exception("Unknown operator: " + operator);
        }

        return String.valueOf(result);
    }

    private static int resolveCellValue(String cellReference, List<List<String>> csvData) throws Exception {
        int row = cellReference.charAt(1) - '1';  // Convert '1', '2', '3' to 0-based index
        int col = cellReference.charAt(0) - 'A';  // Convert 'A', 'B', 'C' to 0-based index

        if (row >= csvData.size() || col >= csvData.get(row).size()) {
            throw new Exception("Cell reference out of bounds: " + cellReference);
        }

        String cellValue = csvData.get(row).get(col);
        if (cellValue.startsWith("=")) {
            return Integer.parseInt(evaluateFormula(cellValue, csvData));  // Recursively evaluate the formula
        } else {
            return Integer.parseInt(cellValue);  // Return the static value
        }
    }

    private static boolean isCellReference(String input) {
        return input.matches("[A-Z]\\d+");
    }

    private static void writeXLS(String outputFile, List<List<String>> csvData) throws IOException {
        try (Workbook workbook = new HSSFWorkbook(); FileOutputStream fileOut = new FileOutputStream(outputFile)) {
            Sheet sheet = workbook.createSheet("Sheet1");

            for (int rowIndex = 0; rowIndex < csvData.size(); rowIndex++) {
                Row row = sheet.createRow(rowIndex);
                List<String> rowData = csvData.get(rowIndex);

                for (int colIndex = 0; colIndex < rowData.size(); colIndex++) {
                    Cell cell = row.createCell(colIndex);
                    String cellValue = rowData.get(colIndex);
                    try {
                        int numericValue = Integer.parseInt(cellValue);
                        cell.setCellValue(numericValue); // If it's a number, set as numeric
                    } catch (NumberFormatException e) {
                        cell.setCellValue(cellValue); // If it's text, set as string
                    }
                }
            }

            workbook.write(fileOut); // Write the workbook to the output file
        }
    }
}

