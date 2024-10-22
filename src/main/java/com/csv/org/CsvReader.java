package com.csv.org;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class CsvReader {

		    public static void main(String[] args) throws Exception {
		        // Input and Output File Paths
		        String inputCSVFile = "input.csv";
		        String outputXLSXFile = "output.xlsx";

		        // Read the CSV into a 2D array (for simplicity, let's assume fixed size here)
		        String[][] data = readCSV(inputCSVFile);

		        // Evaluate formulas
		        String[][] evaluatedData = evaluateFormulas(data);

		        // Write the evaluated data to XLSX with headers
		        writeToExcel(evaluatedData, outputXLSXFile);
		    }

		    private static String[][] readCSV(String csvFile) throws IOException {
		        // Simulated CSV data for simplicity
		        String[][] data = {
		                {"5", "3", "=5+A1"},
		                {"7", "8", "=A2+B2"},
		                {"9", "=4+5", "=C2+B3"}
		        };
		        return data;
		    }

		    private static String[][] evaluateFormulas(String[][] data) {
		        int rows = data.length;
		        int cols = data[0].length;
		        Map<String, Integer> cellReference = new HashMap<>();
		        String[][] result = new String[rows][cols];

		        for (int i = 0; i < rows; i++) {
		            for (int j = 0; j < cols; j++) {
		                String value = data[i][j];
		                result[i][j] = value.startsWith("=") ? evaluateFormula(value, result) : value;
		                cellReference.put("" + (char) ('A' + j) + (i + 1), Integer.parseInt(result[i][j]));
		            }
		        }
		        return result;
		    }

		    private static String evaluateFormula(String formula, String[][] data) {
		        formula = formula.substring(1);  // Remove leading "="
		        String[] parts = formula.split("\\+");

		        int result = 0;
		        for (String part : parts) {
		            part = part.trim();
		            if (Character.isDigit(part.charAt(0))) {
		                result += Integer.parseInt(part);
		            } else {  // This is a cell reference (e.g., A1)
		                int row = part.charAt(1) - '1';
		                int col = part.charAt(0) - 'A';
		                result += Integer.parseInt(data[row][col]);
		            }
		        }
		        return String.valueOf(result);
		    }

		    private static void writeToExcel(String[][] data, String outputFile) throws IOException {
		        Workbook workbook = new XSSFWorkbook();
		        Sheet sheet = workbook.createSheet("Evaluated Data");

		        // Create header row
		        Row headerRow = sheet.createRow(0);
		        String[] headers = {"", "A", "B", "C"};
		        for (int i = 0; i < headers.length; i++) {
		            Cell cell = headerRow.createCell(i);
		            cell.setCellValue(headers[i]);
		        }

		        // Add data rows starting from row 1 (after the header)
		        for (int i = 0; i < data.length; i++) {
		            Row row = sheet.createRow(i + 1);  // Start from row 1
		            for (int j = 0; j < data[i].length; j++) {
		                Cell cell = row.createCell(j);
		                cell.setCellValue(Integer.parseInt(data[i][j]));
		            }
		        }

		        // Write the output to an XLSX file
		        try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
		            workbook.write(fileOut);
		        }

		        // Closing the workbook
		        workbook.close();
		    }
   
}
      

    
    
    
    

    
    
    
        
            
            
            
        
    