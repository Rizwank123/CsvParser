package com.csv.org;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CsvReader {
	private final String outputFile = "output.xlsx";

	public void dataMapFile(String in) throws IOException {
		{
			Map<String, String> dataMap = new HashMap<>();
			String inputCSVFile = "input.csv";

			// Read the CSV data
			try (BufferedReader reader = new BufferedReader(new FileReader(inputCSVFile))) {
				String line;
				while ((line = reader.readLine()) != null) {
					String[] entries = line.split(", ");
					for (String entry : entries) {
						String[] parts = entry.split(": ");
						dataMap.put(parts[0], parts[1]);
					}
				}
			} catch (IOException e) {
				e.printStackTrace();
			}

			// Print the data map
			System.out.println("Data Map: " + dataMap);

			Map<String, Integer> finalMap = new HashMap<>();
			boolean updated;

			do {
				updated = false;
				for (String key : dataMap.keySet()) {
					if (!finalMap.containsKey(key)) {
						try {
							String s = dataMap.get(key);
							if (s.charAt(0) == '=') {
								s = s.substring(1); // Remove the "=" character
							}
							String[] parts = s.split("\\+");
							int sum = 0;
							for (String part : parts) {
								part = part.trim();
								if (Character.isDigit(part.charAt(0))) {
									sum += Integer.parseInt(part);
								} else {
									Integer value = finalMap.get(part);
									if (value != null) {
										sum += value;
									} else {
										throw new IllegalArgumentException("Reference not yet calculated: " + part);
									}
								}
							}
							finalMap.put(key, sum);
							updated = true;
						} catch (Exception e) {
							// Skip updating this key until dependencies are resolved
						}
					}
				}
			} while (updated);
			writeToExcel(finalMap, outputFile);
			System.out.println("Final Map: " + finalMap);
		}

	}

	private void writeToExcel(Map<String, Integer> dataMap, String outputFile) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Evaluated Data");

		// Create header row
		Row headerRow = sheet.createRow(0);
		String[] headers = { "sn", "A", "B", "C" };
		for (int i = 0; i < headers.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(headers[i]);
		}

		// Add data rows starting from row 1 (after the header)
		for (int i = 1; i <= 3; i++) {
			Row row = sheet.createRow(i); // Start from row 1
			Cell snCell = row.createCell(0);
			snCell.setCellValue(i); // Serial number

			for (int j = 0; j < 3; j++) {
				Cell cell = row.createCell(j + 1);
				String cellKey = "" + (char) ('A' + j) + i;
				Integer cellValue = dataMap.get(cellKey);
				if (cellValue != null) {
					cell.setCellValue(cellValue);
				}
			}
		}

		// Write the output to an XLSX file
		try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
			workbook.write(fileOut);
		}

		// Closing the workbook
		// workbook.close();
	}
}
