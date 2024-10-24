package com.csv.org;

import java.io.IOException;

public class MainHandler {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String inputFile = "input.csv";
		String outputFile = "output.xlsx";
		 CsvReader csvReader=new CsvReader();
		 csvReader.dataMapFile(inputFile);
		 //csvReader.csvProcessor(inputFile, outputFile);
		
		 
		 
		 

	}

}
