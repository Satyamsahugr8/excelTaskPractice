package com.javaExcelPractice;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelComparatorByPooja {

	public static void main(String[] args) throws IOException {
		String file1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file1.xlsx";
		String file2Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file2.xlsx";
		FileInputStream file1 = new FileInputStream(file1Path);
		FileInputStream file2 = new FileInputStream(file2Path);
		boolean filesEqual = true;

		// Create Workbook instances for both files
		XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
		XSSFWorkbook workbook2 = new XSSFWorkbook(file2);

		// Iterate through each sheet in both workbooks
		for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
			Sheet sheet1 = workbook1.getSheetAt(i);
			Sheet sheet2 = workbook2.getSheetAt(i);

			// Compare sheet names
			if (!sheet1.getSheetName().equals(sheet2.getSheetName())) {
				System.out.println("Sheet names do not match for sheet " + i);
				filesEqual = false;
				continue;
			}

			// Compare number of rows
			if (sheet1.getLastRowNum() != sheet2.getLastRowNum()) {
				System.out.println("Number of rows do not match for sheet " + i);
				filesEqual = false;
				continue;
			}

			// Compare each cell in each row
			for (int j = 0; j <= sheet1.getLastRowNum(); j++) {
				Row row1 = sheet1.getRow(j);
				Row row2 = sheet2.getRow(j);
				if (row1 == null && row2 == null) {
					continue;
				} else if (row1 == null || row2 == null) {
					System.out.println("Number of columns do not match for row " + j + " in sheet " + i);
					filesEqual = false;
					continue;
				}
				for (int k = 0; k < row1.getLastCellNum(); k++) {
					Cell cell1 = row1.getCell(k);
					Cell cell2 = row2.getCell(k);
					if (!cell1.toString().equals(cell2.toString())) {
						System.out
								.println("Cell values do not match for cell " + k + " in row " + j + " of sheet " + i);
						filesEqual = false;
					}
				}
			}
		}

		if (filesEqual) {
			System.out.println("Files are equal");
		}
		workbook1.close();
		workbook2.close();
		file1.close();
		file2.close();
	}
}