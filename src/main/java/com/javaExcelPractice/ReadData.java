package com.javaExcelPractice;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	private void readExcel(XSSFSheet sheetCreate2) {
		// reading with null excel
		XSSFRow row;
		int totalrowof = sheetCreate2.getLastRowNum();
		int totalcellof = sheetCreate2.getRow(0).getLastCellNum();

		System.out.println("totalNumberOfRows1:" + totalrowof);
		System.out.println("totalNumberOfColumn1:" + totalcellof);

		for (int r = 0; r <= totalrowof; r++) {
			row = sheetCreate2.getRow(r);

			for (int c = 0; c < totalcellof; c++) {
				if (row.getCell(c).getCellType() == CellType.STRING) {
					System.out.print(row.getCell(c).getStringCellValue());
				} else if (row.getCell(c).getCellType() == CellType.NUMERIC) {
					System.out.print(row.getCell(c).getNumericCellValue());
				} else if (row.getCell(c).getCellType() == CellType.BOOLEAN) {
					System.out.print(row.getCell(c).getBooleanCellValue());
				}
			}
			System.out.println();
		}

	}

	public static void main(String[] args) throws IOException {

		String firstExcelPath = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file1.xlsx";
		FileInputStream file1 = new FileInputStream(firstExcelPath);
		XSSFWorkbook workBook1 = new XSSFWorkbook(file1);
		XSSFSheet sheet1 = workBook1.getSheetAt(0);
		XSSFRow row;

		int totalNumberOfRows1 = sheet1.getLastRowNum();
		int totalNumberOfColumn1 = sheet1.getRow(0).getLastCellNum();

		System.out.println("totalNumberOfRows1:" + totalNumberOfRows1);
		System.out.println("totalNumberOfColumn1:" + totalNumberOfColumn1);

		for (int r = 0; r <= totalNumberOfRows1; r++) {
			row = sheet1.getRow(r);

			for (int c = 0; c < totalNumberOfColumn1; c++) {
				if (row.getCell(c).getCellType() == CellType.STRING) {
					System.out.print(row.getCell(c).getStringCellValue());
				} else if (row.getCell(c).getCellType() == CellType.NUMERIC) {
					System.out.print(row.getCell(c).getNumericCellValue());
				} else if (row.getCell(c).getCellType() == CellType.BOOLEAN) {
					System.out.print(row.getCell(c).getBooleanCellValue());
				}
			}
			System.out.println();
		}

		workBook1.close();
	}
}

