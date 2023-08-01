package com.javaExcelPractice;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UserByVivek {
	public static void main(String[] args) {
		try {

			FileInputStream sourceFile = new FileInputStream(
					"C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\List1.xlsx");
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);

			FileInputStream source2File = new FileInputStream(
					"C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\List2.xlsx");
			XSSFWorkbook targetWorkbook = new XSSFWorkbook(source2File);
			XSSFSheet targetSheet = targetWorkbook.getSheetAt(0);

			int rowCount = sourceSheet.getLastRowNum();
			for (int i = 0; i <= rowCount; i++) {
				Row sourceRow = sourceSheet.getRow(i);
				Row destRow = targetSheet.createRow(i);
				if (sourceRow == null) {
					continue;
				}

				int cellCount = sourceRow.getLastCellNum();

				for (int j = 0; j < cellCount; j++) {
					Cell sourceCell = sourceRow.getCell(j);
					Cell destCell = destRow.createCell(j);
					if (sourceCell == null) {
						continue;
					}

					if (sourceCell.getCellType() == CellType.STRING) {
						destCell.setCellValue(sourceCell.getStringCellValue());
					} else if (sourceCell.getCellType() == CellType.NUMERIC) {
						destCell.setCellValue(sourceCell.getNumericCellValue());
					} else if (sourceCell.getCellType() == CellType.BOOLEAN) {
						destCell.setCellValue(sourceCell.getBooleanCellValue());
					}

				}
			}
			
			FileOutputStream outputStream = new FileOutputStream(
					"C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\OutputFolder\\OutputList1.xlsx");
			targetWorkbook.write(outputStream);
			targetWorkbook.close();

			outputStream.close();

			sourceWorkbook.close();
			sourceFile.close();
			System.out.println("Data copied successfully..");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
