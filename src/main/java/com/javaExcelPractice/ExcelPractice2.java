package com.javaExcelPractice;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Done creating new Two Different Excel but with null row included
public class ExcelPractice2 {

	public static void readExcel() throws IOException {

		String firstExcelPath = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file1.xlsx";
		FileInputStream file1 = new FileInputStream(firstExcelPath);
		XSSFWorkbook workBook1 = new XSSFWorkbook(file1);
		XSSFSheet sheet1 = workBook1.getSheetAt(0);

		String secondExcelPath = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file2.xlsx";
		FileInputStream file2 = new FileInputStream(secondExcelPath);
		XSSFWorkbook workBook2 = new XSSFWorkbook(file2);
		XSSFSheet sheet2 = workBook2.getSheetAt(0);

		// workBook1
		int totalNumberOfRows1 = sheet1.getLastRowNum();
		int totalNumberOfColumn1 = sheet1.getRow(0).getLastCellNum();
		int key1 = 0;
		XSSFCell Excel1Row1cellKeyValue = sheet1.getRow(0).getCell(key1);
		XSSFCell cellOfRowKey1;
		XSSFRow rowOfSameKey1;

		// workBook2
		int totalNumberOfRows2 = sheet2.getLastRowNum();
		int totalNumberOfColumn2 = sheet2.getRow(0).getLastCellNum();
		int key2 = 1;
		XSSFCell Excel2Row1cellKeyValue = sheet2.getRow(0).getCell(key2);
		XSSFCell cellOfRowKey2;
		XSSFRow rowOfSameKey2;

		// creating new working and adding new rows
		XSSFWorkbook workBookOutput1 = new XSSFWorkbook(firstExcelPath);
		XSSFSheet sheetCreate1 = workBookOutput1.getSheetAt(0);

		// going to Excel1 key -> row = 1 to last
		for (int r = 1; r <= totalNumberOfRows1; r++) {
			cellOfRowKey1 = sheet1.getRow(r).getCell(Excel1Row1cellKeyValue.getColumnIndex());
			// going to Excel2 key -> row = 1 to last
			for (int e = 1; e <= totalNumberOfRows2; e++) {
				cellOfRowKey2 = sheet2.getRow(e).getCell(Excel2Row1cellKeyValue.getColumnIndex());
				if (cellOfRowKey1.getNumericCellValue() == cellOfRowKey2.getNumericCellValue()) {
//					System.out.println("SameCells1:" + cellOfRowKey1);
					rowOfSameKey1 = sheet1.getRow(r);
					sheet1.removeRow(rowOfSameKey1);
//					sheet1.removeRowBreak(r);
					break;
				}
			}
		}

		String firstExcelPathCopy = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file1.xlsx";
		FileInputStream file1Copy = new FileInputStream(firstExcelPathCopy);
		XSSFWorkbook workBook1Copy = new XSSFWorkbook(file1Copy);
		XSSFSheet sheet1Copy = workBook1Copy.getSheetAt(0);
		XSSFCell cellOfRowKey1Copy;
		
		// going to Excel2 key -> row = 1 to last
		for (int r = 1; r <= totalNumberOfRows2; r++) {
			cellOfRowKey2 = sheet2.getRow(r).getCell(Excel2Row1cellKeyValue.getColumnIndex());
			// going to Excel1 key -> row = 1 to last
			for (int e = 1; e <= totalNumberOfRows1; e++) {
				cellOfRowKey1Copy = sheet1Copy.getRow(e).getCell(Excel1Row1cellKeyValue.getColumnIndex());
				if (cellOfRowKey2.getNumericCellValue() == cellOfRowKey1Copy.getNumericCellValue()) {
//					System.out.println("SameCells2:" + cellOfRowKey2);
					rowOfSameKey2 = sheet2.getRow(r);
					sheet2.removeRow(rowOfSameKey2);
					break;
				}
			}
		}

		String target1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\New Folder\\outputFile_1.xlsx";
		FileOutputStream outputStream1 = new FileOutputStream(target1Path);
		
		String target2Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\New Folder\\outputFile_2.xlsx";
		FileOutputStream outputStream2 = new FileOutputStream(target2Path);
		
		workBook1.write(outputStream1);
		workBook2.write(outputStream2);
		
		workBook1.close();
		workBook2.close();
		
		System.out.println("done....");
	}

	public static void main(String[] args) {
		try {
			readExcel();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
