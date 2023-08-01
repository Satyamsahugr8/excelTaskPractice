package com.javaExcelPractice;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPractice4 {

	public static void removeRow(XSSFSheet sheet, int rowIndex) {
		int lastRowNum = sheet.getLastRowNum();
		if (rowIndex >= 0 && rowIndex < lastRowNum) {
			sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
		}
		if (rowIndex == lastRowNum) {
			XSSFRow removingRow = sheet.getRow(rowIndex);
			if (removingRow != null) {
				sheet.removeRow(removingRow);
			}
		}
	}

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
//					removeRow(sheet1, r);
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

		// Upto here we have to two excel with some null or empty row

		// count null row
		int counter = 0;
		ArrayList<Integer> nullIndex = new ArrayList<Integer>();

		for (int r = 1; r <= totalNumberOfRows1; r++) {
			if (sheet1.getRow(r) == null) {
				counter++;
				nullIndex.add(r);
			}
		}

		int totalNumberOfRowsOfNewSheet = (1+totalNumberOfRows1) - counter;
		System.out.println("totalNumberOfRows1:" + totalNumberOfRows1);
		System.out.println("totalNumberOfRowsOfNewSheet:" + totalNumberOfRowsOfNewSheet);

		
//      creating new working and adding new rows for excel1
		XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
		XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
		XSSFRow rowCreated = null;
		XSSFCell cellCreated = null;

		int u = 0;
		
		for (int r = 0; r <= totalNumberOfRowsOfNewSheet; r++) {
			rowCreated = sheetCreate1.createRow(r);
			
		for (int c = 0; c < totalNumberOfColumn1; c++) {
			cellCreated = rowCreated.createCell(c);
		}
	}
			
		//
		for (int p = 0; p <= totalNumberOfRows1; p++) {
			
			if (sheet1.getRow(p) == null) {
				continue;
			} else {
				rowCreated = sheetCreate1.getRow(u);
				u++;
				for (int d = 0; d < totalNumberOfColumn1; d++) {
//					rowCreated.getCell(d).setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
					if (sheet1.getRow(p).getCell(d).getCellType() == CellType.STRING) {
						rowCreated.getCell(d).setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
					} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
						rowCreated.getCell(d).setCellValue(sheet1.getRow(p).getCell(d).getNumericCellValue());
					} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
						rowCreated.getCell(d).setCellValue(sheet1.getRow(p).getCell(d).getBooleanCellValue());
					}
				}
			}
		}
		
		
		
		
		
		
		
//		//removeRow(sheet1, r);
//		for (int r = 1; r <= totalNumberOfRows1; r++) {
//			if(sheet1.getRow(r) == null) {
//				int lastRowNum=sheet1.getLastRowNum();
//			    if(r>=0 && r<lastRowNum){
//			        sheet1.shiftRows(r+1,lastRowNum, -1);
//			    }
//			    if(r==lastRowNum){
//			        XSSFRow removingRow=sheet1.getRow(r);
//			        if(removingRow!=null){
//			            sheet1.removeRow(removingRow);
//			        }
//			    }
//			}
//		} // removing only one null row not recommended

		
		
		
		
		
		
		
		
		
		
		
//		// copying that row to new excel
////		FileInputStream sourceFile = new FileInputStream(path1);
////		XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
//		XSSFSheet sourceSheet = workBook1.getSheetAt(0);
//
//		String target1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\New Folder2\\outputFile_1.xlsx";
//		
//		FileInputStream targetFile = new FileInputStream(target1Path);
//		XSSFWorkbook targetWorkbook = new XSSFWorkbook(targetFile);
//		XSSFSheet targetSheet = targetWorkbook.getSheetAt(0);
//
//		int rowCount = sourceSheet.getLastRowNum();
//		for (int i = 0; i <= rowCount; i++) {
//			Row sourceRow = sourceSheet.getRow(i);
//			Row destRow = targetSheet.createRow(i);
//			if (sourceRow == null) {
//				continue;
//			}
//
//			int cellCount = sourceRow.getLastCellNum();
//
//			for (int j = 0; j < cellCount; j++) {
//				Cell sourceCell = sourceRow.getCell(j);
//				Cell destCell = destRow.createCell(j);
//				if (sourceCell == null) {
//					continue;
//				}
//				if (sourceCell.getCellType() == CellType.STRING) {
//					destCell.setCellValue(sourceCell.getStringCellValue());
//				} else if (sourceCell.getCellType() == CellType.NUMERIC) {
//					destCell.setCellValue(sourceCell.getNumericCellValue());
//				} else if (sourceCell.getCellType() == CellType.BOOLEAN) {
//					destCell.setCellValue(sourceCell.getBooleanCellValue());
//				}
//			}
//		}
//		
//		FileOutputStream outputStream = new FileOutputStream(target1Path);
//		targetWorkbook.write(outputStream);
//		targetWorkbook.close();
//		outputStream.close();
//		System.out.println("Data copied successfully..");
//		//copying complete

		
		
		
		
		
		
		
		
//		String target1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\New Folder\\outputFile_1.xlsx";
//		FileOutputStream outputStream1 = new FileOutputStream(target1Path);
//		workBook1.write(outputStream1);

//		String target2Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\New Folder\\outputFile_2.xlsx";
//		FileOutputStream outputStream2 = new FileOutputStream(target2Path);
//		workBook2.write(outputStream2);

		
		//removed null excel writing
		String target1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\New Folder\\outputFile_1-Copy.xlsx";
		FileOutputStream outputStream1 = new FileOutputStream(target1Path);
		workBookOutput1.write(outputStream1);
		
		
		
		workBookOutput1.close();
		workBook1.close();
		workBook2.close();
		
		

		System.out.println("done....");
	}

	private void fetchExcel(String path1, String path2, int keyFile1, int keyFile2, String folderPath) {
		try {

			FileInputStream sourceFile = new FileInputStream(path1);
			XSSFWorkbook sourceWorkbook = new XSSFWorkbook(sourceFile);
			XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(0);

			FileInputStream targetFile = new FileInputStream(path2);
			XSSFWorkbook targetWorkbook = new XSSFWorkbook(targetFile);
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

			FileOutputStream outputStream = new FileOutputStream(path2);
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

	public static void main(String[] args) {
		try {
			readExcel();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
