package com.javaExcelPractice;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
//import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//done creating with two excel with unique of file1 to output1 and file2 to output2
//with null space with no exception raised
public class ExcelPractice3 {

	public static void readExcel() throws IOException {

		String firstExcelPath = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file1.xlsx";
//		String firstExcelPath = "C:\\Users\\SATYASAH\\Downloads\\Capg Bench.xlsx";
		FileInputStream file1 = new FileInputStream(firstExcelPath);
		XSSFWorkbook workBook1 = new XSSFWorkbook(file1);
		XSSFSheet sheet1 = workBook1.getSheetAt(0);

		String secondExcelPath = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file2.xlsx";
//		String secondExcelPath = "C:\\Users\\SATYASAH\\Downloads\\CapgBenchhh.xlsx";
		FileInputStream file2 = new FileInputStream(secondExcelPath);
		XSSFWorkbook workBook2 = new XSSFWorkbook(file2);
		XSSFSheet sheet2 = workBook2.getSheetAt(0);

		// workBook1
		int totalNumberOfRowsInExcel1 = sheet1.getLastRowNum();
		int totalNumberOfColumnInExcel1 = sheet1.getRow(0).getLastCellNum();
		int key1 = 0;
//		XSSFCell Excel1Row1cellKeyValue = sheet1.getRow(0).getCell(key1);
		XSSFCell cellOfRowKey1;
		XSSFRow rowOfSameKey1;

		// workBook2
		int totalNumberOfRowsInExcel2 = sheet2.getLastRowNum();
		int totalNumberOfColumnInExcel2 = sheet2.getRow(0).getLastCellNum();
		int key2 = 0;
//		XSSFCell Excel2Row1cellKeyValue = sheet2.getRow(0).getCell(key2);
		XSSFCell cellOfRowKey2;
		XSSFRow rowOfSameKey2;

		// going to Excel1 key -> row = 1 to last
		for (int r = 1; r <= totalNumberOfRowsInExcel1; r++) {
			if (sheet1.getRow(r) == null) {
				continue;
			} else {
				if (sheet1.getRow(r).getCell(key1) == null) {
					continue;
				} else {
					cellOfRowKey1 = sheet1.getRow(r).getCell(key1);
				}

				// going to Excel2 key -> row = 1 to last
				for (int e = 1; e <= totalNumberOfRowsInExcel2; e++) {
					if (sheet2.getRow(e) == null) {
						continue;
					} else {
						if (sheet2.getRow(e).getCell(key2) == null) {
							continue;
						} else {
							cellOfRowKey2 = sheet2.getRow(e).getCell(key2);
						}
//				cellOfRowKey2 = sheet2.getRow(e).getCell(key2);
						if (cellOfRowKey1.getNumericCellValue() == cellOfRowKey2.getNumericCellValue()) {
//					System.out.println("SameCells1:" + cellOfRowKey1);
							rowOfSameKey1 = sheet1.getRow(r);
							sheet1.removeRow(rowOfSameKey1);
//					sheet1.removeRowBreak(r);
//					removeRow(sheet1, r);
							break;
						}

					} // else
				} // for
			} // else
		} // for

		String firstExcelPathCopy = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file1.xlsx";
//		String firstExcelPathCopy = "C:\\Users\\SATYASAH\\Downloads\\Capg Bench.xlsx";
		FileInputStream file1Copy = new FileInputStream(firstExcelPathCopy);
		XSSFWorkbook workBook1Copy = new XSSFWorkbook(file1Copy);
		XSSFSheet sheet1Copy = workBook1Copy.getSheetAt(0);
		XSSFCell cellOfRowKey1Copy;

		// going to Excel2 key -> row = 1 to last
		for (int rr = 1; rr <= totalNumberOfRowsInExcel2; rr++) {
			if (sheet2.getRow(rr) == null) {
				continue;
			} else {
				if (sheet2.getRow(rr).getCell(key2) == null) {
					continue;
				} else {
					cellOfRowKey2 = sheet2.getRow(rr).getCell(key2);
				}
				// going to Excel1 key -> row = 1 to last
				for (int e = 1; e <= totalNumberOfRowsInExcel1; e++) {
					if (sheet1Copy.getRow(e) == null) {
						continue;
					} else {
						if (sheet1Copy.getRow(e).getCell(key1) == null) {
							continue;
						} else {
							cellOfRowKey1Copy = sheet1Copy.getRow(e).getCell(key1);
						}
//				cellOfRowKey1Copy = sheet1Copy.getRow(e).getCell(key1);
						if (cellOfRowKey2.getNumericCellValue() == cellOfRowKey1Copy.getNumericCellValue()) {
//					System.out.println("SameCells2:" + cellOfRowKey2);
							rowOfSameKey2 = sheet2.getRow(rr);
							sheet2.removeRow(rowOfSameKey2);
							break;
						}
					} // else
				} // for
			} // else
		} // for

//		String target1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\Output3\\outputFileWithSpace_1.xlsx";
////		String target1Path = "C:\\Users\\SATYASAH\\Downloads\\output\\outputFileWithSpace_1.xlsx";
//		FileOutputStream outputStream1 = new FileOutputStream(target1Path);
//		workBook1.write(outputStream1);
//
//		String target2Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\Output3\\outputFileWithSpace_2.xlsx";
////		String target2Path = "C:\\Users\\SATYASAH\\Downloads\\output\\outputFileWithSpace_2.xlsx";
//		FileOutputStream outputStream2 = new FileOutputStream(target2Path);
//		workBook2.write(outputStream2);

		// Upto here we have to two excel with some null or empty row
		// sheet1 and sheet2 as output only NO new sheet created

//------------------------------------------------------------------------------------------------------------		

		// counting null row in EXCEL 1
		int counter = 0;
		for (int r = 0; r <= totalNumberOfRowsInExcel1; r++) {
			if (sheet1.getRow(r) == null) {
				counter++;
			}
		}

		System.out.println("totalNumberOfRows1:" + totalNumberOfRowsInExcel1);
		System.out.println("counter:" + counter);

		int totalNumberOfRowsOfNewSheet = (1 + totalNumberOfRowsInExcel1) - counter;

		System.out.println("totalNumberOfRowsOfNewSheet1:" + totalNumberOfRowsOfNewSheet);

		// creating new working and adding new rows for excel1
		XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
		XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
		XSSFRow rowCreated = null;
		XSSFCell cellCreated = null;

		for (int r = 0; r < totalNumberOfRowsOfNewSheet; r++) {
			rowCreated = sheetCreate1.createRow(r);

			for (int c = 0; c < totalNumberOfColumnInExcel1; c++) {
				cellCreated = rowCreated.createCell(c);
			}
		}

		for (int p = 0, u = 0; p <= totalNumberOfRowsInExcel1; p++) {
			if (sheet1.getRow(p) == null) {
				continue;
			} else {
				rowCreated = sheetCreate1.getRow(u);
				u++;
				for (int d = 0; d < totalNumberOfColumnInExcel1; d++) {
					if (sheet1.getRow(p) == null) {
						continue;
					} else {
						if (sheet1.getRow(p).getCell(d) == null) {
							continue;
						} else {
							if (sheet1.getRow(p).getCell(d).getCellType() == null) {
								continue;
							}
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
			}
		}

		// counting null row in EXCEL 2
		int counter2 = 0;
		for (int r = 1; r <= totalNumberOfRowsInExcel2; r++) {
			if (sheet2.getRow(r) == null) {
				counter2++;
			}
		}

		System.out.println("totalNumberOfRows2:" + totalNumberOfRowsInExcel2);
		System.out.println("counter2:" + counter2);

		int totalNumberOfRowsOfNewSheet2 = (1 + totalNumberOfRowsInExcel2) - counter2;
		System.out.println("totalNumberOfRowsOfNewSheet2:" + totalNumberOfRowsOfNewSheet2);

		// creating new working and adding new rows for excel1
		XSSFWorkbook workBookOutput2 = new XSSFWorkbook();
		XSSFSheet sheetCreate2 = workBookOutput2.createSheet();
		XSSFRow rowCreated2 = null;
		XSSFCell cellCreated2 = null;

		int v = 0;
		for (int r = 0; r < totalNumberOfRowsOfNewSheet2; r++) {
			rowCreated2 = sheetCreate2.createRow(r);

			for (int c = 0; c < totalNumberOfColumnInExcel2; c++) {
				cellCreated2 = rowCreated2.createCell(c);
			}
		}

		for (int p = 0; p <= totalNumberOfRowsInExcel2; p++) {
			if (sheet2.getRow(p) == null) {
				continue;
			} else {
				rowCreated2 = sheetCreate2.getRow(v);
				v++;
				for (int d = 0; d < totalNumberOfColumnInExcel2; d++) {
					if (sheet2.getRow(p) == null) {
						continue;
					} else {
						if(sheet2.getRow(p).getCell(d) == null) {
							continue;
						} else {
					if (sheet2.getRow(p).getCell(d).getCellType() == null) {
						continue;
					}
					if (sheet2.getRow(p).getCell(d).getCellType() == CellType.STRING) {
						rowCreated2.getCell(d).setCellValue(sheet2.getRow(p).getCell(d).getStringCellValue());
					} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
						rowCreated2.getCell(d).setCellValue(sheet2.getRow(p).getCell(d).getNumericCellValue());
					} else if (sheet2.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
						rowCreated2.getCell(d).setCellValue(sheet2.getRow(p).getCell(d).getBooleanCellValue());
					}
				}}
				}
			}
		}

		// removed null excel writing
		String target1Path1 = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\New Folder5\\outputFileFinal_1.xlsx";
		FileOutputStream outputStream11 = new FileOutputStream(target1Path1);
		workBookOutput1.write(outputStream11);
		workBookOutput1.close();

		String target1Path2 = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\New Folder5\\outputFileFinal_2.xlsx";
		FileOutputStream outputStream22 = new FileOutputStream(target1Path2);
		workBookOutput2.write(outputStream22);
		

		workBookOutput1.close();
		workBookOutput2.close();

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
