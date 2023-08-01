package com.javaExcelPractice;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Task3ByVivek {

//	private static void readExcel(XSSFSheet sheetCreate2) {
//		// reading with null excel
//		XSSFRow row;
//		int totalrowof = sheetCreate2.getLastRowNum();
//		int totalcellof = sheetCreate2.getRow(0).getLastCellNum();
//
//		System.out.println("totalNumberOfRows1:" + totalrowof);
//		System.out.println("totalNumberOfColumn1:" + totalcellof);
//
//		for (int r = 0; r <= totalrowof; r++) {
//			row = sheetCreate2.getRow(r);
//
//			for (int c = 0; c < totalcellof; c++) {
//				if (row.getCell(c).getCellType() == CellType.STRING) {
//					System.out.print(row.getCell(c).getStringCellValue());
//				} else if (row.getCell(c).getCellType() == CellType.NUMERIC) {
//					System.out.print(row.getCell(c).getNumericCellValue());
//				} else if (row.getCell(c).getCellType() == CellType.BOOLEAN) {
//					System.out.print(row.getCell(c).getBooleanCellValue());
//				}
//			}
//			System.out.println();
//		}
//	}

	@SuppressWarnings({ "resource" })
	public static void main(String[] args) {

		try {

			String firstExcelPath = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\BigexcelFiles\\New\\file1.xlsx";
			String secondExcelPath = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\BigexcelFiles\\New\\file2.xlsx";

			int sheetNo1 = 0;
			int sheetNo2 = 0;

			int key1 = 0;
			int key2 = 0;

//			List<String> sheet1List = new ArrayList();
//			List<String> sheet2List = new ArrayList();

			Set<String> sheet1List = new LinkedHashSet<String>();
			Set<String> sheet2List = new LinkedHashSet<String>();

			// =----------------------------------------------------------------=

			FileInputStream file1 = new FileInputStream(firstExcelPath);
			XSSFWorkbook workBook1 = new XSSFWorkbook(file1);
			XSSFSheet sheet1 = workBook1.getSheetAt(sheetNo1);

			FileInputStream file2 = new FileInputStream(secondExcelPath);
			XSSFWorkbook workBook2 = new XSSFWorkbook(file2);
			XSSFSheet sheet2 = workBook2.getSheetAt(sheetNo2);

			// workBook1
			int totalNumberOfRowsInExcel1 = sheet1.getLastRowNum();
			int totalNumberOfColumnInExcel1 = sheet1.getRow(0).getLastCellNum();

			// workBook2
			int totalNumberOfRowsInExcel2 = sheet2.getLastRowNum();
			int totalNumberOfColumnInExcel2 = sheet2.getRow(0).getLastCellNum();
			
			
			// creating new working and adding new rows for excel1
			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
			XSSFRow rowCreated = null;
			
			

			// =-------------------------------------------------------------------------=

//			System.out.println("totalNumberOfRowsInExcel1 :"+totalNumberOfRowsInExcel1);
//			System.out.println("totalNumberOfRowsInExcel2 :"+totalNumberOfRowsInExcel2);

			for (int r = 1; r <= totalNumberOfRowsInExcel1; r++) {
				sheet1List.add(sheet1.getRow(r).getCell(0).toString());
			}

			for (int r = 1; r <= totalNumberOfRowsInExcel2; r++) {
				sheet2List.add(sheet2.getRow(r).getCell(0).toString());
			}

			List<String> UniqueSet = sheet1List.stream().filter(a -> !sheet2List.contains(a))
					.collect(Collectors.toList());

			List<String> UniqueSet2 = sheet2List.stream().filter(a -> !sheet1List.contains(a))
					.collect(Collectors.toList());
			
			
			for (int r = 0; r <= UniqueSet.size() ; r++) {
				rowCreated = sheetCreate1.createRow(r);

				for (int c = 0; c < totalNumberOfColumnInExcel1; c++) {
					rowCreated.createCell(c);
				}
			}

			System.out.println("Unique From excel 1 ---------------------");
			for (String string : UniqueSet) {
				System.out.println(string);
			}

			System.out.println("Unique From excel 2 ---------------------");
			for (String string2 : UniqueSet2) {
				System.out.println(string2);
			}

			for (int i = 0; i < UniqueSet.size(); i++) {
				for (int j = 1; j <= totalNumberOfRowsInExcel1; j++) {
					if (sheet1.getRow(j) == null) {
						continue;
					} else {
						if ((UniqueSet.get(i).equals(sheet1.getRow(j).getCell(key1).toString()))) {
							
							System.out.println("j: "+j);
							System.out.println(UniqueSet.get(i) + "=" + (sheet1.getRow(j).getCell(key1).toString()));
							
							for (int d = 0; d < totalNumberOfColumnInExcel1; d++) {
								if (sheet1.getRow(j).getCell(d) == null) {
									continue;
								} else {
									if (sheet1.getRow(j).getCell(d).getCellType() == CellType.STRING) {
										rowCreated.getCell(d).setCellValue(sheet1.getRow(j).getCell(d).getStringCellValue());
									} else if (sheet1.getRow(j).getCell(d).getCellType() == CellType.NUMERIC) {
										rowCreated.getCell(d).setCellValue(sheet1.getRow(j).getCell(d).getNumericCellValue());
									} else if (sheet1.getRow(j).getCell(d).getCellType() == CellType.BOOLEAN) {
										rowCreated.getCell(d).setCellValue(sheet1.getRow(j).getCell(d).getBooleanCellValue());
									}
								}
							}
							break;
							
						} 
					}
				}
			}

			// New Excel Created

//			XSSFWorkbook workBookOutput1 = new XSSFWorkbook();
//			XSSFSheet sheetCreate1 = workBookOutput1.createSheet();
//			XSSFRow rowCreated = null;
//
//			int totalNumberOfRowsOfNewSheet = UniqueSet.size();

//			for (int rr = 0; rr <= totalNumberOfRowsOfNewSheet; rr++) {
//				rowCreated = sheetCreate1.createRow(rr);
//
//				for (int c = 0; c < totalNumberOfColumnInExcel1; c++) {
//					rowCreated.createCell(c);
//				}
//			}

//			for (int j = 0; j < totalNumberOfRowsOfNewSheet; j++) {

//				System.out.println("UniqueSet.get(i):"+UniqueSet.get(j));

//				for (int i = 1; i <= totalNumberOfRowsInExcel1; i++) {

//					System.out.println("Excel:"+sheet1.getRow(i).getCell(key1).toString());

//					if(UniqueSet.get(j).equals(sheet1.getRow(i).getCell(key1).toString())) {

//						if (sheet1.getRow(i).getCell(key1).getCellType() == CellType.STRING) {
//							rowCreated.getCell(j).setCellValue(sheet1.getRow(i).getCell(key1).getStringCellValue());
//							break;
//						} else if (sheet1.getRow(i).getCell(key1).getCellType() == CellType.NUMERIC) {
//							rowCreated.getCell(j).setCellValue(sheet1.getRow(i).getCell(key1).getNumericCellValue());
//							break;
//						} else if (sheet1.getRow(i).getCell(key1).getCellType() == CellType.BOOLEAN) {
//							rowCreated.getCell(j).setCellValue(sheet1.getRow(i).getCell(key1).getBooleanCellValue());
//							break;
//						}

//						for (int p = 0, u = 0; p <= totalNumberOfRowsInExcel1; p++) {
//
//							if (sheet1.getRow(p) == null) {
//								continue;
//							} else {
//
//								for (int d = 0; d < totalNumberOfColumnInExcel1; d++) {
//
//									if (sheet1.getRow(p).getCell(d) == null) {
//										continue;
//									} else {
//
//										if (sheet1.getRow(p).getCell(d).getCellType() == CellType.STRING) {
//											rowCreated.getCell(d)
//													.setCellValue(sheet1.getRow(p).getCell(d).getStringCellValue());
//										} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.NUMERIC) {
//											rowCreated.getCell(d)
//													.setCellValue(sheet1.getRow(p).getCell(d).getNumericCellValue());
//										} else if (sheet1.getRow(p).getCell(d).getCellType() == CellType.BOOLEAN) {
//											rowCreated.getCell(d)
//													.setCellValue(sheet1.getRow(p).getCell(d).getBooleanCellValue());
//										}
//
//									}
//								}
//							}
//							u++;
//						}
//					}
//				}	
//			}
//		readExcel(sheetCreate1);

			try {
				System.out.println("Unique Excel1 created");
				String target1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Desktop\\BigexcelFiles\\New\\fileOutput1.xlsx";
				FileOutputStream outputStream1 = new FileOutputStream(target1Path);
				workBook1.write(outputStream1);
				workBook1.close();
			} catch (FileNotFoundException ee) {
				ee.printStackTrace();
			}

//			// Now for same Data
//
//			List<String> intersectSet = sheet1List.stream()
//					.filter(sheet2List::contains).collect(Collectors.toList());
//
//			List<String> intersectSet2 = sheet2List.stream()
//					.filter(sheet1List::contains).collect(Collectors.toList());
//
//			System.out.println("Same From excel 1 ---------------------");
//			
//			for (String string2 : intersectSet) {
//				System.out.println(string2);
//			}
//
//			System.out.println("Same From excel 2 ---------------------");
//
//			for (String string2 : intersectSet2) {
//				System.out.println(string2);
//			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
