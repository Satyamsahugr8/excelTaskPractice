package com.javaExcelPractice;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPractice {

	public static void readExcel() throws IOException {

		String excel1Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file1.xlsx";
		FileInputStream file = new FileInputStream(excel1Path);
		XSSFWorkbook workBook = new XSSFWorkbook(file);
		XSSFSheet sheet = workBook.getSheetAt(0);

		String excel2Path = "C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\Task\\Excel\\Demo\\file2.xlsx";
		FileInputStream file1 = new FileInputStream(excel2Path);
		XSSFWorkbook workBook2 = new XSSFWorkbook(file1);
		XSSFSheet sheet2 = workBook2.getSheetAt(0);

		// workBook1
		int totalNumberOfRows1 = sheet.getLastRowNum();
		int totalNumberOfColumn1 = sheet.getRow(0).getLastCellNum();

		System.out.println("totalNumberOfRows1:" + totalNumberOfRows1);
		System.out.println("totalNumberOfColumn1:" + totalNumberOfColumn1);

		int key1 = 0;

		System.out.println("Set key as key1:" + key1);

		XSSFRow row11 = sheet.getRow(0);
		XSSFRow row12;
		XSSFCell cell1;
		XSSFCell cellOfRowKey1;

		// workBook2
		int totalNumberOfRows2 = sheet2.getLastRowNum();
		int totalNumberOfColumn2 = sheet2.getRow(0).getLastCellNum();

		System.out.println("totalNumberOfRows2:" + totalNumberOfRows2);
		System.out.println("totalNumberOfColumn2:" + totalNumberOfColumn2);

		int key2 = 1;

		System.out.println("Set key as key2:" + key2);

		XSSFRow firstRowOfExcel2 = sheet2.getRow(0);
		XSSFRow row22;
		XSSFCell firstColumn;
		XSSFCell cellOfRowKey2;
		double value;

		
		
		
		// workBook1 Logic
		for (int c = 0; c < totalNumberOfColumn1; c++) {
			cell1 = row11.getCell(c);
//			System.out.println(cell);
			if (cell1.getColumnIndex() == key1) {
				System.out.println("cell1:" + cell1);
				for (int r = 1; r <= totalNumberOfRows1; r++) {
					
					cellOfRowKey1 = sheet.getRow(r).getCell(cell1.getColumnIndex());
					
					System.out.println(cellOfRowKey1);
					row12 = cellOfRowKey1.getRow();
//					System.out.println(row12);
				}
			}
		}

		// workBook2 Logic
		for (int c = 0; c < totalNumberOfColumn2; c++) {
			firstColumn = firstRowOfExcel2.getCell(c);
//			if (firstRowOfExcel2.getCell(c).getColumnIndex() == key2) {
			if (firstColumn.getColumnIndex() == key2) {
//				firstColumn = firstRowOfExcel2.getCell(c);
//				System.out.println("cellOfFirstRow:" + firstColumn);
				for (int r = 1; r <= totalNumberOfRows2; r++) {
					cellOfRowKey2 = sheet2.getRow(r).getCell(firstColumn.getColumnIndex());
					System.out.println(cellOfRowKey2);
					row22 = cellOfRowKey2.getRow();
//					System.out.println(row22);
				}
			}
		}
	} // readExcel

//				switch (cell.getCellType()) {
//				case STRING:
//					System.out.println(cell.getStringCellValue());
//					break;
//				case NUMERIC:
//					System.out.print(cell.getNumericCellValue());
//					break;
//				case BOOLEAN:
//					System.out.print(cell.getBooleanCellValue());
//					break;
//				default:
//					System.out.println("HI");
//					System.out.println();
//				} // switch
//		} // for
//	} // for

//		// Iterator
//		Iterator<Row> rows = sheet.iterator();
//
//		while (rows.hasNext()) {
//
//			Row currRow = rows.next();
//
//			Iterator<Cell> cells = currRow.cellIterator();
//
//			while (cells.hasNext()) {
//
//				Cell currCell = cells.next();
//				CellType cellType = currCell.getCellType();
//
//				String value = "";
//
//				if (cellType == CellType.STRING) {
//					value = "" + currCell.getStringCellValue();
//				} else if (cellType == CellType.NUMERIC) {
//					value = "" + currCell.getNumericCellValue();
//				}
//				System.out.print(" " + value);
//			}
//			System.out.println();
//		}

//		// geting first row as string
//		
//		int rows = sheet.getLastRowNum();
//		int column = sheet.getRow(0).getLastCellNum();
//		
//		
//		String[] rowHeader = new String[column];
//		for (int r = 0; r < 1; r++) {
//			XSSFRow row = sheet.getRow(0);
//			for (int c = 0; c < column; c++) {
//				XSSFCell cell = row.getCell(c);
//				
	// 1
//				System.out.println("cell"+cell);
//				rowHeader[c] = ""+cell;

//				//2
//				switch (cell.getCellType()) {
//				case STRING:
////					System.out.println(cell.getStringCellValue());
//					rowHeader[c] = cell.getStringCellValue();
//					break;
//				case NUMERIC:
////					System.out.print(cell.getNumericCellValue());
//					rowHeader[c] = ""+cell.getNumericCellValue();
//					break;
//				case BOOLEAN:
//					System.out.print(cell.getBooleanCellValue());
//					break;
////				default:
////					System.out.println("default");
////					System.out.println();
//				} // switch
//				
//			} // for
//		} // for

//		
//		
//		String[] bookTitles = new String[] { "Effective Java", "Head First Java", "Thinking in Java",
//				"Java for Dummies" };
//      String rowHeader2[] = {"Emp ID","Global ID","22.0","Grade","Screener","Mobile no.","Interested ?","technology"};

//		JComboBox<String> bookList = new JComboBox<>(bookTitles);

////    add more books
//		bookList.addItem("Java Generics and Collections");
//		bookList.addItem("Beginnning Java 7");
//		bookList.addItem("Java I/O");
//
//		JButton comboButton = new JButton("Show ComboBox selected");
//		JLabel labelCombo = new JLabel("ComboBox is");
//
//		comboButton.addActionListener(new ActionListener() {
//			@Override
//			public void actionPerformed(ActionEvent evt) {
//				// do everything here...
//				// get the selected item:
//				String selectedBook = (String) bookList.getSelectedItem();
//				System.out.println("You seleted the book: " + selectedBook);
//				labelCombo.setText(selectedBook + " ");
//			}
//		});

	// 6
//				bookList.addActionListener(new ActionListener() { 
//				    @Override
//				    public void actionPerformed(ActionEvent event) {
//				        JComboBox<String> combo = (JComboBox<String>) event.getSource();
//				        String selectedBook = (String) combo.getSelectedItem();
//				 
//				        if (selectedBook.equals("Effective Java")) {
//				            System.out.println("Good choice!");
//				        } else if (selectedBook.equals("Head First Java")) {
//				            System.out.println("Nice pick, too!");
//				        }
//				    }
//				});
//				
//				
////				bookList.setForeground(Color.BLUE);
//				bookList.setBackground(Color.WHITE);
//				bookList.setFont(new Font("Arial", Font.BOLD, 14));
//				// And limit the maximum number of items displayed in the drop-down list:
//				bookList.setMaximumRowCount(5); // scroller

	public static void main(String[] args) {
		try {
			readExcel();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
