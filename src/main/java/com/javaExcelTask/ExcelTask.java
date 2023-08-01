package com.javaExcelTask;

import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("serial")
public class ExcelTask extends JFrame {

	String path1;
	String path2;
	int key1;
	int key2;
	String folderPath;

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

	private JLabel labelFILE1 = new JLabel("FILE 1 :");
	private JLabel labelFILE2 = new JLabel("FILE 2 :");
	private JLabel labelKEYFILE1 = new JLabel("KEY 1 :");
	private JLabel labelKEYFILE2 = new JLabel("KEY 2 :");
	private JLabel outputFolder = new JLabel("OUTPUT :");
	private JLabel displayFileName1 = new JLabel();
	private JLabel displayFileName2 = new JLabel();
	private JLabel displayOutputFolder = new JLabel();
	private JLabel emptyLabel = new JLabel();
	private JButton buttonFile1 = new JButton("openFile1");
	private JButton buttonFile2 = new JButton("openFile2");
	private JButton buttonOutput = new JButton("openFolder");
	private JTextField field = new JTextField(15);
	private JTextField field2 = new JTextField(15);
	private JButton buttonENTER = new JButton("ENTER");

	public ExcelTask() {
		super("EXCEL TASK");

		setLayout(new GridBagLayout());
		GridBagConstraints constraints = new GridBagConstraints();
		constraints.anchor = GridBagConstraints.WEST;
		constraints.insets = new Insets(10, 10, 10, 10);

		constraints.gridy = 0;
		constraints.gridx = 0;
		this.add(labelFILE1, constraints);

		constraints.gridy = 0;
		constraints.gridx = 1;
		add(buttonFile1, constraints);

		constraints.gridy = 0;
		constraints.gridx = 2;
		add(displayFileName1, constraints);

		buttonFile1.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == buttonFile1) {

					JFileChooser fileChooser = new JFileChooser();

					fileChooser.setCurrentDirectory(
							new File("C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\task\\Excel"));

					int response = fileChooser.showOpenDialog(null);

					if (response == JFileChooser.APPROVE_OPTION) {

						File file = new File(fileChooser.getSelectedFile().getAbsolutePath());
						File file2 = fileChooser.getSelectedFile();

						displayFileName1.setText(file.getName());
						String s = fileChooser.getSelectedFile().getAbsolutePath();
						path1 = s;
//						System.out.println(path1);

					}
				}
			}
		});

		constraints.gridy = 1;
		constraints.gridx = 0;
		add(labelFILE2, constraints);

		constraints.gridy = 1;
		constraints.gridx = 1;
		add(buttonFile2, constraints);

		constraints.gridy = 1;
		constraints.gridx = 2;
		add(displayFileName2, constraints);

		buttonFile2.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == buttonFile2) {

					JFileChooser fileChooser = new JFileChooser();

					fileChooser.setCurrentDirectory(
							new File("C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\task\\Excel")); // sets

					int response = fileChooser.showOpenDialog(null);

					if (response == JFileChooser.APPROVE_OPTION) {
						File file = new File(fileChooser.getSelectedFile().getAbsolutePath());
						File file2 = fileChooser.getSelectedFile();
						displayFileName2.setText(file2.getName());
						String s = fileChooser.getSelectedFile().getAbsolutePath();
						path2 = s;
//						System.out.println(path2);

					}
				}
			}
		});

		constraints.gridy = 3;
		constraints.gridx = 0;
		add(labelKEYFILE1, constraints);

		constraints.anchor = GridBagConstraints.CENTER;
		constraints.gridy = 3;
		constraints.gridx = 2;
		add(field, constraints);

		constraints.gridy = 3;
		constraints.gridx = 1;
		add(emptyLabel, constraints);

		constraints.gridy = 4;
		constraints.gridx = 0;
		add(labelKEYFILE2, constraints);

		constraints.anchor = GridBagConstraints.CENTER;
		constraints.gridy = 4;
		constraints.gridx = 2;
		add(field2, constraints);

		constraints.gridy = 4;
		constraints.gridx = 1;
		add(emptyLabel, constraints);

		constraints.gridx = 0;
		constraints.gridy = 5;
		add(outputFolder, constraints);

		constraints.gridx = 1;
		constraints.gridy = 5;
		add(buttonOutput, constraints);

		constraints.gridx = 2;
		constraints.gridy = 5;
		add(displayOutputFolder, constraints);

		buttonOutput.addActionListener(new ActionListener() {

			@SuppressWarnings("unused")
			public void actionPerformed(ActionEvent e) {
				if (e.getSource() == buttonOutput) {

					JFileChooser fileChooser = new JFileChooser();
					fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

					fileChooser.setCurrentDirectory(
							new File("C:\\Users\\SATYASAH\\OneDrive - Capgemini\\Documents\\task\\Excel")); // sets

					int response = fileChooser.showOpenDialog(ExcelTask.this);

					if (response == JFileChooser.APPROVE_OPTION) {
						File file = new File(fileChooser.getSelectedFile().getAbsolutePath());
						File file2 = fileChooser.getSelectedFile();
						displayOutputFolder.setText(file2.getName());
//						System.out.println(file);
						String s = fileChooser.getSelectedFile().getAbsolutePath();
						folderPath = s;
					} else {
						displayOutputFolder.setText("");
					}
				}
			}
		});

		constraints.gridx = 0;
		constraints.gridy = 6;
		constraints.gridwidth = 3;
		constraints.anchor = GridBagConstraints.CENTER;
		add(buttonENTER, constraints);

		buttonENTER.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent event) {

				key1 = Integer.parseInt(field.getText());
				key2 = Integer.parseInt(field2.getText());

				fetchExcel(path1, path2, key1, key2, folderPath);
				JOptionPane.showMessageDialog(ExcelTask.this, "Excel created", "Excel", JOptionPane.PLAIN_MESSAGE);
			}
		});

		pack();
//		setExtendedState(JFrame.MAXIMIZED_BOTH);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLocationRelativeTo(null);

	}

	public static void main(String[] args) {

		SwingUtilities.invokeLater(new Runnable() {

			public void run() {
				new ExcelTask().setVisible(true);
			}
		});
	}
}
