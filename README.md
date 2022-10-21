# SagaBhagat

package com.subhash;

import java.awt.FileDialog;
import java.awt.Frame;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToText {

	private static XSSFWorkbook workbook;

	public static void main(String[] args) throws IOException {
		System.out.println("Started");
		FileDialog dialog = new FileDialog((Frame) null, "Select Excel");
		dialog.setMode(FileDialog.LOAD);
		dialog.setDirectory("C:\\Users\\" + System.getProperty("user.name") + "\\Desktop");
		dialog.setVisible(true);
		String directory = dialog.getDirectory();
		String excelFileName = dialog.getFile();
		String path = directory + excelFileName.split("\\.")[0];
		new File(path).mkdir(); // create folder with name of excel
		FileInputStream fileInputStream = new FileInputStream(new File(directory + excelFileName));
		workbook = new XSSFWorkbook(fileInputStream);

		int noOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < noOfSheets; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			String sheetName = sheet.getSheetName();
			String subPath = path + "/" + sheetName;
			new File(subPath).mkdir(); // create sub folder with name of sheet of excel
			int noOfColumns = sheet.getRow(0).getLastCellNum();

			for (int j = 0; j < noOfColumns; j++) {
				DataFormatter df = new DataFormatter();
				if (df.formatCellValue(sheet.getRow(0).getCell(j)).equals("")) {
					break;
				}
				String textFilePath = subPath + "/" + df.formatCellValue(sheet.getRow(0).getCell(j)) + ".txt";
				String textFilePath2 = subPath + "/" + df.formatCellValue(sheet.getRow(0).getCell(j)) + "_Quotes.txt";
				File file = new File(textFilePath);
				File file2 = new File(textFilePath2);

				if (!file.exists()) {
					file.createNewFile();
				}
				if (!file2.exists()) {
					file2.createNewFile();
				}
				FileOutputStream fos = new FileOutputStream(file);
				FileOutputStream fos2 = new FileOutputStream(file2);
				String cellValue = null;
				String data = null;
				String data2 = null;
				byte[] contentInBytes;
				int noOfRows = sheet.getLastRowNum();
				Row row = null;
				for (int k = 1; k <= noOfRows; k++) {

					row = sheet.getRow(k);
					cellValue = df.formatCellValue(row.getCell(j));

					if (k != noOfRows) {
						data = cellValue + ",";
						data2 = "\"" + cellValue + "\",";
					} else {
						data = cellValue;
						data2 = "\"" + cellValue + "\"";
					}
					contentInBytes = data.getBytes();
					fos.write(contentInBytes);
					contentInBytes = data2.getBytes();
					fos2.write(contentInBytes);
				}
				fos.flush();
				fos.close();
				fos2.flush();
				fos2.close();
			}
		}

		workbook.close();
		System.out.println("Stopped");
	}
}
