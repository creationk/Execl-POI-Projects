package com;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class MySheet extends BaseSheet {

	public MySheet(Workbook workbook) {
		super(workbook);
	}

	public void build() {

		String sheetName = "MySheet";
		Sheet sheet = getOrCreateSheet(sheetName);

		int numberOfRows = 5;
		int numOfColumns = 3;

		for (int rowIndex = 0; rowIndex < numberOfRows; rowIndex++) {
			Row row = getOrCreateRow(sheet, rowIndex);
			for (int colIndex = 0; colIndex < numOfColumns; colIndex++) {
				setOrCreateTextCell(row, colIndex, rowIndex * colIndex + "");
			}
		}
	}
}
