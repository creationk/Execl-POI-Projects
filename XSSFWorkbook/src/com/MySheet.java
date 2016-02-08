package com;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

public class MySheet {

	private Workbook workbook;

	public MySheet(Workbook workbook) {
		this.workbook = workbook;
	}

	public void build() {

		String sheetName = "Sheet1";
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

	protected void setOrCreateTextCell(Row row, int j, String text) {
		Cell cell = getOrCreateCell(row, j);
		setCell(cell, text);
	}

	protected Cell getOrCreateCell(Row row, int j) {
		Cell cell = row.getCell(j);
		if (cell == null) {
			cell = row.createCell(j);
		}
		return cell;
	}

	protected Row getOrCreateRow(Sheet sheet, int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}
		return row;
	}

	protected Sheet getOrCreateSheet(String sheetName) {
		Sheet sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			sheet = workbook.createSheet(sheetName);
		}
		return sheet;
	}

	protected void setCell(Cell cell, Object value) {
		if (value == null) {
			cell.setCellValue(new XSSFRichTextString());
		} else if (value.toString().equals("")) {
			cell.setCellValue(new XSSFRichTextString(""));
		} else {
			cell.setCellType(Cell.CELL_TYPE_STRING);
			cell.setCellValue(new XSSFRichTextString(value.toString()));
		}
	}

}
