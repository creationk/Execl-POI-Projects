package com;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

public abstract class BaseSheet {

	private Workbook workbook;

	public BaseSheet() {
	}
	
	public BaseSheet(Workbook workbook) {
		this.workbook = workbook;
	}

	public abstract void build();

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
