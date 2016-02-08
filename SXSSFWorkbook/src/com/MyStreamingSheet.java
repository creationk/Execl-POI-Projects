package com;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

public class MyStreamingSheet extends BaseSheet {
	
	public MyStreamingSheet(Workbook workbook) {
		super(workbook);
	}

	public void build() {

		String sheetName = "MyStreamingSheet";
		Sheet sheet = getOrCreateSheet(sheetName);

		int numberOfRows = 50;
		int numOfColumns = 3;

		for (int rowIndex = 0; rowIndex < numberOfRows; rowIndex++) {
			if ((rowIndex % 10 == 0)&&(rowIndex>0)) {
				try {
					((SXSSFSheet) sheet).flushRows(10);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			Row row = getOrCreateRow(sheet, rowIndex);
			for (int colIndex = 0; colIndex < numOfColumns; colIndex++) {
				setOrCreateTextCell(row, colIndex, rowIndex * colIndex + "");
			}
		}
	}
}
