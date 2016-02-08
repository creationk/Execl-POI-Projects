package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportImpl {
	private static final String TEMPLATE = "TEMPLATE.xlsx";
	private static final String OUTPUT = "Output.xlsx";
	private static final String PATH = "files//";
	private static final String filename = PATH + TEMPLATE;

	private Workbook workbook;

	public static void main(String[] args) throws IOException {
		new ExportImpl().export();
	}

	private void export() throws IOException {
		createWorkbook();
		createAndBuildSheet();
		writeToFile();
	}

	private void createAndBuildSheet() {
		MySheet sheet = new MySheet(workbook);
		sheet.build();
	}

	private void writeToFile() throws IOException {
		FileOutputStream fos = new FileOutputStream(new File(PATH + OUTPUT));
		workbook.write(fos);
		fos.flush();
		fos.close();
	}

	private void createWorkbook() throws IOException {
		InputStream inputStream = new FileInputStream(filename);
		workbook = new XSSFWorkbook(inputStream);
		inputStream.close();
	}

}
