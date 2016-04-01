package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SheetDataWriter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportImpl {

	private static final String TEMPLATE = "TEMPLATE.xlsx";
	private static final String OUTPUT = "Output.xlsx";
	private static final String PATH = "files" + File.separator;
	private static final String filename = PATH + TEMPLATE;

	private Workbook workbook;

	public static void main(String[] args) throws IOException {
		new ExportImpl().export(345);
		new ExportImpl().export(12);
	}

	private static void makePoiDir(File dir) {
		if (!dir.exists()) {
			System.out.println("Directory " + dir + " doesn't exist");
			dir.mkdirs();
			System.out.println("Directory created before export");
		}
	}

	void export(int id) throws IOException {
		// createDirectory(id);
		checkDirectory();
		createWorkbook();
		createAndBuildSheets();
		writeToFile();
		deleteSXSSFTempFiles(workbook);
		// resetSystemProperty(initialSystemPropery);

	}

	private void checkDirectory() {
		File dir=new File(System.getProperty("java.io.tmpdir"),"poifiles");
		if (!dir.exists()){
			System.out.println(dir+" doesn't exist.");
			dir.mkdirs();
		}
		System.setProperty("poi.keep.tmp.files", "true");
	}


	private void createDirectory(int id) {
		File poidir = new File(System.getProperty("java.io.tmpdir")
				+ File.separator + "maindir" + File.separator + id
				+ File.separator
				+ Long.toHexString(Double.doubleToLongBits(Math.random())));
		makePoiDir(poidir);
		System.setProperty("java.io.tmpdir", poidir.getAbsolutePath());
	}

	private void createAndBuildSheets() {
		BaseSheet sheet = new MySheet(workbook);
		sheet.build();
		workbook = new SXSSFWorkbook((XSSFWorkbook) workbook);
		sheet = new MyStreamingSheet(workbook);
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

	private void deleteSXSSFTempFiles(Workbook workbook)
			throws FileNotFoundException {
		int numberOfSheets = workbook.getNumberOfSheets();

		for (int i = 0; i < numberOfSheets; i++) {
			Sheet sheetAt = workbook.getSheetAt(i);

			if (sheetAt instanceof SXSSFSheet) {
				try {
					SheetDataWriter sdw;
					sdw = (SheetDataWriter) getPrivateAttribute(sheetAt,
							"_writer");
					File f = (File) getPrivateAttribute(sdw, "_fd");
					System.out.println("Deleting " + f.getAbsolutePath());
					f.delete();

				} catch (NoSuchFieldException e) {
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					e.printStackTrace();
				}

			}
		}

	}

	public static Object getPrivateAttribute(Object containingClass,
			String fieldToGet) throws NoSuchFieldException,
			IllegalAccessException {
		// get the field of the containingClass instance
		Field declaredField = containingClass.getClass().getDeclaredField(
				fieldToGet);
		// set it as accessible
		declaredField.setAccessible(true);
		// access it
		Object get = declaredField.get(containingClass);
		// return it!
		return get;
	}

}
