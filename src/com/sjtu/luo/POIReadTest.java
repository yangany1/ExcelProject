package com.sjtu.luo;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;



public class POIReadTest {
	public static void main(String[] args) throws Exception {
		HSSFWorkbook wb = null;
		POIFSFileSystem fs = null;
		try {
			fs = new POIFSFileSystem(new FileInputStream("excel/1.xls"));
			wb = new HSSFWorkbook(fs);
		} catch (IOException e) {
			e.printStackTrace();
		}

		HSSFSheet sheet = wb.getSheetAt(0);
		HSSFRow row = sheet.getRow(0);
		HSSFCell cell = row.getCell(0);
		String msg = cell.getStringCellValue();
		System.out.println(msg);
	}

	public static void method2() throws Exception {

		InputStream is = new FileInputStream("e:\\workbook.xls");
		HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(is));

		ExcelExtractor extractor = new ExcelExtractor(wb);
		extractor.setIncludeSheetNames(false);
		extractor.setFormulasNotResults(false);
		extractor.setIncludeCellComments(true);

		String text = extractor.getText();
		System.out.println(text);
	}

	
	
}
