package com.sjtu.luo;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class TestXlsxRead {
	public static void main(String[] args) throws IOException, InvalidFormatException {

		Workbook wb = WorkbookFactory.create(new FileInputStream(
				"excel/1.xlsx"));
		Sheet sheet = wb.getSheetAt(0);
		int count = 0;
		for (Row row : sheet) {
			count++;
			
		}
		System.out.println("total:" + count);
	}
}
