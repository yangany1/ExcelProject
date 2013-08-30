package com.sjtu.luo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

public class DocTest {

	public static void main(String[] args) {

		// createNewExcel();
		// createNewSheet();
		travelExcel();
		// textExtraction();
	}

	// 创建一个Excel
	public static void createNewExcel() {
		try {
			Workbook wb = new HSSFWorkbook();
			FileOutputStream fileOut = new FileOutputStream("excel/test.xls");
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 创建sheet
	public static void createNewSheet() {
		try {
			Workbook wb = new HSSFWorkbook(); // or new XSSFWorkbook();
			Sheet sheet1 = wb.createSheet("new sheet");
			Sheet sheet2 = wb.createSheet("second sheet");
			String safeName = WorkbookUtil
					.createSafeSheetName("[O'Brien's sales*?]"); // returns
																	// " O'Brien's sales   "
			Sheet sheet3 = wb.createSheet(safeName);
			FileOutputStream fileOut = new FileOutputStream("excel/test.xls");
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 遍历一个Excel
	public static void travelExcel() {
		POIFSFileSystem fs;
		try {
			fs = new POIFSFileSystem(new FileInputStream("excel/3.xls"));
			HSSFWorkbook wb = new HSSFWorkbook(fs);

			// 读取sheet的个数
			System.out.println("sheet number=" + wb.getNumberOfSheets());
			// 选择sheet
//			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				Sheet sheet = wb.getSheetAt(0);
				// 按行列遍历sheet
				for (Iterator<Row> rit = sheet.rowIterator(); rit.hasNext();) {
					Row row = rit.next();
					for (Iterator<Cell> cit = row.cellIterator(); cit.hasNext();) {
						Cell cell = cit.next();
						CellReference cellRef = new CellReference(
								row.getRowNum(), cell.getColumnIndex());
						// System.out.print(cellRef.formatAsString());
						// System.out.print(" - ");
						// 判断cell中数据的类型

						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING:
							System.out.print(cell.getRichStringCellValue()
									.getString());
							break;
						case Cell.CELL_TYPE_NUMERIC:
							if (DateUtil.isCellDateFormatted(cell)) {
								System.out.print(cell.getDateCellValue());
							} else {
								System.out.print(cell.getNumericCellValue());
							}
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							System.out.print(cell.getBooleanCellValue());
							break;
						case Cell.CELL_TYPE_FORMULA:
							System.out.print(cell.getCellFormula());
							break;
						default:
							System.out.println();
						}

						System.out.print(" ");
					}
					System.out.println();
				}
//			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// for (Row row : sheet) {
		// for (Cell cell : row) {
		// // Do something here
		// }
		// }
	}

	// Text Extraction
	public static void textExtraction() {
		InputStream inp;
		try {
			inp = new FileInputStream("excel/2.xls");
			HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
			ExcelExtractor extractor = new ExcelExtractor(wb);
			extractor.setFormulasNotResults(true);
			extractor.setIncludeSheetNames(false);
			String text = extractor.getText();
			System.out.println(text);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
}
