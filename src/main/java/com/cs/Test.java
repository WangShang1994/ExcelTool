package com.cs;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {

	public static void main(String[] args) {
		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet = workBook.createSheet("Test");
		sheet.createDrawingPatriarch();
		ExcelUtil.mergeCell(sheet, 0, 0, 0, 15);
		ExcelUtil.setCellValue(sheet, 0, 0, "Title Test Execl");
		XSSFCellStyle style = ExcelUtil.createCellStyle(workBook, 15);
		ExcelUtil.setStyleForCell(sheet, style, 0, 0);
		ExcelUtil.addImageToSheet(workBook, sheet, ExcelUtil.getImage("D:\\ExcelTest\\screen.jpg"), 1, 1, 10, 10);
		try (FileOutputStream fos = new FileOutputStream(new File("D:\\ExcelTest\\hello.xlsx"));) {
			workBook.write(fos);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
