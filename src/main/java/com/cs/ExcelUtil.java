package com.cs;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {

	private ExcelUtil() {

	}

	public static void setCellValue(XSSFSheet sheet, int row, int cell, String value) {
		XSSFRow rowObj = sheet.getRow(row);
		if (rowObj != null) {
			XSSFCell cellObj = rowObj.getCell(cell);
			if (cellObj != null) {
				cellObj.setCellValue(value);
			} else {
				rowObj.createCell(cell).setCellValue(value);
			}
		} else {
			sheet.createRow(row).createCell(cell).setCellValue(value);
		}
	}

	public static void mergeCell(XSSFSheet sheet, int startRow, int endRow, int startCell, int endCell) {
		sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, startCell, endCell));
	}

	public static XSSFCellStyle createCellStyle(XSSFWorkbook workbook, int fontsize) {
		XSSFCellStyle style = workbook.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		XSSFFont font = workbook.createFont();
		font.setFontHeightInPoints((short) fontsize);
		style.setFont(font);
		return style;
	}

	public static void setStyleForCell(XSSFSheet sheet, XSSFCellStyle style, int row, int cell) {
		XSSFRow rowObj = sheet.getRow(row);
		if (rowObj != null) {
			XSSFCell cellObj = rowObj.getCell(cell);
			if (cellObj != null) {
				cellObj.setCellStyle(style);
			} else {
				rowObj.createCell(cell).setCellStyle(style);
			}
		} else {
			sheet.createRow(row).createCell(cell).setCellStyle(style);
		}
	}

	public static void addImageToSheet(XSSFWorkbook workBook, XSSFSheet sheet, byte[] image, int row1, int cell1,
			int row2, int cell2) {
		XSSFDrawing d = sheet.createDrawingPatriarch();
		XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, row1, cell1, row2, cell2);
		d.createPicture(anchor, workBook.addPicture(image, HSSFWorkbook.PICTURE_TYPE_JPEG));
	}

	public static byte[] getImage(String filePath) {
		try (ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();) {
			BufferedImage bufferImg = ImageIO.read(new File(filePath));
			ImageIO.write(bufferImg, "png", byteArrayOut);
			return byteArrayOut.toByteArray();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return new byte[0];
	}

}
