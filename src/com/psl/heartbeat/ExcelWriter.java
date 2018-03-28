package com.psl.heartbeat;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

	static int rowCount;
	static Workbook workbook = null;
	static Sheet sheet = null;
	static FileOutputStream outputStream = null;
	static String excelSheetPath = System.getProperty("user.dir")+"/Reports/Automation_Status.xlsx";
	static String propertyFilePath = System.getProperty("user.dir")+"/Resource/excelColumnName.properties";
	
	public static void main(String[] args) {

		createExcelFile();
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		Date startDate = new Date();
		String startTime = dateFormat.format(startDate);
		System.out.println(dateFormat.format(startDate));
		Date endDate = new Date();
		String endTime = dateFormat.format(endDate);
		Long timeStamp = endDate.getTime() - startDate.getTime();
		try {
			writeDataIntoExcel("testCaseName1", "action", "label", "isSuccess", "statusFailOrPass", "exception",
					"screenShotStatus", "screenShotPath", startTime, endTime, timeStamp);
			
			writeDataIntoExcel("testCaseName2", "action2", "label2", "isSuccess2", "statusFailOrPass2", "exception2",
					"screenShotStatus2", "screenShotPath2", startTime, endTime, timeStamp);
						
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static void createExcelFile() {

		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("Automation Status");
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 11);
		font.setFontName(HSSFFont.FONT_ARIAL);
		font.setBoldweight(HSSFFont.COLOR_NORMAL);
		((XSSFFont) font).setBold(true);
		font.setColor(HSSFColor.DARK_BLUE.index);
		style.setFont(font);
		style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		writeExcelColumnName(style);
	}

	public static void writeExcelColumnName(CellStyle style) {

		Properties prop = new Properties();
		InputStream input = null;
		try {
			input = new FileInputStream(propertyFilePath);
			prop.load(input);
			System.out.println(prop.getProperty("isSuccess"));
		} catch (IOException ex) {
			ex.printStackTrace();
		} finally {
			if (input != null) {
				try {
					input.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		if (rowCount == 0) {
			int column = 0;
			Row rowHead = sheet.createRow(rowCount++);
			Cell cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("TestCaseName"));
			cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("Action"));
			cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("Label"));
			cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("isSuccess"));
			cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("Status_Pass_or_Fail"));
			cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("Exception"));
			cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("Screenshot_Status"));
			cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("Screenshot_Path"));
			cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("Start_Time"));
			cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("End_Time"));
			cell = rowHead.createCell(column++);
			cell.setCellStyle(style);
			cell.setCellValue(prop.getProperty("Timestamp"));

		}

	}

	public static void writeDataIntoExcel(String testCaseName, String action, String label, String isSuccess,
			String statusFailOrPass, String exception, String screenShotStatus, String screenShotPath, String startTime,
			String endTime, Long timeStamp) {

		try {
			if (outputStream != null) {
				File file = new File(excelSheetPath);
				FileInputStream fileInputStream = new FileInputStream(file);
				workbook = new XSSFWorkbook(fileInputStream);
				sheet = workbook.getSheetAt(0);
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}finally{
			try {
				outputStream = new FileOutputStream(excelSheetPath);
				workbook.write(outputStream);
				workbook.close();
			} catch (IOException e) {				
				e.printStackTrace();
			}
		}
	}

	public static void assignCellValueIntoExcel(String testCaseName, String action, String label, String isSuccess,
			String statusFailOrPass, String exception, String screenShotStatus, String screenShotPath, String startTime,
			String endTime, Long timeStamp) {
		
		int rowCount = sheet.getLastRowNum();
		Row row = sheet.createRow(++rowCount);
		int columnCount = 0;
		Cell cell = row.createCell(columnCount++);
		cell.setCellValue(testCaseName);
		cell = row.createCell(1);
		cell.setCellValue(action);
		cell = row.createCell(2);
		cell.setCellValue(label);
		cell = row.createCell(3);
		cell.setCellValue(isSuccess);
		cell = row.createCell(4);
		cell.setCellValue(statusFailOrPass);
		cell = row.createCell(5);
		cell.setCellValue(exception);
		cell = row.createCell(6);
		cell.setCellValue(screenShotStatus);
		cell = row.createCell(7);
		cell.setCellValue(screenShotPath);
		cell = row.createCell(8);
		cell.setCellValue(startTime);
		cell = row.createCell(9);
		cell.setCellValue(endTime);
		cell = row.createCell(10);
		cell.setCellValue(timeStamp);
		
		int numberOfSheets = workbook.getNumberOfSheets();

		for (int i = 0; i < numberOfSheets; i++) {
			sheet = workbook.getSheetAt(i);
			if (sheet.getPhysicalNumberOfRows() > 0) {
				Row rowNum = sheet.getRow(0);
				Iterator<Cell> cellIterator = rowNum.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cellNum = cellIterator.next();
					int columnIndex = cellNum.getColumnIndex();
					sheet.autoSizeColumn(columnIndex);
				}
			}
		}		
	}
	
}