package excelops;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelcreator {

	public static void main(String args[]) throws IOException, illegalSheetIndexException {

		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("mysheet");
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("1. cell");

		cell = row.createCell(1);
		DataFormat format = wb.createDataFormat();
		CellStyle dtStyle = wb.createCellStyle();
		dtStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
		cell.setCellStyle(dtStyle);
		cell.setCellValue(new Date());

		row.createCell(2).setCellValue("karthik");

		sheet.autoSizeColumn(2);

		wb.write(new FileOutputStream("exl.xlsx"));
		wb.close();

		XSSFWorkbook wb1 = new XSSFWorkbook(new FileInputStream("exl.xlsx"));
		XSSFSheet sheet1 = wb1.getSheetAt(0);
		XSSFRow row1 = sheet1.getRow(0);
		System.out.println(row1.getCell(0).getStringCellValue());
		System.out.println(row1.getCell(1).getDateCellValue());
		System.out.println(row1.getCell(2).getStringCellValue());
		

	}
}
