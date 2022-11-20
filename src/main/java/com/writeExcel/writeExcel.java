package com.writeExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeExcel {
	XSSFSheet sheet;
	public void WriteExcel(String Sheetname, String cellValue, int row,int col) throws Exception {
		String filePath="C:\\\\Users\\\\Pranay\\\\eclipse-workspace\\\\ReadExcel\\\\TestData\\\\Data.xlsx";
		File file=new File(filePath);
		FileInputStream fis=new FileInputStream(file);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		sheet=wb.getSheet(Sheetname);
		//sheet.getSheet(filePath).getRow(row).createCell(col).setCellValue(cellValue);
		sheet.getRow(row).createCell(col).setCellValue(cellValue);
		FileOutputStream fos=new FileOutputStream(new File(filePath));
		wb.write(fos);
		wb.close();
	}
	
}
