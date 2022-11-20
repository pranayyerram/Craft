package readExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.testng.annotations.Test;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcel {
	
	@Test
	public void readExcel() throws Exception {
		
		String excelPath="C:\\Users\\Pranay\\eclipse-workspace\\ReadExcel\\TestData\\Data.xlsx";
		String fileName="TestData";
		String sheetName="Test";
		
		File file=new File(excelPath);
		FileInputStream fis=new FileInputStream(file);
		
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet(sheetName);
		
		int rowCount = sheet.getLastRowNum();
		System.out.println("Total Number of rows: "+rowCount);
		
		String data=sheet.getRow(5).getCell(0).getStringCellValue();
		System.out.println(data);
		
		for(int i=0;i<=rowCount;i++) {
			Row row=sheet.getRow(i);
			for(int j=0;j<row.getLastCellNum();j++) {
				String data1=sheet.getRow(i).getCell(j).getStringCellValue();
				System.out.print(data1+" ");
			}
			System.out.println();
		}
	}
	
}
