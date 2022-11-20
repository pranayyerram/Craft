package readExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelLibrary {

	XSSFWorkbook wb;
	XSSFSheet sh;
	
	public ExcelLibrary() throws Exception {
		String filePath="C:\\\\Users\\\\Pranay\\\\eclipse-workspace\\\\ReadExcel\\\\TestData\\\\Data.xlsx";
		File file=new File(filePath);
		FileInputStream fis=new FileInputStream(file);
		
		wb=new XSSFWorkbook(fis);
		
	}
	
	public void readData(String Sheetname, int row, int col) {
		sh=wb.getSheet(Sheetname);
		
		for(int i=0;i<sh.getLastRowNum();i++) {
			Row row1=sh.getRow(i);
			for(int j=0;j<row1.getLastCellNum();j++) {
				String data=sh.getRow(i).getCell(j).getStringCellValue();
				System.out.println(data);
			}
		}
		
		//sh=wb.getSheet(Sheetname);
		//String dataa=sh.getRow(row).getCell(col).getStringCellValue();
		//System.out.println(dataa);
		
	}
}
