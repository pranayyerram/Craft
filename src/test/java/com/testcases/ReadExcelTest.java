package com.testcases;

import org.testng.annotations.Test;

import readExcel.ExcelLibrary;

public class ReadExcelTest {
	
	@Test
	public void TestExcel() throws Exception {
		ExcelLibrary eb=new ExcelLibrary();
		eb.readData("test", 6, 0);
		eb.readData("test", 3, 1);
	}
	
	
}
