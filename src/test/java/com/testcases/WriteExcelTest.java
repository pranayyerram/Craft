package com.testcases;

import org.testng.annotations.Test;

import com.writeExcel.writeExcel;

public class WriteExcelTest {
	
	@Test
	public void Writetest() throws Exception {
		writeExcel we=new writeExcel();
		we.WriteExcel("test", "Hello", 3, 9);
	}
}
