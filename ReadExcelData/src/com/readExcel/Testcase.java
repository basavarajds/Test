package com.readExcel;

import java.io.IOException;

public class Testcase extends CommonFunction {
//THis is modified by basavaraj
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String path = "D:\\Keyword_Driven_Framework\\ReadExcelData\\ReadExcelData.xls";
		String data[][] = fetchExcelData( path);
		
		int rowcount = data.length;
		
		System.out.println(rowcount);

	}

}
