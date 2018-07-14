package com.readExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CommonFunction {
	
	public static String[][] fetchExcelData(String path) throws IOException
	{
		//This is second comment
		
		File fileexcel = new File(path);
		FileInputStream fis = new FileInputStream(fileexcel);
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		HSSFSheet ws = wb.getSheet("Sheet1");
		
		int rowNum = ws.getLastRowNum();
		System.out.println(rowNum);
		int colNum= ws.getRow(0).getLastCellNum();
		
		String [][] Data = new String[rowNum][colNum];
		
		for( int i=0; i < rowNum ; i++)
		{
			HSSFRow row = ws.getRow(i);
			for(int j=0;j<colNum;j++ )
			{
				HSSFCell cell = row.getCell(j);
				
				String value = cellToString(cell);
				Data[i][j] = value;
				System.out.println(Data[i][j]);
				
			}
		}
		wb.close();
		fis.close();
		
		return Data;
	}

	
	public static String cellToString(HSSFCell cell)
	{
		
		Object result;
		String strReturn = null;
		if(cell == null)
		{
			strReturn="";
		}
		else
		{
			if(cell.getCellTypeEnum() == CellType.STRING)
			{
				result = cell.getStringCellValue();
				strReturn = result.toString();
			}
			else if(cell.getCellTypeEnum() == CellType.NUMERIC)
			{
				result = cell.getNumericCellValue();
				strReturn = result.toString();
			}
			else
			{
				strReturn ="";
			}
			
		}
			
		
		
		return strReturn;
		
	}
	

	}
