package com.app.libraries;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelLib
{
	public String readData(String filePath,String sheetname,int rownum,int cellno) throws EncryptedDocumentException, InvalidFormatException, IOException 
	{
		FileInputStream fis =new FileInputStream(filePath);
		Workbook wb =WorkbookFactory.create(fis);
		return wb.getSheet(sheetname).getRow(rownum).getCell(cellno).getStringCellValue();
	}
	
		
	public void writeData(String filePath,String sheetname,int rownum,int cellno,String data) throws Exception, InvalidFormatException, IOException 
	{
		FileInputStream fis =new FileInputStream(filePath);
		 Workbook wb =WorkbookFactory.create(fis);
		//XSSFWorkbook workbook = new XSSFWorkbook(fis);
		wb.getSheet("Sheet1").getRow(rownum).createCell(4).setCellValue(data);
		FileOutputStream fos =new FileOutputStream(filePath);
		wb.write(fos);
	    fos.close();
	}
	
	public int TotalRows(String filePath,String sheetname) throws Exception
	{
		FileInputStream fis =new FileInputStream(filePath);
		Workbook wb =WorkbookFactory.create(fis);
		return wb.getSheet(sheetname).getLastRowNum();
		
	}
	
	public int TotalCells(String filePath,String sheetname,int rownum) throws Exception
	{
		FileInputStream fis =new FileInputStream(filePath);
		Workbook wb =WorkbookFactory.create(fis);
		return wb.getSheet(sheetname).getRow(rownum).getLastCellNum();
		
	}

}
