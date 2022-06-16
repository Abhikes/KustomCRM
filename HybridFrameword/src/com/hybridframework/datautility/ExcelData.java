package com.hybridframework.datautility;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Properties;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelData {
   //DataFetch
	public String fetchDataExcel(String sheet,int row,int cell,String path) throws Throwable
    {
    	FileInputStream fis = new FileInputStream(path);
    	Workbook book = WorkbookFactory.create(fis);
    	Sheet sh = book.getSheet(sheet);
    	DataFormatter format= new DataFormatter();
    	String data=format.formatCellValue(sh.getRow(row).getCell(cell));
    	return data;
    }
    public String fetchDataProp(String key,String path) throws Throwable
    {
    	FileInputStream fis = new FileInputStream(path);
    	Properties pobj = new Properties();
    	pobj.load(fis);
    	String data =pobj.getProperty(key);
		
    	return data;
    	
    }
}
