package com.readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;	

public class ReadExcel {
	static	FileInputStream fil;
	static XSSFWorkbook wb;

public static void main(String[] args) {
	
	File filepath = new File("./Data Folder/Test Data Financial.xlsx");
	try {
	 fil = new FileInputStream(filepath);//data>> Stream
		
	} catch(FileNotFoundException e ) {
		e.printStackTrace();
	}
	
//Apache POI >> jar to handle excle >>> read the excel data as Stream
	
	try {
	 wb = new XSSFWorkbook(fil);
		
	}catch (IOException e) {
		e.printStackTrace();
		}
	XSSFCell cellValue =wb.getSheetAt(0).getRow(0).getCell(1);
	System.out.println(cellValue.toString());
	
}
}


