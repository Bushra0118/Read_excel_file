package com.readexcel;

/*import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.LinkedHashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelColumnValue {

	static FileInputStream fis;
	static XSSFWorkbook wb;

	public static void main(String[] args) {*/

		/*File filepath = new File("./Data Folder/Test Data Financial.xlsx");

		try {
			fis = new FileInputStream(filepath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		// Apache POI
		try {
			wb = new XSSFWorkbook(fis);
		} catch (IOException e) {
			e.printStackTrace();
		}

		// column code >>> 1) all rows/row count(?)

		XSSFSheet sht = wb.getSheetAt(0);
		int RowCount = sht.getPhysicalNumberOfRows();// row count/rownum/row /number/total row
		System.out.println("Row number =" + RowCount);// 2

		// loop through all rows
		int columnCount = 0;
		Set<String> mySet = new LinkedHashSet<>();
		for (int i = 0; i < RowCount; i++) {// loop for row

			XSSFRow row = sht.getRow(i);

			columnCount = row.getLastCellNum();

			for (int j = 0; j < columnCount; j++) {// loop for column
				XSSFCell cellValue = row.getCell(j);
				DataFormatter df = new DataFormatter();

				mySet.add(df.formatCellValue(cellValue).toString());

			}

		}

		System.out.println("Column number =" + columnCount);// 3
		System.out.println(mySet);
	}
*/
		
		

		import java.io.File;
		import java.io.FileInputStream;
		import java.util.LinkedHashSet;
		import java.util.Set;

		import org.apache.poi.ss.usermodel.Cell;
		import org.apache.poi.ss.usermodel.DataFormatter;
		import org.apache.poi.xssf.usermodel.XSSFRow;
		import org.apache.poi.xssf.usermodel.XSSFSheet;
		import org.apache.poi.xssf.usermodel.XSSFWorkbook;

		public class ReadExcelColumnValue {
			static FileInputStream fis;
			static XSSFWorkbook wb;

			public static void main(String[] args) {

				File filepath = new File("./Data Folder/Test Data Financial.xlsx");

				/*
				 * try { fis = new FileInputStream(filepath); } 
				 * catch (FileNotFoundException e)
				 * {  e.printStackTrace(); } 
				 * // Apache POI try
				 * { wb = new XSSFWorkbook(fis); } catch (IOException e) {  e.printStackTrace(); }
				 */

				try {
					fis = new FileInputStream(filepath);
					wb = new XSSFWorkbook(fis);
				} catch (Exception e) {
					e.printStackTrace();
				}

				// column code >> all row

				XSSFSheet sht = wb.getSheetAt(0);
				// Row row = sht.getRow(0);//for 1 row
				int RowCount = sht.getPhysicalNumberOfRows();// total row num/row count
				System.out.println(RowCount);

				// loop through all rows

				int columnCount = 0;
				Set<String> myset = new LinkedHashSet<String>();
				
				for (int i = 0; i < RowCount; i++) {
					
					XSSFRow row = sht.getRow(1);
					columnCount = row.getLastCellNum();

					for (int j = 0; j < columnCount; j++) {
						Cell cellvalue = row.getCell(j);
						// myset.add(cellvalue.toString());
						DataFormatter df = new DataFormatter();
						myset.add(df.formatCellValue(cellvalue).toString());

						/*
						 * Cell cellvalue = row.getCell(0);//to print specific column
						 * System.out.println(cellvalue);//to print everything
						 */ }
					// to print one time only we have to print it outside of the loop
					// Cell cellvalue = row.getCell(0);//to print specific column
					// System.out.println(cellvalue);//to print e
				}
					System.out.println("Column number= " + columnCount);
					System.out.println(myset);
				
			}	
		
}
