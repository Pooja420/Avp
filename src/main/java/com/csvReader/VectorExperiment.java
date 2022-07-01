package com.csvReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Vector;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class VectorExperiment {

	public static void main(String[] args) throws IOException {
		
		FileInputStream file= new FileInputStream(new File( "C:\\Excelcode\\tournament.xlsx"));
		@SuppressWarnings("resource")
		XSSFWorkbook workbook= new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
//		XSSFSheet sheet2 = workbook.getSheetAt(1);
//		XSSFRow row;
//		row = sheet2.createRow(0);
//		Cell cell =	row.createCell(0);
//		cell.setCellValue("Pooja");
//		
//		FileOutputStream out = new FileOutputStream(
//        new File("C:\\Excelcode\\tournament2.xlsx"));
//		workbook.write(out);
		
//	    int rowcount=sheet.getLastRowNum();  //return the row count
//		System.out.println(rowcount);
	    int colcount=sheet.getRow(1).getLastCellNum();   //returns column/cell count
//	    System.out.println(colcount);
	    
	    String[] data = new String[colcount];
	    for ( int i = 0; i < colcount; i++)
	    {
	    	data[i] = sheet.getRow(1).getCell(i).toString();
//	    	System.out.println(data[i]);
//		    driver.findElement(By.cssSelector(".event-input")).sendKeys();
//		    driver.findElement(By.id("exampleFormControlTextarea1")).click();
//		    //description//
//		    driver.findElement(By.id("exampleFormControlTextarea1")).sendKeys("Event");
    	}
		// TODO Auto-generated method stub
		Vector<String> integerVector = new Vector<String>();
		integerVector.add(data[15]);
		integerVector.add(data[16]);
		integerVector.add(data[17]);
		integerVector.add(data[18]);
		integerVector.add(data[19]);
		integerVector.add(data[20]);
		integerVector.add(data[33]);
		
		for(int i=0; i<integerVector.size(); i++)
		{
			String s = integerVector.get(i);
			s = s.substring(0, s.length() - 2);
			integerVector.set(i, s);
			
		}
		
		System.out.println(integerVector.get(6));
	}
}
