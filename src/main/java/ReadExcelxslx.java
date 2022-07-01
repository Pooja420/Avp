


import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;


public class ReadExcelxslx 

{

	public static void main(String[] args)
	{

    	System.setProperty("webdriver.chrome.driver", "C:\\Users\\pooja\\eclipse-workspace\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = (WebDriver) new ChromeDriver();
		
		driver.get("https://promoter.applination.in");
		
		try 
		{
			
			FileInputStream file= new FileInputStream(new File( "C:\\Excelcode\\tournament.xlsx"));
			@SuppressWarnings("resource")
			XSSFWorkbook workbook= new XSSFWorkbook(file);
			XSSFSheet sheet= workbook.getSheet("Sheet1"); 
			Iterator<Row> rowIterator= sheet.iterator();
			while(rowIterator.hasNext())
              {
				Row row= rowIterator.next();
				Iterator<Cell>cellIterator= row.cellIterator();
				while(cellIterator.hasNext())
				{
					Cell cell= (Cell )cellIterator.next();
					switch(cell.getCellType())
					{
					case NUMERIC:
					 System.out.print(cell.getNumericCellValue()+"\t");
					break;			    
					case STRING :System.out.print(cell.getStringCellValue()+"\t");
				                  break;
					
					}
				}
				
            	 System.out.println(""); 
			}
			file.close();
		} 
		
		catch (Exception e)
		{
			e.printStackTrace();
			System.out.println();
		}
	}
}

