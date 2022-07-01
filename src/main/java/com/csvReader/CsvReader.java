package com.csvReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;
import java.util.Vector;
import java.util.concurrent.ThreadLocalRandom;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.github.javafaker.Faker;



public class CsvReader 

{

	private static LocalDate integerVector;

	public static void main(String[] args) throws InterruptedException
	{

		
    	System.setProperty("webdriver.chrome.driver", "C:\\Users\\pooja\\eclipse-workspace\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = (WebDriver) new ChromeDriver();
		WebDriverWait wait = new WebDriverWait(driver, 10);

		driver.get("https://promoter.applination.in");
		driver.manage().window().maximize();
		
		driver.findElement(By.xpath("//input[@name='email']")).sendKeys("ashishgargvdoit@gmail.com");
		driver.findElement(By.xpath("//input[@name='email']")).sendKeys(Keys.TAB);

		Thread.sleep(3000);
		//driver.findElement(By.xpath("//input[@name='password'")).click();
		driver.findElement(By.xpath("//input[@name='password']")).sendKeys("123456");
		driver.findElement(By.cssSelector(".LoginButton")).click();
		Thread.sleep(2000);
	    driver.findElement(By.xpath("//div[contains(text(),'+ New Tournament')]")).click();
        
//	    Faker faker = new Faker();
		
		try 
		{
			
			FileInputStream file= new FileInputStream(new File( "C:\\Excelcode\\tournament.xlsx"));
			@SuppressWarnings("resource")
			XSSFWorkbook workbook= new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
//			XSSFSheet sheet2 = workbook.getSheetAt(1);
//			XSSFRow row;
//			row = sheet2.createRow(0);
//			Cell cell =	row.createCell(0);
//			cell.setCellValue("Pooja");
//			
//			FileOutputStream out = new FileOutputStream(
//            new File("C:\\Excelcode\\tournament2.xlsx"));
//			workbook.write(out);
			
//		    int rowcount=sheet.getLastRowNum();  //return the row count
//			System.out.println(rowcount);
		    int colcount=sheet.getRow(1).getLastCellNum();   //returns column/cell count
//		    System.out.println(colcount);
		    
		    String[] data = new String[colcount];
		    for ( int i = 0; i < colcount; i++)
		    {
		    	data[i] = sheet.getRow(1).getCell(i).toString();
//		    	System.out.println(data[i]);
//			    driver.findElement(By.cssSelector(".event-input")).sendKeys();
//			    driver.findElement(By.id("exampleFormControlTextarea1")).click();
//			    //description//
//			    driver.findElement(By.id("exampleFormControlTextarea1")).sendKeys("Event");
	    	}
//		    System.out.println(data[0]);
		  //  boolean skipFlag = true;
		    JavascriptExecutor js = (JavascriptExecutor) driver;
		  //  if(skipFlag) {
		 // event name//
			driver.findElement(By.cssSelector(".event-input")).sendKeys(data[0]);
			driver.findElement(By.id("exampleFormControlTextarea1")).click();
			// description//
			driver.findElement(By.id("exampleFormControlTextarea1")).sendKeys(data[1]);
			Thread.sleep(1000);
			js.executeScript("window.scrollBy(0,200)", "");
			// js.executeScript("arguments[0].scrollIntoView(true)",driver.findElement(By.xpath("//body/div[@id='root']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[3]/a[1]/div[1]/img[1]")));
			driver.findElement(By.xpath(
					"//body/div[@id='root']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[3]/a[1]/div[1]/img[1]"))
					.click();
			Thread.sleep(2000);
			// Address //
			int dropdownIdentifier = ThreadLocalRandom.current().nextInt(2, 7);
			driver.findElement(By.xpath(
					"//body/div[@id='root']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[3]/a[1]/div[1]/span[1]/div["
							+ Integer.toString(dropdownIdentifier) + "]/div[2]"))
					.click();
//		    driver.findElement(By.xpath("//body/div[@id='root']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[3]/a[1]/div[1]/span[1]/div[2]/div[2]")).click();
			Thread.sleep(2000);
			js.executeScript("window.scrollBy(0,200)", "");
			
			// Date code----
			// SimpleDateFormat formatter= new SimpleDateFormat("MM/dd/yyyy");
			// Date date = new Date(System.currentTimeMillis());
			LocalDate date1 = LocalDate.now().plusDays(5);
			LocalDate date2 = LocalDate.now().plusDays(7);

			// System.out.println(date1.format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
			// System.out.println(formatter.format(date));
			driver.findElement(By.xpath(
					"//body/div[@id='root']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[4]/div[2]/div[1]/div[3]/div[1]/div[1]/input[1]"))
					.sendKeys(date1.format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
			// end date//
			driver.findElement(By.xpath(
					"//body/div[@id='root']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[4]/div[3]/div[1]/div[3]/div[1]/div[1]/input[1]"))
					.sendKeys(date2.format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
			
			Faker faker = new Faker();
			// Time----
			String randomTime = faker.date().toString();
			driver.findElement(By.xpath(
					"//body/div[@id='root']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[5]/div[2]/div[1]/div[3]/div[1]/div[1]/input[1]"))
					.sendKeys("11:00 A");
			Thread.sleep(1000);
			driver.findElement(By.xpath("//body/div[4]/div[1]/div[1]/div[1]/div[1]/div[2]/ul[1]/li[2]/button[1]")).click();

			// Registration---
			driver.findElement(By.xpath(
					"//body/div[@id='root']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[6]/div[2]/div[1]/div[1]/div[3]/div[2]/a[1]/div[1]/img[1]"))
					.click();
			Thread.sleep(500);
			driver.findElement(By.xpath("//li[contains(text(),'Yes')]")).click();
			driver.findElement(By.xpath(
					"//body/div[@id='root']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[6]/div[3]/div[1]/div[1]/div[4]/a[1]/div[1]/img[1]"))
					.click();
			Thread.sleep(1000);
			driver.findElement(By.xpath("//li[contains(text(),'silver')]")).click();
			Thread.sleep(1000);
			// close on date//
			LocalDate date3 = LocalDate.now().plusDays(4);
			LocalDate date4 = LocalDate.now().plusDays(2);
			driver.findElement(By.xpath(
					"//body/div[@id='root']/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[6]/div[4]/div[1]/div[3]/div[1]/div[1]/input[1]"))
					.sendKeys(date3.format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
			driver.findElement(
					By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[6]/div[5]/div/div[3]/div/div/input"))
					.sendKeys(date4.format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
			// End At//
			driver.findElement(
					By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[6]/div[6]/div/div[3]/div/div/input"))
					.sendKeys("08:00 P");
			Thread.sleep(1000);
			driver.findElement(By.xpath("/html/body/div[7]/div/div/div/div/div[2]/ul/li[2]/button")).click();
		
			
			//template-----
	         driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[7]/div[1]/img")).click();
//	         driver.findElement(By.xpath("//div[contains(text(),'Template')]")).click();
             Thread.sleep(1000);
			js.executeAsyncScript("let confirmVal = '';function myFunction() {\r\n"
					+ "  let text = \"Do you want to create a new template? Click 'Cancel' to use an existing one.\";\r\n"
					+ "  if (confirm(text) == true) {\r\n"
					+ "    text = \"You pressed OK!\";\r\n"
					+ "    confirmVal = document.createElement('div'); confirmVal.id = 'hiddenEl'; confirmVal.setAttribute('val','true'); document.body.appendChild(confirmVal)\r\n"
					+ "    console.log(confirmVal);\r\n"					
					+ "    return confirmVal;\r\n"					
					+ "  } else {\r\n"
					+ "    text = \"You canceled!\";\r\n"
					+ "  }\r\n"
					+ "};"
					+ "myFunction();");
			Thread.sleep(10000);
			String confirmVal = js.executeScript("return document.getElementById('hiddenEl').getAttribute('val')").toString();
			
			if(confirmVal.equals("true"))
			{
				System.out.println(confirmVal);
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]/span")).click();
											
				Thread.sleep(300);
				
				// New button//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();
				Thread.sleep(2000);
				// Age bracket dropdown//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/img"))
						.click();
				// adult//

				int AgeBracketDropdown = ThreadLocalRandom.current().nextInt(4, 7);
//			  int AgeBracketDropdown = 4;

				Thread.sleep(1000);
//			  driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/span/ul/li[1]")).click();
				driver.findElement(
						By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/span/ul/li["
								+ Integer.toString(AgeBracketDropdown) + "]"))
						.click();
//			  driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/span/ul/li["+Integer.toString(2)+"]")).click();
				Thread.sleep(1000);
				// Next//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();
				if (AgeBracketDropdown == 1) {
					Thread.sleep(2000);
					int var = 2;
					int GenderGroup = ThreadLocalRandom.current().nextInt(1, var);
					Thread.sleep(2000);
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/label["
							+ Integer.toString(GenderGroup) + "]/input")).click();
//				  Next button
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					// Next Screen
					int i = ThreadLocalRandom.current().nextInt(1, 3);
					int x = 3;
					if (i == 1) {
						x = 4;
					} else {
						x = 3;
					}
					int j = ThreadLocalRandom.current().nextInt(1, x);
					// Skills level for division
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div["
							+ Integer.toString(i) + "]/label[" + Integer.toString(j) + "]/input")).click();
					Thread.sleep(1000);
					// Next Button
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]/span"))
							.click();
					Thread.sleep(500);
					i = ThreadLocalRandom.current().nextInt(1, 3);
					j = ThreadLocalRandom.current().nextInt(1, 3);
					// Skill Level for Division
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div["
							+ Integer.toString(i) + "]/label[" + Integer.toString(j) + "]/input")).click();
					Thread.sleep(1000);
					// next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/img")).click();
					// driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();
					// Total number of team size image 8//
					// size//
					Thread.sleep(300);
					driver.findElement(By
							.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/span/ul/li[5]"))
							.click();
					// next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					// player per team//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/img")).click();
					i = ThreadLocalRandom.current().nextInt(1, 4);
					Thread.sleep(1000);
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/span/ul/li["
									+ Integer.toString(i) + "]"))
							.click();
					// next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					// discount section//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/h1/div[1]/label/input"))
							.click();
					// Next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[3]/button[2]"))
							.click();
					// Again next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					// Enter Early bird date//
					LocalDate date5 = LocalDate.now().plusDays(2);

//					driver.findElement(
//							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/div/div/input"))
//							.clear();
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/div/div/input"))
							.sendKeys(date5.format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
					// Next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					// Early bird price//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div[1]/div/input"))
							.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
					// Next //
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					// Normal entry//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/input"))
							.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
					// Next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					// Late entry//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/input"))
							.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
					// Next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					String templateName = faker.lorem().word().toString();
					// Template name//

					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/label[1]/input"))
							.sendKeys(templateName);
					Thread.sleep(3000);
					// next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/button"))
							.click();

				}

				else if (AgeBracketDropdown == 2) {
					// Junior Section
					Thread.sleep(2000);
					// Boys
					int GenderGroup = ThreadLocalRandom.current().nextInt(1, 3);
					Thread.sleep(2000);
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div["
							+ Integer.toString(GenderGroup) + "]/label/input")).click();
					Thread.sleep(2000);
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					int i = ThreadLocalRandom.current().nextInt(1, 3);
					int j = ThreadLocalRandom.current().nextInt(1, 2);
					// Age Range for Division for Junior
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div["
							+ Integer.toString(i) + "]/label[" + Integer.toString(j) + "]/input")).click();
					Thread.sleep(1000);
					// Next
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]/span"))
							.click();

					// skill level devision//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/label["
							+ Integer.toString(i) + "]/input")).click();
					// /html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[2]/label[1]/input
					Thread.sleep(500);
//				 //Next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					Thread.sleep(50);
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					Thread.sleep(2000);
//		     	 //total img icon//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/img")).click();
					Thread.sleep(2000);
					// Total_Number of team//
					// Only 8 teams//
					js.executeScript("window.scrollBy(0,300)", "");

					driver.findElement(By
							.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/span/ul/li[5]"))
							.click();
					// Next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					// player per team//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/img")).click();
					i = ThreadLocalRandom.current().nextInt(1, 4);
					Thread.sleep(1000);
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/span/ul/li["
									+ Integer.toString(i) + "]"))
							.click();
					// next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					// discount section//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/h1/div[1]/label/input"))
							.click();
					// Next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[3]/button[2]"))
							.click();
					// Again next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					Thread.sleep(1000);

					// Enter Early bird date//
					LocalDate date6 = LocalDate.now().plusDays(2);
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/div/div/input"))
							.sendKeys(date6.format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
					// Next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					// Early bird price//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div[1]/div/input"))
							.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
					// Next //
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					// Normal entry//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/input"))
							.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
					// Next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					// Late entry//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/input"))
							.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
					// Next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					String templateName = faker.lorem().word().toString();
					// Template name//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/label[1]/input"))
							.sendKeys(templateName);
					Thread.sleep(3000);
					// next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/button"))
							.click();

				}
				// 3 step dinosour//
				else if (AgeBracketDropdown == 3) {
					//// Boys//
					Thread.sleep(2000);
					int var = 3;
					int GenderGroup = ThreadLocalRandom.current().nextInt(1, var);
					var = 2;
					GenderGroup = ThreadLocalRandom.current().nextInt(1, var);

					Thread.sleep(2000);
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/label["
							+ Integer.toString(GenderGroup) + "]/input")).click();
					// Next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					// Next Screen
					int i = ThreadLocalRandom.current().nextInt(1, 2);
					int x = 2;
					if (i == 1) {
						x = 4;
					} else {
						x = 2;
					}
					int j = ThreadLocalRandom.current().nextInt(1, x);

					// Age range//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div["
							+ Integer.toString(i) + "]/label[" + Integer.toString(j) + "]/input")).click();
					Thread.sleep(500);
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					Thread.sleep(2000);
//		     	 //total img icon//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/img")).click();
					Thread.sleep(2000);
					// Total_Number of team//
					// Only 8 teams//
					js.executeScript("window.scrollBy(0,300)", "");

					driver.findElement(By
							.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/span/ul/li[5]"))
							.click();
					// Next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					// discount section//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/h1/div[1]/label/input"))
							.click();

					Thread.sleep(1000);

					// Next button//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[3]/button[2]"))
							.click();
					Thread.sleep(1000);

					// Again next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
//			    Thread.sleep(100);
//			     driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();

					Thread.sleep(1000);

					// Enter Early bird date//
					LocalDate date7 = LocalDate.now().plusDays(2);
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/div/div/input"))
							.sendKeys(date7.format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
					// Next button//
					Thread.sleep(1000);

					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					// Early bird price//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div[1]/div/input"))
							.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
					// Next //
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					Thread.sleep(1000);

					// Normal entry//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/input"))
							.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
					// Next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();
					Thread.sleep(1000);

					// Late entry//
					driver.findElement(
							By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/input"))
							.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
					// Next//
					Thread.sleep(1000);
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]"))
							.click();

					String templateName = faker.lorem().word().toString();
					// Template name//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/label[1]/input"))
							.sendKeys(templateName);
					Thread.sleep(3000);
					// next//
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/button"))
							.click();

					// /html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[2]/div/div/div[4]/a/div/img
				}

				// 4 step relation father//

				else if (AgeBracketDropdown >= 4 || AgeBracketDropdown <= 7) {
					// Father-Daughter//
//				 Thread.sleep(2000);
//				int GenderGroup =ThreadLocalRandom.current().nextInt(4, 7);
//				 Thread.sleep(2000);
//				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/label[1]/input["+Integer.toString(GenderGroup)+"]")).click();
					Thread.sleep(2000);
					int i = ThreadLocalRandom.current().nextInt(1, 3);
					int j = ThreadLocalRandom.current().nextInt(1, 3);
					driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div["
							+ Integer.toString(i) + "]/label[" + Integer.toString(j) + "]/input")).click();
				}
				//Thread.sleep(20000);
				// next button//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();
//		 	 //total img icon//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/img"))
						.click();
				Thread.sleep(200);
				// Total_Number of team//
				// Only 8 teams//
				js.executeScript("window.scrollBy(0,50)", "");

				driver.findElement(
						By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div/a/div/span/ul/li[5]"))
						.click();
				Thread.sleep(200);

//				  Next button//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();

				// discount section//
				driver.findElement(
						By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/h1/div[1]/label/input"))
						.click();

				Thread.sleep(1000);

				// Next button//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[3]/button[2]")).click();
				Thread.sleep(1000);

				// Again next//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();
//		    Thread.sleep(100);
//		     driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();

				Thread.sleep(1000);

				// Enter Early bird date//
				LocalDate date7 = LocalDate.now().plusDays(2);
				driver.findElement(
						By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/div/div/input"))
						.sendKeys(date7.format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
				// Next button//
				Thread.sleep(1000);

				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();

				// Early bird price//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div[1]/div/input"))
						.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
		// Next //
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();
				Thread.sleep(1000);

				// Normal entry//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/input"))
						.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
		//Next//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();
				Thread.sleep(1000);

				// Late entry//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/div[1]/input"))
						.sendKeys(Integer.toString(ThreadLocalRandom.current().nextInt(0, 11)));
				// Next//
				Thread.sleep(1000);
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[2]/button[2]")).click();

				String templateName = faker.lorem().word().toString();
				// Template name//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/label[1]/input"))
						.sendKeys(templateName);
				Thread.sleep(3000);
				// next//
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/button")).click();

				Thread.sleep(1000);

				// template-----
//			   driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[7]/div[1]/img")).click();
//			     driver.findElement(By.xpath("//div[contains(text(),'Template')]")).click();
//			     Thread.sleep(1000);
//			   driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/a/div/img")).click();
				// driver.findElement(By.xpath("//body/div[3]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/div[1]/div[1]/a[1]")).click();
				// driver.findElement(By.xpath("//li[contains(text(),'Test_applination')]")).click();
				// Thread.sleep(1000);
				// git test//

//				driver.findElement(
//						By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[2]/div/div/div[4]/a/div/img"))
//						.click();
			}
			
			else
			{
				driver.findElement(By.xpath("/html/body/div[8]/div/div[2]/div/div/div/div/div/form/div[1]/div/a/div/img")).click();
											
				Thread.sleep(300);
			}
			
			//Pool code//
			Thread.sleep(10000);
			WebElement button = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[2]/div/div/div[4]/a/div/img")));
			button.click();
			Thread.sleep(1000);
			//driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[2]/div/div/div[4]/a/div/img")).click();
            //Value//
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[2]/div/div/div[4]/a/div/span/ul/li[9]")).click();
			
			
			//Season code//
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[3]/div/div/div[4]/a/div/img")).click();

			
			int SeasonDropdown = ThreadLocalRandom.current().nextInt(1, 5);
//			  int SeasonDropdown = 4;

				Thread.sleep(1000);
				driver.findElement(
						By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[3]/div/div/div[4]/a/div/span/ul/li["
								+ Integer.toString(SeasonDropdown) + "]"))
						.click();

				//Placements Point excel//
			//data[23]
			int colIndex = 23;
			Thread.sleep(1000);
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[4]/div/div/div[4]/a/div/img")).click();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//li[text()='"+data[colIndex++]+"']")).click();
			Thread.sleep(1000);
			//seeding image//
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[5]/div/div/div[4]/a/div/img")).click();
			driver.findElement(By.xpath("//li[text()='"+data[colIndex++]+"']")).click();
			Thread.sleep(200);
			//Surface Type//
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[6]/div/div/div[4]/a/div/img")).click();
			Thread.sleep(200);
			driver.findElement(By.xpath("//li[text()='"+data[25]+"']")).click();
			Thread.sleep(1000);
			//Host clinic//
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[7]/div/div[4]/a/div/img")).click();
			Thread.sleep(200);
			driver.findElement(By.xpath("//li[text()='"+data[26]+"']")).click();
			//Show entries//
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[8]/div/div[4]/a/div/img")).click();
			driver.findElement(By.xpath("//li[text()='"+data[27]+"']")).click();
			//scroll//
			js.executeScript("window.scrollBy(0,200)", "");
			// Team Listing//
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[9]/div/div/div[4]/a/div/img")).click();
			Thread.sleep(200);
			driver.findElement(By.xpath("//li[text()='"+data[28]+"']")).click();
			Thread.sleep(300);
			
			//Number of sets//
			int collIndexP =15;
			//WebDriverWait wait1 = new WebDriverWait(driver, 10);
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[10]/div/div/div[3]/div[2]/a/div/img")).click();
			Thread.sleep(800);
			Vector<String> integerVector = new Vector<String>();
			integerVector.add(data[15]);
			integerVector.add(data[16]);
			integerVector.add(data[17]);
			integerVector.add(data[18]);
			integerVector.add(data[19]);
			integerVector.add(data[20]);
			integerVector.add(data[33]);
			integerVector.add(data[34]);
			integerVector.add(data[35]);
			integerVector.add(data[36]);
			integerVector.add(data[37]);
			integerVector.add(data[38]);

			
			for(int i=0; i<integerVector.size(); i++)
			{
				String s = integerVector.get(i);
				s = s.substring(0, s.length() - 2);
				integerVector.set(i, s);
				
			}
			System.out.println(integerVector);

			driver.findElement(By.xpath("//li[text()='"+integerVector.get(0)+"']")).click();
			// Match Time//
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[11]/div/div/div[3]/div[2]/a/div/img")).click();

			Thread.sleep(1000);
			driver.findElement(By.xpath("//li[contains(.,'"+integerVector.get(1)+"')]")).click();
			//Max points per set//
			Thread.sleep(1000);
//			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[12]/div/div[3]/input")).click();
			Thread.sleep(500);
//			driver.findElement(By.xpath("//li[text()='"+integerVector.get(2)+"']")).click();
			driver.findElement(By.xpath("//*[@id=\"app-content\"]/div[2]/div[1]/div/div/div[8]/div[12]/div/div[3]/input")).clear();
			driver.findElement(By.xpath("//*[@id=\"app-content\"]/div[2]/div[1]/div/div/div[8]/div[12]/div/div[3]/input")).sendKeys(integerVector.get(2));
//			driver.findElement(By.xpath("//*[@id=\"app-content\"]/div[2]/div[1]/div/div/div[8]/div[13]/div/div/div[3]/div[1]")).sendKeys(integerVector.get(3));
										//*[@id="app-content"]/div[2]/div[1]/div/div/div[8]/div[12]/div/div[3]/input
			//Number of playoff//
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[13]/div/div/div[3]/div[2]/a/div/img")).click();
			Thread.sleep(500);
			driver.findElement(By.xpath("//li[text()='"+integerVector.get(3)+"']")).click();
			
			//Team in PlayOff 1//
			Thread.sleep(1000);
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[14]/div/div[3]/input")).sendKeys(integerVector.get(3));
//			driver.findElement(By.xpath("//li[text()='"+data[collIndexP++]+"']")).click();
			//Team Playoff 2//
			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[8]/div[15]/div/div[3]/input")).sendKeys(integerVector.get(4));
//			driver.findElement(By.xpath("//li[text()='"+data[collIndexP++]+"']")).click();
			
			js.executeScript("window.scrollBy(0,300)", "");
			Thread.sleep(1000);
		    
			//Contact//
			

			driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div[2]/div[1]/div/div/div[9]/div[2]/div/div[4]/a/div/img")).click();
			Thread.sleep(1000);
			driver.findElement(By.xpath("//div[text()='"+data[29]+"']")).click();
             //Main method//
			driver.findElement(By.xpath("/html/body/div/div/div[2]/div[2]/div[1]/div/div/div[9]/div[3]/div/div[4]/a/div/img")).click();
			Thread.sleep(600);
			driver.findElement(By.xpath("//div[text()='"+data[30]+"']")).click();
           //Finance//
			driver.findElement(By.xpath("/html/body/div/div/div[2]/div[2]/div[1]/div/div/div[10]/div[2]/div/div[3]/div[2]/a/div/img")).click();
			Thread.sleep(600);
			driver.findElement(By.xpath("//li[text()='"+data[31]+"']")).click();
           //Purse amount//
			driver.findElement(By.xpath("/html/body/div/div/div[2]/div[2]/div[1]/div/div/div[10]/div[3]/div/div[3]/div[2]/a/div/img")).click();
			Thread.sleep(600);
			driver.findElement(By.xpath("//li[text()='"+data[32]+"']")).click();
			Thread.sleep(600);
            //Amount percent//
			driver.findElement(By.xpath("/html/body/div/div/div[2]/div[2]/div[1]/div/div/div[10]/div[4]/div/div[3]/input")).sendKeys(integerVector.get(6));
			//Text//
			js.executeScript("window.scrollBy(0,100)", "");
			Thread.sleep(1000);
//			Thread.sleep(1000);
//			driver.findElement(By.xpath("/html/body/div/div/div[2]/div[2]/div[1]/div/div/di[10]/div[5]/div/div[3]/input")).clear();
			driver.findElement(By.xpath("/html/body/div/div/div[2]/div[2]/div[1]/div/div/di[10]/div[5]/div/div[3]/input")).sendKeys(integerVector.get(7));
			
			///Donation 1//
			Thread.sleep(600);
			driver.findElement(By.xpath("/html/body/div/div/div[2]/div[2]/div[1]/div/div/div[10]/div[6]/div/div[3]/input")).sendKeys(integerVector.get(8));
			//Donation 2//
			Thread.sleep(600);
			driver.findElement(By.xpath("/html/body/div/div/div[2]/div[2]/div[1]/div/div/div[10]/div[7]/div/div[3]/input")).sendKeys(integerVector.get(9));
			//Donation 3//
			Thread.sleep(600);
		 driver.findElement(By.xpath("/html/body/div/div/div[2]/div[2]/div[1]/div/div/div[10]/div[8]/div/div[3]/input")).sendKeys(integerVector.get(10));
			
			
		   // } // SkipFlag
				Thread.sleep(1000000);
		}
		

		catch (Exception e)
		{
			e.printStackTrace();
			System.out.println();
		}
	}
}
