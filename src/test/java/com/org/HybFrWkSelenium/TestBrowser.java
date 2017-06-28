package com.org.HybFrWkSelenium;


import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class TestBrowser {
	
	@Test
	public void OpenBrowser() throws IOException {
		System.out.println("siteURL");
		System.setProperty("webdriver.gecko.driver", "C://Users//asyedzia//workspace//myHybridFramework//geckodriver.exe");
		WebDriver driver = new FirefoxDriver();
		//System.setProperty("webdriver.IE.driver", "C://Users//asyedzia//workspace//myHybridFramework//geckodriver.exe");
		//WebDriver driver = new InternetExplorerDriver();
		driver.get("https://assurancetech.wordpress.com");
		//driver.findElement(By.xpath("//a[@href='https://assurancetech.wordpress.com/2016/07/07/hello-world-for-selenium-a-beginners-guide-with-sample-code/']")).click();
		
		String excelFilePath = "C://Users//asyedzia//workspace//myHybridFramework//myDatanKeyWordFile.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
         
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firstSheet.iterator();
         
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
             
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                 
               
                System.out.print(cell.getStringCellValue());
                if (cell.getStringCellValue().equals("ClickLink")) {
                	
                	System.out.println("Next Cell to Get the Link");
	                cell = cellIterator.next();
	                
	                System.out.println(cell.getStringCellValue());
	                driver.findElement(By.xpath(cell.getStringCellValue())).click();  
	                driver.manage().timeouts().implicitlyWait(45, TimeUnit.SECONDS);
	                
	                System.out.println("Getting back");
	                driver.get("https://assurancetech.wordpress.com");
                }
	            break;                
	            
                
            }
            System.out.println("Next Row");            
        }
         
        workbook.close();
        inputStream.close();

		
	}

}
