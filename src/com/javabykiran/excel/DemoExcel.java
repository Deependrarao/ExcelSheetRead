package com.javabykiran.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class DemoExcel 
{
	WebDriver driver;

	public void driverSetting(String Url) 
	{
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get(Url);
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	}

	public void readExcel() 
	{
		try {
			FileInputStream fis = new FileInputStream("D:\\javabykiran-Selenium-Softwares\\excelFiles\\myexcel.xlsx");
           XSSFWorkbook workbook = new XSSFWorkbook(fis);
           XSSFSheet sheet = workbook.getSheetAt(0);
          // XSSFRow row = sheet.getRow(0);
           XSSFRow row = sheet.getRow(1);
           XSSFCell cell = row.getCell(0);
           System.out.println("Data Usename:=>"+cell.getStringCellValue());
           System.out.println("Data Password:=>"+row.getCell(1).getNumericCellValue());
           

		} catch (FileNotFoundException e)
		{ 
			e.printStackTrace();
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}
}
