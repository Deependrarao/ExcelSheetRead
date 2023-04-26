package com.javabykiran.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class DemoExcel1 {

	WebDriver driver;

	public void driverSetting(String url)
	{
		System.setProperty("webdriver.chrome.driver","chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get(url);
	}

	public void excelData() 
	{
		try {
			FileInputStream fis = new FileInputStream("D:\\javabykiran-Selenium-Softwares\\excelFiles\\myexcel.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			int rowcount = sheet.getLastRowNum();
			System.out.println("Number of row "+rowcount);
			for(int i =0; i<=rowcount; i++) {
				XSSFRow row = sheet.getRow(i);

				int cellcount = row.getLastCellNum();
				System.out.println("Number of cell "+cellcount);
				for(int j=0; j<cellcount; j++) {
					if(j!=0 && i!=0) 
					{
						System.out.println("if condition is true");
						System.out.println("==>"+row.getCell(j).getNumericCellValue());
					}else {
						System.out.println("else condition");
						System.out.println("==>"+row.getCell(j).getStringCellValue());
					}
				}
			}
		} 
		catch (FileNotFoundException e) 
		{
			e.printStackTrace();
		}
		catch (Exception e)
		{
			e.printStackTrace();



		}
	}

}
