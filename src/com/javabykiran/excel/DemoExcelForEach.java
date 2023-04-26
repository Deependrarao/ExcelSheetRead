package com.javabykiran.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class DemoExcelForEach {

	WebDriver driver;

	public void driverSetting(String url) {
		System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get(url);
	}

	public void excelData() {
		try {
			FileInputStream fis = new FileInputStream("D:\\javabykiran-Selenium-Softwares\\excelFiles\\myexcel.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheet("Sheet1");

			for (Row row : sheet) {
				for (Cell cell : row) {
					switch (cell.getCellType()) {
					case STRING:
						String cellvalue = cell.getStringCellValue();
						System.out.print(cellvalue + "\t");
						break;

					case NUMERIC:
						double cellvaluenum = cell.getNumericCellValue();
						System.out.print(cellvaluenum + "\t");
						break;

					}
				}
				System.out.println();

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();

		}
	}

}
