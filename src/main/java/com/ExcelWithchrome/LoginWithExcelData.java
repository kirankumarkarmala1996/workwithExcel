package com.ExcelWithchrome;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class LoginWithExcelData {
	public static void main(String[] args) throws InterruptedException {
		LoginWithExcelData led = new LoginWithExcelData();
//		System.out.println(led.getExcelData2("kiran", 0, 0));
//		System.out.println(led.getExcelData2("kiran", 0, 1));

		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.get("https://demo.actitime.com/login.do");
		driver.manage().window().maximize();
		driver.manage().timeouts().pageLoadTimeout(25, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);

		int lastrn = getLastRowNum("kiran");
		
		for (int i = 1; i <=lastrn; i++) {
			String Un = led.getExcelData2("kiran", i, 0);
			String Pa = led.getExcelData2("kiran", i, 1);
			System.out.println(Un);
			System.out.println(Pa);
			driver.findElement(By.id("username")).sendKeys(Un);
			driver.findElement(By.name("pwd")).sendKeys(Pa);
			driver.findElement(By.xpath("//div[text()='Login ']")).click();
			Thread.sleep(3000);
			driver.findElement(By.id("logoutLink")).click();
		}
	}

	public String getExcelData2(String sheetname, int rowno, int cellno) {
		String rtv = null;

		try {
			FileInputStream file = new FileInputStream("G:\\Exeldata\\actitime.xlsx");
			Workbook wb = WorkbookFactory.create(file);
			Sheet sh = wb.getSheet(sheetname);
			Row rr = sh.getRow(rowno);
			Cell ce = rr.getCell(cellno);
			rtv = ce.getStringCellValue();

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return rtv;
	}

	public static int getLastRowNum(String Sheet) {
		int lrn = 0;

		FileInputStream file;
		try {
			file = new FileInputStream("G:\\Exeldata\\actitime.xlsx");
			Workbook wb = WorkbookFactory.create(file);
			Sheet st = wb.getSheet(Sheet);
			lrn = st.getLastRowNum();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return lrn;
	}
}



//Exception in thread "main" org.openqa.selenium.StaleElementReferenceException: stale element reference:
//		element is not attached to the page document

