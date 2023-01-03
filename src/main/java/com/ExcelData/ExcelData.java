package com.ExcelData;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class ExcelData {
	public static void main(String[] args) {
		Excel ec = new Excel();
//		ec.getExcelData();
//		ec.getExcelData1("raja", 1, 1);
//		ec.getExcelData1("raja", 1, 2);
		
		
		String ss=ec.getExcelData2("raja", 1, 2);
		System.out.println(ss);
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.get("https://demo.actitime.com/login.do");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
		
		driver.findElement(By.id("username")).sendKeys("admin");
		driver.findElement(By.name("pwd")).sendKeys("manager");
		driver.findElement(By.xpath("//div[text()='Login ']")).click();
	}

}
