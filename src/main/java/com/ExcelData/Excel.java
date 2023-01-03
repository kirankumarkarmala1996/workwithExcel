package com.ExcelData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel {
	public void getExcelData() {
		try {
			FileInputStream file = new FileInputStream("G:\\Exeldata\\kiran.xlsx");
			Workbook wb = WorkbookFactory.create(file);
			Sheet sh = wb.getSheet("raja");
			Row rr = sh.getRow(1);
			
			Cell ce = rr.getCell(5);
			String st = ce.getStringCellValue();
			System.out.println(st);

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
	}

	public void getExcelData1(String sheetname, int rowno, int cellno) {
		try {
			FileInputStream file = new FileInputStream("G:\\Exeldata\\kiran.xlsx");
			Workbook wb = WorkbookFactory.create(file);
			Sheet sh = wb.getSheet(sheetname);
			Row rr = sh.getRow(rowno);
			Cell ce = rr.getCell(cellno);
			String st = ce.getStringCellValue();
			System.out.println(st);

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
	}

	public String getExcelData2(String sheetname, int rowno, int cellno) {
		String rtv = null;

		try {
			FileInputStream file = new FileInputStream("G:\\Exeldata\\kiran.xlsx");
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

}
