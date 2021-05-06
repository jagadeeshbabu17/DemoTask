package org.sample1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Employee {
	public static void main(String[] args) throws IOException {
		File file = new File("C:\\Users\\jb\\selinium\\Employee\\file\\jb.xls");
		 FileInputStream stream = new FileInputStream(file);
		 System.out.println("hai");
		 System.out.println("hai");
		 Workbook workbook = new HSSFWorkbook(stream);
		 Sheet sheet = workbook.getSheet("Sheet1");
//		 Row row = sheet.getRow(7);
//		 Cell cell = row.getCell(9);
//		 int count = sheet.getPhysicalNumberOfRows();
//		 System.out.println(count);
//		FileInputStream stream = new FileInputStream(file);
//
//		Workbook workbook = new XSSFWorkbook(stream);
//
//		Sheet sheet = workbook.getSheet("Sheet1");
//
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {

			Row row = sheet.getRow(i);

			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {

				Cell cell = row.getCell(j);

				System.out.println(cell);

			}

		}
	}
}
