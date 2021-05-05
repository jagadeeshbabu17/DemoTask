package org.sample1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Employee2 {
	public static void main(String[] args) throws IOException {
		File file = new File("C:\\Users\\jb\\selinium\\Employee\\file\\jb2.xls");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new HSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Sheet1");
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				int type = cell.getCellType();
				if (type == 1) {
					String name = cell.getStringCellValue();
					System.out.println(name);
				}
				if (type == 0) {
					if (DateUtil.isCellDateFormatted(cell)) {
						// SimpleDateFormat f = new SimpleDateFormat("dd-MMM");
						// String date = f.format(cell.getDateCellValue());
						String date = new SimpleDateFormat("dd-MMM").format(cell.getDateCellValue());
						System.out.println(date);
					} else {
						// double d = cell.getNumericCellValue();
						// long l= (long) d;
						// String value = String.valueOf(l);
						String value = String.valueOf((long) cell.getNumericCellValue());
						System.out.println(value);
					}
				}
			}
		}
	}

}
