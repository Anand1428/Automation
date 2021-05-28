package com.Automation_Maven1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data_Driven {

	public static void particular_Data() throws IOException {
		File f = new File("C:\\Users\\POPPY\\eclipse-workspace\\Automation_Maven1\\user.xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fi);
		Sheet sheetAt = wb.getSheetAt(0);
		Row row = sheetAt.getRow(2);
		Cell cell = row.getCell(0);
		CellType cellType = cell.getCellType();
		if (cellType.equals(CellType.STRING)) {
			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);
		} else if (cellType.equals(CellType.NUMERIC)) {
			double numericCellValue = cell.getNumericCellValue();
			int value = (int) numericCellValue;
			System.out.println(value);

		}

	}

	public static void alldata() throws IOException {
		File f = new File("C:\\Users\\POPPY\\eclipse-workspace\\Automation_Maven1\\user.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fis);
		Sheet sheetAt = w.getSheetAt(0);
		int row_size = sheetAt.getPhysicalNumberOfRows();
		for (int i = 0; i < row_size; i++) {
			Row row = sheetAt.getRow(i);
			int cell_size = row.getPhysicalNumberOfCells();
			for (int j = 0; j < cell_size; j++) {
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				if (cellType.equals(CellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
				} else if (cellType.equals(CellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int value = (int) numericCellValue;
					System.out.println(value);

				}

			}
		}

	}
	 
	public void instruction() {

	}	
	

	public static void main(String[] args) throws IOException {
		particular_Data();
		System.out.println("####alldata####");
		alldata();
	}

}
