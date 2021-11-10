package org.login;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Browser {
	
	public static void main(String[] args) throws IOException {
		
		File f = new File("E:\\Suthakar Study Material\\FrameWork.xlsx");
		
		FileInputStream stream = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(stream);
		
		Sheet sheet = w.getSheet("Sheet1");
		
		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		
		for (int i = 0; i < physicalNumberOfRows ; i++) {
			
			Row row = sheet.getRow(i);
			
			int physicalNumberOfCells = row.getPhysicalNumberOfCells();
			
			for (int j = 0; j < physicalNumberOfCells ; j++) {
				
				Cell cell = row.getCell(physicalNumberOfCells);
				
				int cellType = cell.getCellType();
				
			if (cellType==1) {
				String stringCellValue = cell.getStringCellValue();
				System.out.println(stringCellValue);
			} 
			
			else if (DateUtil.isCellDateFormatted(cell)) {
				Date dateCellValue = cell.getDateCellValue();
				System.out.println(dateCellValue);
				
			}
			else {
				double numericCellValue = cell.getNumericCellValue();
				
				long l = (long)numericCellValue;
				System.out.println(l);
			}
	
		
			}

		}
		
		
	}

}
