package test.java;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelPlugin {
	
	
	
	public XSSFWorkbook workbook;
	public ExcelPlugin() throws IOException
	{
		
		System.out.println("Loading Excel...");
		File f = new File("C:/Users/Dinesh/Desktop/TestData/TestData.xlsx");
		FileInputStream fis = new FileInputStream(f);
		workbook = new XSSFWorkbook(fis);
		System.out.println("Loading Excel Completed...");
		
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		int type = sheet.getRow(1).getCell(1).getCellType();
		switch(type)
		{
		case 0:
			System.out.println(sheet.getRow(1).getCell(0).getNumericCellValue());
			break;
		case 1:
			System.out.println(sheet.getRow(1).getCell(0).getStringCellValue());
			break;
			
		default:
			break;
		}
				
		}
	
	public int getRowCount(String sheet) 
	{
	int	getRowCount=0;
	getRowCount =workbook.getSheet("Sheet1").getLastRowNum()+1;
	return getRowCount;
	}
	
	public int getColumnCount(String sheet)
	{
		int getColumnCount=0;
		
		getColumnCount=workbook.getSheet("Sheet1").getRow(0).getLastCellNum();
		
		return getColumnCount;
	}

}
