package Excelreader;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelRead1 
{
	
	public void readExceldata() throws IOException
	{
		FileInputStream fis = new FileInputStream("C:/Users/user/eclipse-workspace/dataread/dataExcel.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("sheet1");
	XSSFRow	row = sheet.getRow(0);
	//XSSFCell cell = row.getCell(0);
	
	int rowCount = sheet.getPhysicalNumberOfRows();
	int colCount =row.getPhysicalNumberOfCells();
	
	System.out.println("row count ==>"+ rowCount);
	System.out.println("column count ==>"+ colCount);
	
	for(int i=1;i<rowCount;i++)
	{
		row = sheet.getRow(i);
		
		for(int j=0;j<colCount;j++)
		{
			System.out.println(row.getCell(j).getStringCellValue());
		}
		
		System.out.println("------------");
	}	
	}
	
	

	public static void main(String[] args) throws IOException 
	{
	
		excelRead1 read1 = new excelRead1();
		read1.readExceldata();

	}

}
