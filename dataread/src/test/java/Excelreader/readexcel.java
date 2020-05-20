package Excelreader;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.impl.SheetTypeImpl;

public class readexcel 
{
		
	public FileInputStream fis = null;
	public XSSFWorkbook workbook = null;
	public XSSFSheet sheet = null;
	public XSSFRow row = null;
	public XSSFCell cell = null;
	String value="";
	
	public readexcel(String xlfilepath) throws IOException
	{
	
		FileInputStream fis = new FileInputStream(xlfilepath);
		workbook = new XSSFWorkbook(fis);
		fis.close();
	}
		

	public String getCelldata(String sheetname,int colNum,int rowNum)
	{
		
		sheet = workbook.getSheet(sheetname);
		row = sheet.getRow(rowNum);
		cell = row.getCell(colNum);

	
		
try {		
		if(cell.getCellType() == CellType.STRING)
		{
			value = cell.getStringCellValue();
		}
		else if(cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA)
		{
			value = String.valueOf(cell.getNumericCellValue());			
		}
		else if(cell.getCellType() == CellType.BOOLEAN)
		{
			value = String.valueOf(cell.getBooleanCellValue());
		}
		else if(cell.getCellType() == CellType.BLANK)
		{
			value = "";
		}
}
catch(Exception e)
{
	e.printStackTrace();
	return "No match found";
}
return value;

	}

	public static void main(String[] args) throws IOException
	{
	
		readexcel read = new readexcel("C:/Users/user/eclipse-workspace/dataread/dataExcel.xlsx");
		System.out.println(read.getCelldata("sheet1", 0, 1));
		System.out.println(read.getCelldata("sheet1", 1, 1));
		System.out.println(read.getCelldata("sheet1", 2, 1));
		
		System.out.println("------------");
		
		System.out.println(read.getCelldata("sheet1", 0, 2));
		System.out.println(read.getCelldata("sheet1", 1, 2));
		System.out.println(read.getCelldata("sheet1", 2, 2));
		
		System.out.println("------------");
		
		System.out.println(read.getCelldata("sheet1", 0, 3));
		System.out.println(read.getCelldata("sheet1", 1, 3));
		System.out.println(read.getCelldata("sheet1", 2, 3));
		
		System.out.println("------------");

		System.out.println(read.getCelldata("sheet1", 0, 4));
		System.out.println(read.getCelldata("sheet1", 1, 4));
		System.out.println(read.getCelldata("sheet1", 2, 4));
		
	}
}
