package Oct15;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteExcel 
{ 
	public static void main(String[] args)
	{
		try
		{
			XSSFWorkbook book = new XSSFWorkbook("D:/SeleniumW/Excel/src/Oct15/Data.xlsx");//object of work book
			XSSFSheet sheet = book.getSheet("Sheet1");//create object of sheet
			XSSFRow row2 = sheet.getRow(1);//create object of row which wants to read
			
			Cell cell1 = row2.getCell(0);
			Cell cell2 = row2.getCell(1);
	 		Cell cell3 = row2.getCell(2);
			
			System.out.println(cell1 + "" + cell2 + "" + cell3);//reads string from cell
			System.out.println(cell1.getStringCellValue() +"");//to fetch specific value, if its a date then call a method cell1.getDateCellValue()
			book.close();			
		} 
		catch (IOException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}


}
