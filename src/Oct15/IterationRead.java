package Oct15;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import javax.imageio.IIOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class IterationRead 
{

	public static void main(String[] args)
	{
		IterationRead.readExcel();		
		IterationRead Ir = new IterationRead();
		Ir.writeExcel();
	}

	public static void readExcel()
	{
		try
		{
			XSSFWorkbook book = new XSSFWorkbook("D:/SeleniumW/Excel/src/Oct15/Data.xlsx");//object of work book
			XSSFSheet sheet = book.getSheet("Sheet1");//create object if sheet

			Iterator<Row> rite = sheet.rowIterator();// created row iterator for row

			while(rite.hasNext())//hasnext is a method returns true if there exist new row ahead
			{
				Row row = rite.next();
				Iterator<Cell> cite = row.cellIterator();// created cell iterator for cell

				while(cite.hasNext())
				{
					Cell cell = cite.next();
					System.out.println(cell.getStringCellValue());

				}
			}

		} 
		catch (IOException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void writeExcel()
	{
		try{
			File file = new File("D:/SeleniumW/Excel/src/Oct15/Data.xlsx");
			FileInputStream fis = new FileInputStream(file);

			XSSFWorkbook book = new XSSFWorkbook(fis);//object of work book
			XSSFSheet sheet = book.getSheet("Sheet1");//create object if sheet
			XSSFRow row = sheet.getRow(1);
			Cell cell = row.getCell(2);

			cell.setCellValue("Hello");
			fis.close();

			file = new File("D:/SeleniumW/Excel/src/Oct15/Data.xlsx");
			FileOutputStream fos = new FileOutputStream(file);
			book.write(fos);//write to excel
			book.close();

		}
		catch(IOException ex)
		{

		}

// this is a comment added to test the pull request



	}


}
