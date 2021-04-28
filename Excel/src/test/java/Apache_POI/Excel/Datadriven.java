package Apache_POI.Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Datadriven {
	
	public static void main(String args[]) throws IOException
	{
		FileInputStream fis = new FileInputStream("C:\\Users\\ganes\\eclipse-workspace\\Excel\\Testdata.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheets = workbook.getNumberOfSheets();
		
		for(int i=0; i<sheets; i++)
		{
			if(workbook.getSheetName(1).equalsIgnoreCase("testdata"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
			
				Iterator<Row> rows = sheet.iterator();
				
				Row firstrow = rows.next();
				
				Iterator<Cell> cell = firstrow.cellIterator();
				int k=0;
				int column=0;
				while(cell.hasNext())
				{
					Cell value  = cell.next();
					if(value.getStringCellValue().equalsIgnoreCase("Testcases"))
					{
						column =k;
					}
					k++;
				}
			
				while(rows.hasNext())
				{
					Row r = rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase"))
					{
						Iterator<Cell> cv = r.cellIterator();
						while(cv.hasNext())
						{
							System.out.println(cv.next().getStringCellValue());
						}
					}
				}
			}
		}
		
	
	}

}
