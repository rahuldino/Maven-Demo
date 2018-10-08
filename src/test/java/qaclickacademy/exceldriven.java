package qaclickacademy;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class exceldriven {
	// Identify testcases column by scanning the entire 1st row
	// once column is identified then scan entire testcase column to identify purchase testcase row
	// after you grab purchase testcase row pull all the data of that row and feed into test 

		// TODO Auto-generated method stub
		public  ArrayList<String> getData(String testcasename) throws IOException
		{
			ArrayList<String>a= new ArrayList<String>();
			FileInputStream fis=new FileInputStream("C:\\Users\\Rahul Philip-DELL\\Desktop\\data.xlsx"); // reading the file
			XSSFWorkbook workbook=new XSSFWorkbook(fis); // setting the object in the workbook
			int sheets=workbook.getNumberOfSheets(); // get the number of sheets present
			for(int i=0;i<sheets;i++) // looping to get that particular sheet
			{
				if(workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) // get this sheet
				{
				XSSFSheet sheet=workbook.getSheetAt(i); // getting the index
				Iterator<Row> rows=sheet.iterator(); // go to each and every row, sheet is collection of cells
				Row firstrow=rows.next(); // go to the next row
				Iterator<Cell> cell=firstrow.cellIterator(); // row is collection of cells
				int k=0; // declaring a variable for looping
				int column = 0;
				while(cell.hasNext()) // check if next cell is present
				{
					Cell value=cell.next();
					if(value.getStringCellValue().equalsIgnoreCase(testcasename))
					{
						column=k;
					}
					k++; // incrementing the variable

				}
				System.out.println(column);
				// once column is identified then scan entire testcase column to identify purchase testcase row
				while(rows.hasNext())
				{
					Row r=rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcasename))
					{
						// after you grab purchase testcase row pull all the data of that row and feed into test 
						Iterator<Cell> cv=r.cellIterator(); // iterate to each cell
						while(cv.hasNext())
						{
							Cell c =cv.next();
							if(c.getCellTypeEnum()==CellType.STRING)  // checks what cell type it is
							{
							a.add(c.getStringCellValue()); // get the string value
						}
							else
							{
								a.add(NumberToTextConverter.toText(c.getNumericCellValue())); // get numeric cell value
							}

							}
					}
				}
			}
			

		}
			return a;

		}
	
		
	}

