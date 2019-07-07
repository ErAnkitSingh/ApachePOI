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

public class DataDriven {

	
	// Strategy (Hold XLsx file --> Access Sheets --> Access specific Sheets --> Access Rows --> Access specific row -->
	// Access Cells --> Access specific Cell)
	
	
	public ArrayList<String> getData(String testcasename) throws IOException
	{
		ArrayList<String> a=new ArrayList<String>();

		FileInputStream fis = new FileInputStream("C://Users//Acer//Desktop//Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);  //Hold xlsx file in object workbook

		int sheets = workbook.getNumberOfSheets(); // count number of sheets inside file 
		System.out.println("Num of Sheet " + sheets);
		for (int i=0;i<sheets;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("Demo")) //Access the sheet of name Demo
			{
				XSSFSheet sheet = workbook.getSheetAt(i); // XssfSheet hold that specific sheet

				Iterator<Row> rows = sheet.iterator(); // Hold the all rows
				Row firstrow = rows.next(); // start searching in first row 
				Iterator<Cell> ce = firstrow.cellIterator();  // Hold the cell of the first rows.
				
				int j = 0;
				int col = 0 ;
				
				while(ce.hasNext())  //loop is running till cell is present
				{
					Cell value= ce.next();   //Hold the cells
					if(value.getStringCellValue().equalsIgnoreCase("Testcases"))  //Access the Testcases rows value
					{
						col = j;
					}
					j++;
				}
				System.out.println("Print the col " + col); // Here We successfully identify the column index of Testcases i.e; 0

				while(rows.hasNext()) // Start searching in col one (TestCases)
				{
					Row r = rows.next();
					if(r.getCell(col).getStringCellValue().equalsIgnoreCase(testcasename)) // In col one identify testcase name
					{
						Iterator<Cell> cv=r.cellIterator(); //Hold the cells
						while(cv.hasNext())
						{
							//System.out.println(cv.next().getStringCellValue() + " Inside cell");
							Cell c=cv.next();  // Get all the value of cells inside object c
							System.out.println(c+ " cells");
							if(c.getCellTypeEnum()==CellType.STRING)
							{
								//send data to array list
								a.add(c.getStringCellValue());
							}
							else
							{
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));

							}
						}
					}
				}

			}
		}
		return a;

	}
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub


	}

}
