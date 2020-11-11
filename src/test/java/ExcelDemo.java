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

public class ExcelDemo {

	public ArrayList<String> getData(String testCaseName) throws IOException {
		
		ArrayList<String> al = new ArrayList<String>();		
		FileInputStream fis = new FileInputStream("D:\\Vinaya_REST\\Rest_demo.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheets = workbook.getNumberOfSheets();
		
		//get value of Testdate column and Purchase row data values
		for(int i=0;i<sheets;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
				Row firstRow = rows.next();
				Iterator<Cell> ce = firstRow.cellIterator();
				int k=0;
				int column=0;
				while(ce.hasNext()) {
					Cell value = ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						column=k;
						
					}
					k++;
				}
				
				//iterate the row which has "purchase"
				while(rows.hasNext()) {
					
					Row r = rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase")) {
						Iterator<Cell> cv = r.cellIterator();
						while(cv.hasNext()) {
							//System.out.println(cv.next().getStringCellValue());
							Cell c = cv.next();
							//If excel has both int and string type
							if(c.getCellTypeEnum()==CellType.STRING)
							{							
							al.add(c.getStringCellValue());
							}
							else {
								//Poi has NumberToTextConverter to convert int/double to string. 
								//Arraylist created here is of String type here, so make all values to string
								al.add(NumberToTextConverter.toText(c.getNumericCellValue()));
								
							}
						}
						
					}
					
					
				}
				
				
				}
			
			
			
		}
		
		return al;
	}
	
	public static void main(String[] args) throws IOException {
		
		
		
	}
}
