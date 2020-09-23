package totalQaApacheReadXLSX;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadXLSX {

	public static void main(String[] args) throws Exception {
		
		{
		File f= new File("TestCaseData.xlsx")	;
		
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook excelworkbook= new XSSFWorkbook(fis);
		XSSFSheet excelsheet =  excelworkbook.getSheetAt(0);
		
		XSSFCell cell;
		
		int rows = excelsheet.getPhysicalNumberOfRows();
		int cols = excelsheet.getRow(0).getPhysicalNumberOfCells();
		
		String data[][]= new String[rows][cols];
		
			for( int i =0;i<rows; i++) 
			{
				
				for( int j=0; j<cols;j++)
				{
					cell= excelsheet.getRow(i).getCell(j);
					String cellcontents= cell.getStringCellValue();
					data[i][j]= cellcontents;
					System.out.println(cellcontents);
					
				}
				
				
			}
			fis.close();
			}
		
			
			
			
			
			
}
	}