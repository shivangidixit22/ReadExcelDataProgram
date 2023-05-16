package assignment5;

import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadDataFromExcel {
	@Test
	public void readData() throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook("C:\\Users\\SH360592\\Programs\\Jenkins5\\TestData.xlsx");
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int rowCount = sheet.getPhysicalNumberOfRows();
		int colCount = sheet.getRow(0).getPhysicalNumberOfCells();
		
		for(int i =0; i <rowCount; i++) {
			for(int j=0; j<colCount; j++) {
				if(sheet.getRow(i).getCell(j).getCellType().toString().equals("STRING")) {
					String s = sheet.getRow(i).getCell(j).getStringCellValue();
					System.out.print(s+"   |   ");
				}else {
					double a = sheet.getRow(i).getCell(j).getNumericCellValue();
					int x = (int)a;
					System.out.print(x+"   |   ");
				}
				
			}
			System.out.println();
		}
 
	}	
}
