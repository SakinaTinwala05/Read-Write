package excelReadAndWrite;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	public static void main (String[] args) {
		readDataFile("D:\\ExcelFiles\\ExcelReadWriteData.xlsx");
	}

	private static void readDataFile(String File) {
		
		try {
			XSSFWorkbook work = new XSSFWorkbook(new FileInputStream(File));
			
			XSSFSheet sheet = work.getSheet("Employee");
			
			XSSFRow row = null;
			
			int i=0;
			while ((row = sheet.getRow(i))!=null) {
				
				System.out.println("No. :"+row.getCell(0).getNumericCellValue());
				System.out.println("Name :"+row.getCell(1).getStringCellValue());
				i++;
			}
			
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	
	
}
