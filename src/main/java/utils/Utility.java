package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.io.FileHandler;

public class Utility {
	// to take screenshot of failed test cases
		public static void captureScreenShot(WebDriver driver,String testCaseID) throws IOException {
			String filePath ="test-output\\FailedScreenShot";
			TakesScreenshot ts = (TakesScreenshot) driver;
			File src = ts.getScreenshotAs(OutputType.FILE);    //SS saved to unknown location
			
			LocalDateTime currentTime =LocalDateTime.now();
			DateTimeFormatter formatter =DateTimeFormatter.ofPattern("  yyyy-MM-dd HH.mm.ss.SSS");
			String formattedDateTime = currentTime.format(formatter);
			
			File dest = new File(filePath+"\\"+testCaseID+formattedDateTime+".jpeg");
			FileHandler.copy(src, dest);
			
		}
		

		static String filePath ="C:\\Users\\adity\\eclipse-workspace\\AjioAutomationMaven\\src\\main\\resources\\testCases\\AjioTestCases.xlsx";
		
		// another way to get cell data
		
			public static String getExcelData ( String sheetName, int rowIndex, int cellIndex) throws EncryptedDocumentException, FileNotFoundException, IOException {
			
			FileInputStream file = new FileInputStream(filePath);
			Workbook book = WorkbookFactory.create(file);
			Sheet sheet = book.getSheet(sheetName);
			Row row = sheet.getRow(rowIndex);
			Cell cell = row.getCell(cellIndex);  
			
			String data;
			DataFormatter formatter = new DataFormatter();    
			try {
				data = formatter.formatCellValue(cell);  //this method will get any type of cell data and convert it into string 
			}
			catch(Exception e) {
				data="";
			}
			return data;
			
		}
			
			
//		//to get data from excel file
//		public static String excelData (String sheetName, int rowIndex,int cellIndex) throws EncryptedDocumentException, IOException {
//			Workbook book =WorkbookFactory.create(new FileInputStream(filePath));
//			try {
//			String cellData = book.getSheet(sheetName).getRow(rowIndex).getCell(cellIndex).getStringCellValue();
//			return cellData;
//			}
//			catch(IllegalStateException e){
//				
//				double cellData = book.getSheet(sheetName).getRow(rowIndex).getCell(cellIndex).getNumericCellValue();//if cell has DATE type 
//				String cellData1 = cellData+"";          //conversion from double to string   ANOTHER APPROACH =>  Double.toString(d);      OR  =>String s1 = String.valueOf(d);
//				return cellData1;
//				
//			}
//		}
			
			
}
