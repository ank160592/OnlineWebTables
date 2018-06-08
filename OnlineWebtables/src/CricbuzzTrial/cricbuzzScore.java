package CricbuzzTrial;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.NoSuchElementException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class cricbuzzScore {
public void writeToExcel(String filepath, String filename, String sheetname, String[] valueToWrite, int colCount, int excelRowCount) throws IOException {
		
		File file= new File(filepath+"\\"+filename);
		FileInputStream inputStream = new FileInputStream(file);
		String fileNameExtension = filename.substring(filename.indexOf("."));
		Workbook workbook = null;
		if(fileNameExtension.equalsIgnoreCase(".xlsx")) {
			workbook = new XSSFWorkbook(inputStream);
		}else 
			if(fileNameExtension.equalsIgnoreCase(".xls")) {
				workbook = new HSSFWorkbook(inputStream);
			}
		
		Sheet sheet = workbook.getSheet(sheetname);
		int excelRow = sheet.getLastRowNum()-sheet.getFirstRowNum();
		
		if(excelRow==0) {
			cricbuzzScore emptyObj =  new cricbuzzScore();
			emptyObj.emptyTable(colCount, valueToWrite, sheet, excelRowCount);
		}
		else {
			cricbuzzScore dataObj =  new cricbuzzScore();
			dataObj.dataTable(colCount, valueToWrite, sheet, excelRow);
		}
				
		inputStream.close();
		FileOutputStream outputStream = new FileOutputStream(file);
		workbook.write(outputStream);
		outputStream.close();		
	}

public void emptyTable(int colCount, String[] valueToWrite, Sheet sheet, int excelRowCount) {
	
	Row newRow = sheet.createRow(excelRowCount);
	for (int j=0; j<colCount;j++) {
		Cell cell = newRow.createCell(j);
		cell.setCellValue(valueToWrite[j]);
	}
	System.out.println("Data written to excel File in row number: " + (excelRowCount+1));	
}

public void dataTable(int colCount, String[] valueToWrite, Sheet sheet, int excelRow) {
	
	Row newRow = sheet.createRow(excelRow+1);
	for (int j=0; j<colCount;j++) {
		Cell cell = newRow.createCell(j);
		cell.setCellValue(valueToWrite[j]);
	}
	System.out.println("Data written to excel File in row number: " + (excelRow+2));
}

	public static void main(String[] args) throws IOException {
		
		WebDriver driver = new FirefoxDriver();
		driver.get("http://www.cricbuzz.com/live-cricket-scorecard/20348/afg-vs-ban-3rd-t20i-afghanistan-v-bangladesh-in-india-2018");
				
		int x = 3;
		int rowCount=-3;
		String rowStart = "html/body/div[1]/div[2]/div[4]/div[2]/div[2]/div[1]/div[";
		String rowEnd = "]/div[1]";
		
		try{
			while(driver.findElements(By.xpath(rowStart+x+rowEnd)).size() !=0) {			
			x++;
			rowCount++;
			}
		}
		catch (NoSuchElementException e) {
			System.out.println("No element found");
		}
		
		System.out.println("Row count: " + rowCount);
		
		int y = 1;
		int colCount=0;
		String colStart = "html/body/div[1]/div[2]/div[4]/div[2]/div[2]/div[1]/div[3]/div[";
		String colEnd = "]";
		
		while(driver.findElements(By.xpath(colStart+y+colEnd)).size() !=0) {			
			y++;
			colCount++;
		}
		System.out.println("Cols count: " + colCount);
		
		String tableStart = "html/body/div[1]/div[2]/div[4]/div[2]/div[2]/div[1]/div[";
		String tableMid = "]/div[";
		String tableEnd = "]";
		cricbuzzScore objFile = new cricbuzzScore();
		String filepath="C:\\Users\\ankitkumagarwal\\Desktop\\Selenium\\workspace\\OnlineWebtables";
		int excelRowCount = 0;
		int totalRuns=0;
		String valueToWrite[] = new String[colCount];
		
		for(int a=3;a<=(rowCount+2);a++) {
			
			for(int b=1;b<=colCount;b++) {
				
	//Storing the value of a column in an array
				valueToWrite[b-1] = driver.findElement(By.xpath(tableStart+a+tableMid+b+tableEnd)).getText();
							
				System.out.print((driver.findElement(By.xpath(tableStart+a+tableMid+b+tableEnd)).getText())+"        ");
				if(b==3) {
					int runs = Integer.parseInt(valueToWrite[b-1]);
					totalRuns = totalRuns+runs;
				}
				
			}
	//Calling method to write values into excel
			
			objFile.writeToExcel(filepath,"cricbuzz.xlsx","Sheet1",valueToWrite,colCount,excelRowCount);
			excelRowCount++;
			System.out.println();
			System.out.println("------------------------------------------------------------------");
		}
		System.out.println();
		
	//Compare the total calculated out of each runs column with the total displayed in the website
		
		int extras = Integer.parseInt((driver.findElement(By.xpath("html/body/div[1]/div[2]/div[4]/div[2]/div[2]/div[1]/div[11]/div[2]"))).getText());
		totalRuns = totalRuns+extras;
		System.out.println(totalRuns);
		int webTotalRuns = Integer.parseInt((driver.findElement(By.xpath("html/body/div[1]/div[2]/div[4]/div[2]/div[2]/div[1]/div[12]/div[2]"))).getText());
		if(totalRuns==webTotalRuns) {
			System.out.println("Total Matches");
		}else {
			System.out.println("Total Mismatch");
		}
		
	}


}
