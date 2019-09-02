package Property_Excel_relatedCode;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class PrintLinksInto_Excel {
	public WebDriver driver;
	String excel_file = "D:\\workspaceFinal\\Selenium3_Latest\\src\\Property_Excel_relatedCode\\Testdata2.xlsx";

	@Test(priority = 1)
	public void init() {
		System.setProperty("webdriver.chrome.driver", "D:\\Library\\chrome_76\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS);
		driver.get("https://www.google.com");

	}

	@Test(priority = 2)
	public void printLinks() throws IOException {

		List<WebElement> links = driver.findElements(By.tagName("a"));
		System.out.println(links.size());
		FileInputStream fis = new FileInputStream(new File(excel_file));

		// Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		// Create a blank sheet
		//XSSFSheet sheet = workbook.createSheet("Links");
		
		// For existing sheet
		//XSSFSheet sheet = workbook.getSheetAt(2);
		// OR
		XSSFSheet sheet = workbook.getSheet("Links");
		
		int rownum = 1;
		for (WebElement link : links) {
			// this creates a new row in the sheet
			Row row = sheet.createRow(rownum++);
			// this line creates a cell in the next column of that row
			Cell cell = row.createCell(2);
			if(!link.getText().equalsIgnoreCase("")){
			cell.setCellValue(link.getText());
			System.out.println("Links are :" + link.getText());
			}
		}
		try {
			// this Writes the workbook gfgcontribute
			FileOutputStream out = new FileOutputStream(new File(excel_file));
			workbook.write(out);
			out.close();
			System.out.println("gfgcontribute.xlsx written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
