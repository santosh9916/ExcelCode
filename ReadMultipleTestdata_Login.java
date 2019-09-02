package Property_Excel_relatedCode;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.DataProvider;

public class ReadMultipleTestdata_Login {
	
	private static WebDriver driver;
	String excel_file = "D:\\workspaceFinal\\Selenium3_Latest\\src\\Property_Excel_relatedCode\\ReadMultipleTestdata.xlsx";
	FileInputStream file;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	
	public void init() {
		System.setProperty("webdriver.chrome.driver", "D:\\Library\\chrome_76\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS);
		driver.get("https://www.flipkart.com");
	}
	// Not completed successfully
	public  Object[][] getTestData() {
		FileInputStream file = null;
		try {
			file = new FileInputStream(new File(excel_file));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		try {
			workbook = new XSSFWorkbook(file);
		} catch (IOException e) {
			e.printStackTrace();
		}
		sheet = workbook.getSheet("Sheet1");
		WebElement uname = driver.findElement(By.xpath("//input[@class='_2zrpKA _1dBPDZ']"));
		WebElement pwd = driver.findElement(By.xpath("//input[@class='_2zrpKA _3v41xv _1dBPDZ']"));
		WebElement loginbtn = driver.findElement(By.xpath("//button[@class='_2AkmmA _1LctnI _7UHT_c']//span[contains(text(),'Login')]"));
		Object[][] data = new Object[sheet.getLastRowNum()][sheet.getRow(1).getLastCellNum()];
		// System.out.println(sheet.getLastRowNum() + "--------" +
		// sheet.getRow(0).getLastCellNum());
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			for (int k = 0; k < sheet.getRow(1).getLastCellNum(); k++) {
				data[i][k] = sheet.getRow(i + 1).getCell(k).toString();
				// System.out.println(data[i][k]);
				uname.sendKeys();
				pwd.sendKeys();
				loginbtn.click();
			}
		}
		return data;
	}
	@DataProvider
	public void login_flipkart(){
		
		
	}
public static void main(String[] args) {
	ReadMultipleTestdata_Login login = new ReadMultipleTestdata_Login();
	login.init();
	
}
}
