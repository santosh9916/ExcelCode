package Property_Excel_relatedCode;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

import javafx.scene.layout.Priority;

import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;

public class Sample {
	public WebDriver driver;
	public String str;
	String file = "D:\\workspaceFinal\\Selenium3_Latest\\src\\Property_Excel_relatedCode\\ReadMultipleTestdata.xlsx";
	
	@BeforeMethod
	public void init() {
		System.setProperty("webdriver.chrome.driver", "D:\\Library\\chrome_76\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS);
		driver.get("https://www.flipkart.com");
	}
	@Test
	public void verifyLogin() throws Exception {
		// TimeStamp
		DateFormat df = new SimpleDateFormat("yyyy_MMM_dd hh_mm_ss a");
		Date d = new Date();
		String time = df.format(d);
		// To read the test data file
		FileInputStream fi = new FileInputStream(new File(file));
		XSSFWorkbook workbook = new XSSFWorkbook(fi);
		XSSFSheet sheet = workbook.getSheetAt(0);
		

		for (int i = 1; i < sheet.getLastRowNum(); i++) {
			
			// Enter username,Password and click on signin by taking data from
			// testdata file
			driver.findElement(By.xpath("//input[@class='_2zrpKA _1dBPDZ']"))
					.sendKeys(sheet.getRow(0).getCell(i).getStringCellValue());
			driver.findElement(By.xpath("//input[@class='_2zrpKA _3v41xv _1dBPDZ']"))
					.sendKeys(sheet.getRow(1).getCell(i).getStringCellValue());
			driver.findElement(By.xpath("//button[@class='_2AkmmA _1LctnI _7UHT_c']//span[contains(text(),'Login')]"))
					.click();

			// Validate Logout, if available assign pass to str, else assign
			// fail to str
			try {
				Actions actions = new Actions(driver);
				actions.moveToElement(driver.findElement(By
						.xpath("//body/div[@id='container']/div/div[@class='_3ybBIU']/div[@class='_1tz-RS']/div[@class='_3pNZKl']/div[3]/div[1]")))
						.build().perform();
				driver.findElement(By.xpath("//div[contains(text(),'Logout')]")).click();
				// str="Pass";
			} catch (Exception e) {
				// str="Fail";
				e.getMessage();
			}

		}
	}

	
}
