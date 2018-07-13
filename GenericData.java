package org.openqa.selenium.winium;
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.*;
import org.openqa.selenium.By;
import org.sikuli.script.Pattern;

public class GenericData {
	
	protected static XSSFSheet sheet;
	protected static XSSFWorkbook workbook1;
	protected static XSSFCell Cell;
	protected static File InputSheet= new File("D:\\New Workspace\\WPFAutomation\\WPF_InputSheet.xlsx");
	protected static String DefaultTemplate;
	
	public static void setExcelFile(String SheetName) throws Exception {
		
        FileInputStream src = new FileInputStream(InputSheet);
        workbook1 = new XSSFWorkbook(src);
        sheet = workbook1.getSheet(SheetName);
       }

	//Patterns
	protected Pattern Macro_ScrollBarBtn = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Macro_ScrollBar.png");
	protected Pattern Macro_ScrollDownBtn = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Macro_ScrollDown.png");
	
	public void login() throws Exception {
        try {
            CommonPropertiesClass.setExcelFile("Login");

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                XSSFCell cell = sheet.getRow(i).getCell(0);

                driver.findElement(Username).sendKeys(cell.getStringCellValue());
                cell = sheet.getRow(i).getCell(1);
                driver.findElement(Password).sendKeys(cell.getStringCellValue());
                driver.findElement(LoginBtn).click();
                Thread.sleep(3000);
            }
        } catch (Exception e) {
            
        }
    }
	
	
	 public void SelectBrowser() throws Exception {
        try {
            this.DeleteProcess();
            CommonPropertiesClass.setExcelFile("Browser");

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                XSSFCell cell    = sheet.getRow(i).getCell(0);
                String   Browser = cell.getStringCellValue();

                if (Browser.equals("IE")) {
                    System.setProperty("webdriver.ie.driver", "D:\\Jars for Selenium\\IEDriverServer.exe");
                    driver = new InternetExplorerDriver();
                } else if (Browser.equals("Chrome")) {
                    System.setProperty("webdriver.chrome.driver", "D:\\Jars for Selenium\\chromedriver.exe");
                    driver = new ChromeDriver();
                } else if (Browser.equals("Firefox")) {
                    System.setProperty("webdriver.gecko.driver", "D:\\Jars for Selenium\\geckodriver.exe");
                    driver = new FirefoxDriver();
                } else {
                    System.out.println("Browser not available");
                }
            }
        } catch (Exception e) {
            
        }
		 
		 
	 public boolean isElementPresent(By element) {
        boolean status = false;

        try {
            driver.findElement(element).isDisplayed();
            status = true;
        } catch (NoSuchElementException e) {
            status = false;
        }

        return status;
    }
		 
		 
		 
		 
	public boolean VerifyTextPresent(String value)
	{
		boolean status = false;
		String PageSourse= driver.getPageSource();
		if(PageSourse.contains(value))
		{
			status = false;
		}
		else
		{
			status = true;
		}
		return status;
	}
	
}

