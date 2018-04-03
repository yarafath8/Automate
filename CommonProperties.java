package CommonProperties;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
import org.apache.poi.xssf.usermodel.*;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;
import org.openqa.selenium.winium.WiniumDriverService;
import org.openqa.selenium.os.WindowsUtils;
import org.openqa.selenium.winium.*;



public class ObjectProperties {
	
	protected static XSSFSheet sheet;
	protected static XSSFWorkbook workbook1;
	protected static XSSFCell Cell;
    
	public static void setExcelFile(String SheetName) throws Exception {
        FileInputStream src = new FileInputStream(filepath);
        workbook1 = new XSSFWorkbook(src);
        sheet = workbook1.getSheet(SheetName);
       }
	
	public void OpenWindowsApplication() throws IOException, Exception {
	 
	 DesktopOptions options= new DesktopOptions();
     options.setApplicationPath(EXEPath);
     WiniumDriverService service = new WiniumDriverService.Builder().usingPort(9999).withVerbose(true).withSilent(false).buildDesktopService();
	 WPF1=new WiniumDriver(service,options);
     
 }

	File src=new File("D:\\New Workspace\\Scheduling\\ObjectRepo");	  
	FileInputStream fis;
	protected static Properties pro=new Properties();
	{
	try {
		fis = new FileInputStream(src);
	} catch (FileNotFoundException e) {
		e.printStackTrace();
	}
	
	try {
		pro.load(fis);
	} catch (IOException e) {
		e.printStackTrace();
	}
	System.out.println("Property class loaded");
	
	}	
}
