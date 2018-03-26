package CommonProperties;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.*;


public class ObjectProperties {
	
	protected static XSSFSheet sheet;
	protected static XSSFWorkbook workbook1;
	protected static XSSFCell Cell;
    
	public static void setExcelFile(String SheetName) throws Exception {
        FileInputStream src = new FileInputStream("D:\\New Workspace\\Scheduling\\Input_Sheet_Scheduling.xlsx");
        workbook1 = new XSSFWorkbook(src);
        sheet = workbook1.getSheet(SheetName);
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
