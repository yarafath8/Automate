package org.openqa.selenium.winium;
import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.*;
import org.openqa.selenium.By;
import org.sikuli.script.Pattern;

public class CommonProperties {
	
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

	//Advance Options
	protected By AdvanceOptions = By.id("txtAdvOption");
	protected By ProdURL = By.id("txtHost");
	protected By ProdURLResetBtn = By.id("btnReset");
	protected By ProdURLResetOKBtn = By.id("2");
	protected By TestURL = By.id("txtHostTest");
	protected By TestURLResetBtn = By.id("btnResetTest");
	protected By TestURLResetOKBtn = By.id("2");
	protected By AdvanceOptionsCloseBtn = By.id("Button_1");
	
	//Login Section
	protected By Username = By.className("TextBox");
	protected By Username1 = By.id("txtName");
	protected By Password = By.className("PasswordBox");
	protected By Password1 = By.id("pwbPassword");
	protected By Login = By.id("LoginButton");
	protected By ServerDropdown = By.id("cboxServer");
	
	//Meta data
	protected By MetaData = By.name("Updating Metadata...56%");
	
	//Confirm Modal
	protected By RoomDropdown = By.id("roomloc");
	protected By OKBtn = By.id("Button_1");
	protected By UnknownBtn = By.id("Button_2");
	protected By OR_Room = By.xpath("//*[@Name='OR2']");
	protected By OR_Room1 = By.name("OR1");
	
	//Server Conflict
	protected By ConflictYesBtn = By.id("6");
	
	//New Patient Modal //*[@id='email' or @name='email']
	protected By PatientList_NewPatient = By.id("TextBlock_28");
	protected By PatientList_NewPatient1 = By.name("New Patient");
	protected By MRNbox = By.id("tboxMRN");
	protected By AccountNo = By.id("tboxAccountNumber");
	protected By FirstName = By.id("tboxFirstName");
	protected By MiddleName = By.id("tboxMiddleName");
	protected By LastName = By.id("tboxLastName");
	protected By DOBDate = By.id("Btn");
	protected By GenderDropdown = By.id("coboxSex");
	protected By GenderDropdown1 = By.className("ComboBox");
	protected By GenderMale = By.id("ComboBoxItem_1");
	protected By GenderFemale = By.id("ComboBoxItem_2");
	protected By OKbtn = By.id("btnOK");
	protected By Emergency = By.id("btnEm");
	
	//IntraOp
	protected By IntraOp_Max = By.id("PART_MaximizeButton");
	protected By IntraOp_Maximize = By.id("imgMax");
	protected By IntraOp_ExitBtn = By.id("btnExit"); 
	protected By ExitBtn_Yes = By.id("Button_1"); 
	protected By Macro1 = By.id("btn"); 
	protected By Macro_Save = By.id("btnSave");
	protected By Macro_ScrollUp = By.id("LineUp");
	protected By Macro_ScrollDown = By.id("LineDown");
	protected By Macro_ScrollBar = By.name("Thumb");
	protected By Macro_Space = By.id("ScrollViewer_1");
	protected By Case_Stuff = By.id("TextBlock_7");
	protected By CaseStaff_Signin_Yes = By.name("Yes");
	protected By CaseStaff_Signin_No = By.name("No");
	
	//Restore Session
	protected By Restore_Previous_Session_Modal= By.name("Your last PlexusChart session closed unexpectedly.  Do you want to restore from your previous session?");
	protected By Restore_Previous_Session_Yes = By.name("Yes");
	protected By Restore_Previous_Session_No = By.name("No");
	
	//Case Staff
	protected By Signature_Window = By.id("personnelWindow");
	protected By CaseStaff_CloseBtn = By.id("btnClose");
	protected By CaseStaff_SignBtn = By.id("btnSign");
	protected By CaseStaff_UnsignBtn = By.id("btnUnsign");
	protected By CaseStaff_ClickToSign = By.name("<Click to Sign>");
	
	//Sign in Modal
	protected By SignModal  = By.xpath("//*[contains(@id,'tbMessage') or contains(@name, 'You are currently not signed onto the case')]");
	protected By Signin_Modal = By.id("tbMessage");
	protected By Signin_Modal_Message = By.name("You are currently not signed onto the case. Would you like to sign in?");
	protected By Signin_Modal_Yes = By.name("Yes");
	protected By Signin_Modal_No = By.name("No");
	
	//Statements
	protected By Statement_Induction  = By.name("I was present for induction");
	protected By Statement_Available = By.name("I was immediately available");
	protected By Statement_Emergence = By.name("I was present for emergence");
	protected By Statement_medical = By.name("I was medically directing case");
	protected By Statement_Save = By.id("Button_1");
	
	//Template
	protected By SelectTemlpate = By.id("LabTitle");
	protected By TemplateName = By.name("OB");
	 
	//MacroData
	protected By MacroTextBox = By.className("TextBox"); 
	protected By MacroTabs_Add = By.className("Add"); 
	protected By MacroScrollDownBtn = By.id("LineDown");
	protected By MacroScrollUpBtn = By.id("LineUp");
	protected By MacroScrollBar = By.className("Thumb");
	
	//Patterns
	protected Pattern Macro_ScrollBarBtn = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Macro_ScrollBar.png");
	protected Pattern Macro_ScrollDownBtn = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Macro_ScrollDown.png");
	/*protected Pattern Conflict_Yes= new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Conflict_Yes.png");
	protected Pattern Conflict_No= new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Conflict_No.png");
	protected Pattern Confirm_OK= new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Confirm_OK.png");
	protected Pattern Confirm_Unknown = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Confirm_Unknown.png");
	protected Pattern Macro_ScrollBarBtn = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Macro_ScrollBar.png");
	protected Pattern Macro_ScrollDownBtn = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Macro_ScrollDown.png");
	protected Pattern Checkbox_Checked = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Checkbox_Checked.png");
	protected Pattern Checkbox_Unchecked = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Checkbox_Unchecked.png");
	protected Pattern Signature_Modal = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Signature_Modal.png");
	protected Pattern Restore_Previous_Session = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\Restore_Previous_Session.png");
	protected Pattern Macro_TextBox = new Pattern("D:\\New Workspace\\WPFAutomation\\images\\TextBox.png");*/
	
	
	
}

