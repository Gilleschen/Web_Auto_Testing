package AutoTesting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import java.net.URL;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;

import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class method {
	static LoadTestCase TestCase = new LoadTestCase();
	LoadExpectResult ExpectResult = new LoadExpectResult();

	static String CaseErrorList[][] = new String[TestCase.CaseList.size()][TestCase.DeviceInformation.BrowserList
			.size()];// �����U�רҩ�U�˸m�����O���G (2���}�C)CaseErrorList[CaseList][Devices]
	String ErrorList[] = new String[TestCase.DeviceInformation.BrowserList.size()];// �����U�˸m�����O���G

	static int port = 5555;
	static int command_timeout = 15;// 30sec
	static String appElemnt;// APP����W��
	static String appInput;// ��J��
	static String appInputXpath;// ��J�Ȫ�Xpath�榡
	static WebDriver driver[] = new WebDriver[TestCase.DeviceInformation.BrowserList.size()];
	static WebDriverWait[] wait = new WebDriverWait[TestCase.DeviceInformation.BrowserList.size()];
	String element[] = new String[driver.length];
	static int CurrentCaseNumber = -1;// �ثe�����ĴX�Ӵ��ծצC
	static Boolean CommandError = true;// �P�w���檺���O�O�_�X�{���~�Fture�����T�Ffalse�����~
	XSSFSheet Sheet;
	XSSFWorkbook workBook;

	public static void main(String[] args) throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException, IOException {
		initial();
		invokeFunction();
		System.out.println("���յ���!!!!!!!!");
		Process proc = Runtime.getRuntime().exec("explorer C:\\TUTK_QA_TestTool\\TestReport");// �}��TestReport��Ƨ�
	}

	public static void initial() {// ��l��CaseErrorList�x�}
		for (int i = 0; i < CaseErrorList.length; i++) {
			for (int j = 0; j < CaseErrorList[i].length; j++) {
				CaseErrorList[i][j] = "";
			}
		}
	}

	public static void invokeFunction() throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException {
		Object obj = new method();
		Class c = obj.getClass();
		String methodName = null;

		for (int CurrentCase = 0; CurrentCase < TestCase.StepList.size(); CurrentCase++) {
			CommandError = true;// �w�]CommandError��True
			for (int CurrentCaseStep = 0; CurrentCaseStep < TestCase.StepList.get(CurrentCase)
					.size(); CurrentCaseStep++) {
				if (!CommandError) {
					break;// �Y�ثe���ծרҥX�{CommandError=false�A�h���X�ثe�רҨð���U�@�Ӯר�
				}
				switch (TestCase.StepList.get(CurrentCase).get(CurrentCaseStep).toString()) {

				case "Launch":
					methodName = "Launch";
					break;

				case "Byid_SendKey":
					methodName = "Byid_SendKey";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

				case "Byid_Click":
					methodName = "Byid_Click";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_SendKey":
					methodName = "ByXpath_SendKey";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

				case "ByXpath_Click":
					methodName = "ByXpath_Click";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Byid_Wait":
					methodName = "Byid_Wait";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Wait":
					methodName = "ByXpath_Wait";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Byid_Result":
					methodName = "Byid_Result";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Result":
					methodName = "ByXpath_Result";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Scroll":
					methodName = "ByXpath_Scroll";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Byid_Scroll":
					methodName = "Byid_Scroll";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Sleep":
					methodName = "Sleep";
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ScreenShot":
					methodName = "ScreenShot";
					break;

				case "Quit":
					methodName = "Quit";
					break;

				}

				Method method;
				method = c.getMethod(methodName, new Class[] {});
				method.invoke(c.newInstance());
			}
		}
	}

	public void Byid_Result() {
		boolean result[] = new boolean[driver.length];// �����wBoolean�ȡA�w�]��False
		boolean ErrorResult[] = new boolean[driver.length];

		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				element[i] = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.id(appElemnt))).getText();
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				element[i] = "ERROR";// �䤣��Ӫ���A�^��Error
			}

			if (element[i].equals("ERROR")) {
				ErrorResult[i] = true;

			} else {
				// �^�Ǵ��ծרҲM�檺�W�ٵ�ExpectResult.LoadExpectResult�A�æs����浲�G��ResultList�M��
				ExpectResult.LoadExpectResult(TestCase.CaseList.get(CurrentCaseNumber).toString());
				for (int j = 0; j < ExpectResult.ResultList.size(); j++) {
					if (element[i].equals(ExpectResult.ResultList.get(j)) == true) {
						result[i] = true;
						break;
					} else {
						result[i] = false;
					}
				}
			}
		}
		SubMethod_Result(ErrorResult, result);// �I�ssubmethod_result�x�s���յ��G��Excel
		// CurrentCaseNumber = CurrentCaseNumber + 1;

	}

	public void ByXpath_Result() {
		boolean result[] = new boolean[driver.length];// �����wBoolean�ȡA�w�]��False
		boolean ErrorResult[] = new boolean[driver.length];

		for (int i = 0; i < driver.length; i++) {

			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				element[i] = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)))
						.getText();

			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				element[i] = "ERROR";// �䤣��Ӫ���A�^��Error
			}

			if (element[i].equals("ERROR")) {
				ErrorResult[i] = true;

			} else {
				// �^�Ǵ��ծרҲM�檺�W�ٵ�ExpectResult.LoadExpectResult�A�æs����浲�G��ResultList�M��
				ExpectResult.LoadExpectResult(TestCase.CaseList.get(CurrentCaseNumber).toString());
				for (int j = 0; j < ExpectResult.ResultList.size(); j++) {
					if (element[i].equals(ExpectResult.ResultList.get(j)) == true) {
						result[i] = true;
						break;
					} else {
						result[i] = false;
					}
				}
			}

		}
		SubMethod_Result(ErrorResult, result);

		// CurrentCaseNumber = CurrentCaseNumber + 1;

	}

	public void Byid_Click() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.id(appElemnt))).click();
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
				driver[i].quit();// �Y�䤣����w����A�h����Browser
			}
		}
	}

	public void ByXpath_Click() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).click();
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
				driver[i].quit();// �Y�䤣����w����A�h����Browser
			}
		}
	}

	public void Byid_SendKey() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.id(appElemnt))).sendKeys(appInput);
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
				driver[i].quit();// �Y�䤣����w����A�h����Browser
			}
		}
	}

	public void ByXpath_SendKey() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).sendKeys(appInput);
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
				driver[i].quit();// �Y�䤣����w����A�h����Browser
			}
		}
	}

	public void ByXpath_Scroll() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				WebElement target = wait[i].until(ExpectedConditions.presenceOfElementLocated((By.xpath(appElemnt))));
				Actions actions = new Actions(driver[i]);
				actions.moveToElement(target);
				// actions.click(target);
				actions.perform();
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
				driver[i].quit();// �Y�䤣����w����A�h����Browser
			}
		}
	}

	public void Byid_Scroll() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				WebElement target = wait[i].until(ExpectedConditions.presenceOfElementLocated((By.id(appElemnt))));
				Actions actions = new Actions(driver[i]);
				actions.moveToElement(target);
				// actions.click(target);
				actions.perform();
				ErrorList[i] = "Pass";
				CaseErrorList[CurrentCaseNumber] = ErrorList;
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
				driver[i].quit();// �Y�䤣����w����A�h����Browser
			}
		}
	}

	public void Launch() {
		CurrentCaseNumber = CurrentCaseNumber + 1;
		DesiredCapabilities cap[] = new DesiredCapabilities[TestCase.DeviceInformation.BrowserList.size()];

		for (int i = 0; i < driver.length; i++) {
			cap[i] = new DesiredCapabilities();
			cap[i].setBrowserName(TestCase.DeviceInformation.BrowserList.get(i));

			try {
				driver[i] = new RemoteWebDriver(new URL("http://localhost:" + port + "/wd/hub"), cap[i]);
				driver[i].manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
				driver[i].manage().window().maximize();
				driver[i].get(TestCase.DeviceInformation.URL);

			} catch (Exception e1) {
				System.out.println("[Error] Can't launch " + TestCase.DeviceInformation.BrowserList.get(i).toString());
			}
			// port++;
		}
	}

	public void Quit() {

		for (int j = 0; j < driver.length; j++) {
			driver[j].quit();// ���}APP��A�g�J���յ��GPass��Error

			// �}��Excel
			try {
				workBook = new XSSFWorkbook(
						new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm"));
			} catch (Exception e) {
				System.out.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm");
			}
			for (int i = 0; i < driver.length; i++) {

				if (TestCase.DeviceInformation.BrowserList.get(i).toString().length() > 20) {// Excel�u�@��W�ٳ̱`31�r���]�A�G�ݧP�_BrowserName���׬O�_�j��31
					char[] NewBrowserName = new char[20];// �]�ݥ]�t_TestReport�r��(�@11�r��)�A�G�]�w20��r���}�C(31-11)
					TestCase.DeviceInformation.BrowserList.get(i).toString().getChars(0, 20, NewBrowserName, 0);// ���XBrowserName�e20�r����NewBrowserName
					Sheet = workBook.getSheet(String.valueOf(NewBrowserName) + "_TestReport");// �ھ�NewUdid�A���w�Y�x�˸m��TestReport
					// sheet
				} else {
					Sheet = workBook.getSheet(TestCase.DeviceInformation.BrowserList.get(i).toString() + "_TestReport");// ���w�Y�x�˸m��TestReport
																														// sheet
				}

				if (CaseErrorList[CurrentCaseNumber][i].equals("Pass")) {// ���XCaseErrorList����CurrentCaseNumber�Ӵ���������i�x��ʸ˸m�����G
					Sheet.getRow(CurrentCaseNumber + 1).getCell(1).setCellValue("Pass");// ��J��i�x��ʸ˸m����CurrentCaseNumber�Ӵ������GPass
				}

			}
			// ����g�JExcel�᪺�s�ɰʧ@
			try {
				FileOutputStream out = new FileOutputStream(
						new File("C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm"));
				workBook.write(out);
				out.close();
				workBook.close();
			} catch (Exception e) {
				System.out.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm");
			}

		}

	}

	public void Sleep() {
		String NewString = "";// �s�r��
		char[] r = { '.' };// �p���I�r��
		char[] c = appInput.toCharArray();// �N�r���ন�r���}�C
		for (int i = 0; i < c.length; i++) {
			if (c[i] != r[0]) {// �P�_�r���O�_���p���I
				NewString = NewString + c[i];// �_�A�N�r���զX���s�r��
			} else {
				break;// �O�A���X�j��
			}
		}

		try {
			System.out.println("[driver] [start] Sleep(): " + NewString + " second...");
			Thread.sleep(Integer.valueOf(NewString) * 1000);// �N�r���ন���
			System.out.println("[driver] [end] Sleep");
		} catch (Exception e) {
			;
		}
	}

	public void ScreenShot() {

		Calendar date = Calendar.getInstance();
		String month = Integer.toString(date.get(Calendar.MONTH) + 1);
		String day = Integer.toString(date.get(Calendar.DAY_OF_MONTH));
		String hour = Integer.toString(date.get(Calendar.HOUR_OF_DAY));
		String min = Integer.toString(date.get(Calendar.MINUTE));
		String sec = Integer.toString(date.get(Calendar.SECOND));
		for (int i = 0; i < driver.length; i++) {
			File screenShotFile = (File) ((TakesScreenshot) driver[i]).getScreenshotAs(OutputType.FILE);

			try {
				FileUtils.copyFile(screenShotFile, new File("C:\\TUTK_QA_TestTool\\TestReport\\"
						+ TestCase.CaseList.get(CurrentCaseNumber) + "_" + month + day + hour + min + sec + ".jpg"));
				System.out.println("[Log] " + "ScreenShot Successfully!! (Name:CaseName+Month+Day+Hour+Minus+Second)");

			} catch (IOException e) {
				;
			}
		}
	}

	public void SubMethod_Result(boolean ErrorResult[], boolean result[]) {
		// �}��Excel
		try {
			workBook = new XSSFWorkbook(new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm"));
		} catch (Exception e) {
			System.out.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm");
		}
		for (int i = 0; i < driver.length; i++) {

			if (TestCase.DeviceInformation.BrowserList.get(i).toString().length() > 20) {// Excel�u�@��W�ٳ̱`31�r���]�A�G�ݧP�_UDID���׬O�_�j��31
				char[] NewUdid = new char[20];// �]�ݥ]�t_TestReport�r��(�@11�r��)�A�G�]�w20��r���}�C(31-11)
				TestCase.DeviceInformation.BrowserList.get(i).toString().getChars(0, 20, NewUdid, 0);// ���XUDID�e20�r����NewUdid
				Sheet = workBook.getSheet(String.valueOf(NewUdid) + "_TestReport");// �ھ�NewUdid�A���w�Y�x�˸m��TestReport
																					// sheet
			} else {
				Sheet = workBook.getSheet(TestCase.DeviceInformation.BrowserList.get(i).toString() + "_TestReport");// ���w�Y�x�˸m��TestReport
																													// sheet
			}

			if (ErrorResult[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Error!!");
			} else if (result[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Pass");
			} else if (result[i] == false) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Fail");
			}
		}
		// ����g�JExcel�᪺�s�ɰʧ@
		try {
			FileOutputStream out = new FileOutputStream(
					new File("C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm"));
			workBook.write(out);
			out.close();
			workBook.close();
		} catch (Exception e) {
			System.out.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm");
		}
	}
}