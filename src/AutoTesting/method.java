package AutoTesting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Platform;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import junit.framework.TestCase;

public class method {
	static LoadTestCase TestCase = new LoadTestCase();
	LoadExpectResult ExpectResult = new LoadExpectResult();
	static int port = 5555;
	static int command_timeout = 30;// 30sec
	static String appElemnt;// APP元件名稱
	static String appInput;// 輸入值
	static String appInputXpath;// 輸入值的Xpath格式
	static WebDriver driver[] = new WebDriver[TestCase.DeviceInformation.BrowserList.size()];
	static WebDriverWait[] wait = new WebDriverWait[TestCase.DeviceInformation.BrowserList.size()];
	String element[] = new String[driver.length];
	static int CurrentCaseNumber = -1;// 目前執行到第幾個測試案列
	XSSFSheet Sheet;
	XSSFWorkbook workBook;

	public static void main(String[] args) throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException, IOException {

		invokeFunction();
		System.out.println("測試結束!!!!!!!!");
		// Process proc = Runtime.getRuntime().exec("explorer
		// C:\\TUTK_QA_TestTool\\TestReport");// 開啟TestReport資料夾
	}

	public static void invokeFunction() throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException {
		Object obj = new method();
		Class c = obj.getClass();
		String methodName = null;

		for (int i = 0; i < TestCase.StepList.size(); i++) {

			switch (TestCase.StepList.get(i).toString()) {

			case "Launch":
				methodName = "Launch";
				break;

			case "Byid_SendKey":
				methodName = "Byid_SendKey";
				appElemnt = TestCase.StepList.get(i + 1);
				appInput = TestCase.StepList.get(i + 2);
				i = i + 2;
				break;

			case "Byid_Click":
				methodName = "Byid_Click";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "ByXpath_SendKey":
				methodName = "ByXpath_SendKey";
				appElemnt = TestCase.StepList.get(i + 1);
				appInput = TestCase.StepList.get(i + 2);
				i = i + 2;
				break;

			case "ByXpath_Click":
				methodName = "ByXpath_Click";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "Byid_Wait":
				methodName = "Byid_Wait";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "ByXpath_Wait":
				methodName = "ByXpath_Wait";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "Byid_Result":
				methodName = "Byid_Result";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "ByXpath_Result":
				methodName = "ByXpath_Result";
				appElemnt = TestCase.StepList.get(i + 1);
				i = i + 1;
				break;

			case "Sleep":
				methodName = "Sleep";
				appInput = TestCase.StepList.get(i + 1);
				i = i + 1;
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

	public void Byid_Result() {
		boolean result[] = new boolean[driver.length];// 未給定Boolean值，預設為False
		boolean ErrorResult[] = new boolean[driver.length];

		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				element[i] = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.id(appElemnt))).getText();
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				element[i] = "ERROR";// 找不到該物件，回傳Error
			}

			if (element[i].equals("ERROR")) {
				ErrorResult[i] = true;

			} else {
				// 回傳測試案例清單的名稱給ExpectResult.LoadExpectResult，並存放期望結果至ResultList清單
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
		SubMethod_Result(ErrorResult, result);// 呼叫submethod_result儲存測試結果於Excel
		// CurrentCaseNumber = CurrentCaseNumber + 1;

	}

	public void ByXpath_Result() {
		boolean result[] = new boolean[driver.length];// 未給定Boolean值，預設為False
		boolean ErrorResult[] = new boolean[driver.length];

		for (int i = 0; i < driver.length; i++) {

			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				element[i] =wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).getText();
				
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
				element[i] = "ERROR";// 找不到該物件，回傳Error
			}

			if (element[i].equals("ERROR")) {
				ErrorResult[i] = true;

			} else {
				// 回傳測試案例清單的名稱給ExpectResult.LoadExpectResult，並存放期望結果至ResultList清單
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
				// driver[i].findElement(By.xpath(appElemnt)).click();
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
			}
		}
	}

	public void ByXpath_Click() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).click();
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
			}
		}
	}

	public void Byid_SendKey() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.id(appElemnt))).sendKeys(appInput);
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
			}
		}
	}

	public void ByXpath_SendKey() {
		for (int i = 0; i < driver.length; i++) {
			try {
				wait[i] = new WebDriverWait(driver[i], command_timeout);
				wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).sendKeys(appInput);
			} catch (Exception e) {
				System.out.println("[Error] Can't find " + appElemnt);
			}
		}
	}

	public void Launch() {
		CurrentCaseNumber = CurrentCaseNumber + 1;
		DesiredCapabilities cap[] = new DesiredCapabilities[TestCase.DeviceInformation.BrowserList.size()];
		//for (int i = 0; i < driver.length; i++) {
			//cap[i] = new DesiredCapabilities();
		//}

		for (int i = 0; i < driver.length; i++) {
			cap[i] = new DesiredCapabilities();
			cap[i].setBrowserName(TestCase.DeviceInformation.BrowserList.get(i));

			try {
				driver[i] = new RemoteWebDriver(new URL("http://localhost:" + port + "/wd/hub"), cap[i]);
				driver[i].manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
				driver[i].manage().window().maximize();
				driver[i].get(TestCase.DeviceInformation.URL);

			} catch (Exception e1) {
				;
			}
			// port++;
		}
	}

	public void Quit() {
		for (int i = 0; i < driver.length; i++) {
			driver[i].quit();
		}
	}

	public void SubMethod_Result(boolean ErrorResult[], boolean result[]) {
		// 開啟Excel
		try {
			workBook = new XSSFWorkbook(new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm"));
		} catch (Exception e) {
			System.out.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm");
		}
		for (int i = 0; i < driver.length; i++) {

			if (TestCase.DeviceInformation.BrowserList.get(i).toString().length() > 20) {// Excel工作表名稱最常31字元因，故需判斷UDID長度是否大於31
				char[] NewUdid = new char[20];// 因需包含_TestReport字串(共11字元)，故設定20位字元陣列(31-11)
				TestCase.DeviceInformation.BrowserList.get(i).toString().getChars(0, 20, NewUdid, 0);// 取出UDID前20字元給NewUdid
				Sheet = workBook.getSheet(String.valueOf(NewUdid) + "_TestReport");// 根據NewUdid，指定某台裝置的TestReport
																					// sheet
			} else {
				Sheet = workBook.getSheet(TestCase.DeviceInformation.BrowserList.get(i).toString() + "_TestReport");// 指定某台裝置的TestReport
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
		// 執行寫入Excel後的存檔動作
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