package Auto;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;

import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.io.FileHandler;
import org.openqa.selenium.logging.LogEntries;
import org.openqa.selenium.logging.LogEntry;
import org.openqa.selenium.logging.LogType;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.google.common.base.Stopwatch;

public class method {
	static LoadTestCase TestCase = new LoadTestCase();
	LoadExpectResult ExpectResult = new LoadExpectResult();
	static String CaseErrorList[] = new String[TestCase.CaseList.size()];// 紀錄各案例於各裝置之指令結果
																			// (1維陣列)CaseErrorList[CaseList]

	static int command_timeout = 30;// 30sec
	static String appElemnt;// APP元件名稱
	static String appInput;// 輸入值
	static String appInputXpath;// 輸入值的Xpath格式
	static String element;
	static WebDriver driver;
	static int CurrentCaseNumber = -1;// 目前執行到第幾個測試案列
	static Boolean CommandError = true;// 判定執行的指令是否出現錯誤；ture為正確；false為錯誤
	static long totaltime;// 統計所有案例測試時間
	XSSFSheet Sheet;
	XSSFWorkbook workBook;
	static int CurrentCase;

	public static void main(String[] args) throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException, IOException {
		initial();
		invokeFunction();
		System.out.println("測試結束!!!" + "(" + totaltime + " s)");
		Runtime.getRuntime().exec("explorer C:\\TUTK_QA_TestTool\\TestReport");// 開啟TestReport資料夾
	}

	public static void initial() {// 初始化CaseErrorList矩陣
		for (int i = 0; i < CaseErrorList.length; i++) {
			CaseErrorList[i] = "";// 填入空字串，避免取值時，出現錯誤
		}
	}

	public static void invokeFunction() throws IllegalAccessException, IllegalArgumentException,
			InvocationTargetException, InstantiationException, NoSuchMethodException, SecurityException {

		Object obj = new method();
		Class c = obj.getClass();
		String methodName = null;

		for (CurrentCase = 0; CurrentCase < TestCase.StepList.size(); CurrentCase++) {
			CommandError = true;// 預設CommandError為True
			Stopwatch timer = Stopwatch.createStarted();// 開始計時
			System.out.println("[info] CaseName:|" + TestCase.CaseList.get(CurrentCase).toString() + "|");
			// ResultList = new ArrayList();
			// ResultList.add(TestCase.CaseList.get(CurrentCase).toString());
			for (int CurrentCaseStep = 0; CurrentCaseStep < TestCase.StepList.get(CurrentCase)
					.size(); CurrentCaseStep++) {
				if (!CommandError) {
					break;// 若目前測試案例出現CommandError=false，則跳出目前案例並執行下一個案例
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

				case "Byid_VerifyText":
					methodName = "Byid_VerifyText";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_VerifyText":
					methodName = "ByXpath_VerifyText";
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

				case "Byid_invisibility":
					methodName = "Byid_invisibility";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_invisibility":
					methodName = "ByXpath_invisibility";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Byid_Clear":
					methodName = "Byid_Clear";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Clear":
					methodName = "ByXpath_Clear";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ScreenShot":
					methodName = "ScreenShot";
					break;

				case "Back":
					methodName = "Back";
					break;

				case "Next":
					methodName = "Next";
					break;

				case "Refresh":
					methodName = "Refresh";
					break;

				case "Goto":
					methodName = "Goto";
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Sleep":
					methodName = "Sleep";
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Quit":
					methodName = "Quit";
					break;

				}

				Method method;
				method = c.getMethod(methodName, new Class[] {});
				method.invoke(c.newInstance());
			}
			System.out.println("[info] Time:|" + timer.stop() + "|");
			totaltime += timer.elapsed(TimeUnit.SECONDS);
			System.out.println("");
		}

	}

	public void ErrorCheck(Object... elements) throws IOException {
		DateFormat df = new SimpleDateFormat("MMM dd, yyyy h:mm:ss a");
		Date today = Calendar.getInstance().getTime();
		String reportDate = df.format(today);

		if (elements.length > 1) {

			String APPElement = "";
			int i = 0;
			for (Object element : elements) {
				APPElement = APPElement + element;
				if (i != (elements.length - 1)) {// 判斷是否處理到最後一個element
					APPElement = APPElement + " or ";// 組成" A元件 or B元件 or
														// C元件"字串
				}
				i++;
			}
			System.err.print("[Error] Can not found " + APPElement + " on screen.");
		} else {
			for (Object element : elements) {
				if (element.equals("Sleep")) {
					System.err.print("[Error] Fail to sleep.");
				} else if (element.equals("ScreenShot")) {
					System.err.print("[Error] Fail to ScreenShot.");
				} else if (element.equals("Quit")) {
					System.err.print("[Error] Can't close browser.");
				} else if (element.equals("Launch")) {
					System.err.print("[Error] Can't launch browser.");
				} else if (element.equals("BACK") || element.equals("Refresh") || element.equals("Next")) {
					System.err.print("[Error] Can't execute " + element + " button.");
				} else if (element.equals("Goto")) {
					System.err.println("[Error] Can't execute Goto " + appInput);
				} else if (element.equals("Customized")) {
					System.err.print("[Error] Can't execute Customized_Method.");
				} else {
					System.err.print("[Error] Can not found " + element + " on screen.");
				}
			}
		}
		System.err.println(" " + reportDate);
		String FilePath = MakeErrorFolder();// 建立各案例資料夾存放log資訊及Screenshot資訊
		ErrorScreenShot(FilePath);// Screenshot Error畫面
		// logcat(FilePath);// 收集閃退logcat
		CommandError = false;// 設定CommandError=false
		System.out.println("[info] " + TestCase.CaseList.get(CurrentCaseNumber).toString() + ":ERROR!");
	}

	// 目前Driver 無法取得log資訊
	public void logcat(String FilePath) throws IOException {
		// 收集log
		// System.out.println("[info] Saving device log...");
		DateFormat df = new SimpleDateFormat("yyyy_MM_dd_HH-mm-ss");
		Date today = Calendar.getInstance().getTime();
		String reportDate = df.format(today);

		LogEntries logEntries = driver.manage().logs().get(LogType.BROWSER);
		try {
			FileWriter fw = new FileWriter(FilePath + TestCase.CaseList.get(CurrentCaseNumber).toString() + "_"
					+ reportDate + "_log" + ".txt");
			for (LogEntry entry : logEntries) {
				fw.write("Timestamp:" + entry.getTimestamp() + "\r\n");
				fw.write("Level:" + entry.getLevel() + "\r\n");
				fw.write("Message:" + entry.getMessage() + "\r\n");
			}
			fw.flush();
			fw.close();
			System.out.println("[info] Saving browser log - Done.");
		} catch (Exception e) {
			System.err.println("[Error] Fail to save browser log.");
		}
	}

	public void ErrorScreenShot(String FilePath) {
		try {
			System.out.println("[info] Taking a screenshot of error.");
			DateFormat df = new SimpleDateFormat("yyyy_MM_dd_HH-mm-ss");
			Date today = Calendar.getInstance().getTime();
			String reportDate = df.format(today);
			File screenShotFile = (File) ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
			FileHandler.copy(screenShotFile,
					new File(FilePath + TestCase.CaseList.get(CurrentCaseNumber) + "_" + reportDate + ".jpg"));
		} catch (IOException e) {
			System.err.println("[Error] Fail to ErrorScreenShot.");
		}
	}

	public String MakeErrorFolder() {
		// 資料夾結構 C:\TUTK_QA_TestTool\TestReport\TestURL\CaseName\Browser\log\
		String filePath = "C:\\TUTK_QA_TestTool\\TestReport\\"
				+ TestCase.DeviceInformation.URL.replaceAll("//", "").replaceAll("https:", "").replaceAll("/", "")
						.replaceAll("http:", "").toString()
				+ "\\" + TestCase.CaseList.get(CurrentCase).toString() + "\\" + TestCase.DeviceInformation.Browser
				+ "\\log\\";
		File file = new File(filePath);
		if (!file.exists()) {
			file.mkdirs();
		}
		return filePath;
	}

	public void Byid_VerifyText() throws IOException {

		boolean result = false;// 未給定Boolean值，預設為False
		boolean ErrorResult = false;

		try {
			System.out.println("[info] Executing:|Byid_VerifyText|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(appElemnt))).getText();

		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}

		if (element.equals("ERROR")) {
			ErrorResult = true;

		} else {
			// 回傳測試案例清單的名稱給ExpectResult.LoadExpectResult，並存放期望結果至ResultList清單
			ExpectResult.LoadExpectResult(TestCase.CaseList.get(CurrentCaseNumber).toString());
			for (int j = 0; j < ExpectResult.ResultList.size(); j++) {
				if (element.equals(ExpectResult.ResultList.get(j)) == true) {
					result = true;
					break;
				} else {
					result = false;
				}
			}
		}
		SubMethod_Result(ErrorResult, result);
	}

	public void ByXpath_VerifyText() throws IOException {

		boolean result = false;// 未給定Boolean值，預設為False
		boolean ErrorResult = false;

		try {
			System.out.println("[info] Executing:|ByXpath_VerifyText|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).getText();

		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}

		if (element.equals("ERROR")) {
			ErrorResult = true;

		} else {
			// 回傳測試案例清單的名稱給ExpectResult.LoadExpectResult，並存放期望結果至ResultList清單
			ExpectResult.LoadExpectResult(TestCase.CaseList.get(CurrentCaseNumber).toString());
			for (int j = 0; j < ExpectResult.ResultList.size(); j++) {
				if (element.equals(ExpectResult.ResultList.get(j)) == true) {
					result = true;
					break;
				} else {
					result = false;
				}
			}
		}
		SubMethod_Result(ErrorResult, result);
	}

	public void Byid_Click() throws IOException {

		try {
			System.out.println("[info] Executing:|Byid_Click|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(appElemnt))).click();
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}
	}

	public void ByXpath_Click() throws IOException {

		try {
			System.out.println("[info] Executing:|ByXpath_Click|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).click();
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}
	}

	public void Byid_SendKey() throws IOException {

		try {
			System.out.println("[info] Executing:|Byid_SendKey|" + appElemnt + "|" + appInput + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(appElemnt))).sendKeys(appInput);
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}
	}

	public void ByXpath_SendKey() throws IOException {

		try {
			System.out.println("[info] Executing:|ByXpath_SendKey|" + appElemnt + "|" + appInput + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).sendKeys(appInput);
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}

	}

	public void ByXpath_Scroll() throws IOException {

		try {
			System.out.println("[info] Executing:|ByXpath_Scroll|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			WebElement target = wait.until(ExpectedConditions.presenceOfElementLocated((By.xpath(appElemnt))));
			Actions actions = new Actions(driver);
			actions.moveToElement(target);
			// actions.click(target);
			actions.perform();
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}
	}

	public void Byid_Scroll() throws IOException {

		try {
			System.out.println("[info] Executing:|Byid_Scroll|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			WebElement target = wait.until(ExpectedConditions.presenceOfElementLocated((By.id(appElemnt))));
			Actions actions = new Actions(driver);
			actions.moveToElement(target);
			// actions.click(target);
			actions.perform();
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}

	}

	public void ByXpath_invisibility() throws IOException {

		try {
			System.out.println("[info] Executing:|ByXpath_invisibility|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(appElemnt)));
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}
	}

	public void Byid_invisibility() throws IOException {

		try {
			System.out.println("[info] Executing:|Byid_invisibility|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.id(appElemnt)));
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}
	}

	public void Byid_Wait() throws IOException {

		try {
			System.out.println("[info] Executing:|Byid_Wait|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			wait.until(ExpectedConditions.presenceOfElementLocated(By.id(appElemnt)));
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}

	}

	public void ByXpath_Wait() throws IOException {

		try {
			System.out.println("[info] Executing:|ByXpath_Wait|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(appElemnt)));
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}
	}

	public void Byid_Clear() throws IOException {

		try {
			System.out.println("[info] Executing:|Byid_Clear|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(appElemnt))).clear();
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}
	}

	public void ByXpath_Clear() throws IOException {

		try {
			System.out.println("[info] Executing:|ByXpath_Clear|" + appElemnt + "|");
			WebDriverWait wait = new WebDriverWait(driver, command_timeout);
			wait = new WebDriverWait(driver, command_timeout);
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).clear();
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck(appElemnt);
		}
	}

	public void Launch() throws IOException {
		CurrentCaseNumber = CurrentCaseNumber + 1;
		System.out.println("[info] Executing:|Launch Browser|" + TestCase.DeviceInformation.Browser + "|"
				+ TestCase.DeviceInformation.URL + "|");

		try {
			// 僅支援Chrome與FireFox
			switch (TestCase.DeviceInformation.Browser) {
			case "chrome":
				System.setProperty("webdriver.chrome.driver", TestCase.DeviceInformation.DriverPath.toString());
				driver = new ChromeDriver();
				break;
			case "firefox":
				System.setProperty("webdriver.gecko.driver", TestCase.DeviceInformation.DriverPath.toString());
				driver = new FirefoxDriver();
				break;
			}
			driver.manage().timeouts().pageLoadTimeout(command_timeout, TimeUnit.SECONDS);
			driver.manage().window().maximize();
			driver.get(TestCase.DeviceInformation.URL);

		} catch (Exception e1) {
			ErrorCheck("Launch");
		}
	}

	public void Quit() throws IOException {
		boolean state = false;
		try {

			System.out.println("[info] Executing:|Quit Browser|");
			driver.quit();// 離開APP後，寫入測試結果Pass或Error

			for (int i = 0; i < TestCase.StepList.get(CurrentCaseNumber).size(); i++) {
				if (TestCase.StepList.get(CurrentCaseNumber).get(i).equals("Byid_VerifyText")
						|| TestCase.StepList.get(CurrentCaseNumber).get(i).equals("ByXpath_VerifyText")) {
					state = true;// true代表找到Verify
					break;
				}
			}

			if (!state) {
				// 開啟Excel
				try {
					workBook = new XSSFWorkbook(
							new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm"));
				} catch (Exception e) {
					System.out.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm");
				}

				if (TestCase.DeviceInformation.Browser.toString().length() > 20) {// Excel工作表名稱最常31字元因，故需判斷BrowserName長度是否大於31
					char[] NewBrowserName = new char[20];// 因需包含_TestReport字串(共11字元)，故設定20位字元陣列(31-11)
					TestCase.DeviceInformation.Browser.toString().getChars(0, 20, NewBrowserName, 0);// 取出BrowserName前20字元給NewBrowserName
					Sheet = workBook.getSheet(String.valueOf(NewBrowserName) + "_TestReport");// 根據NewUdid，指定某台裝置的TestReport
					// sheet
				} else {
					Sheet = workBook.getSheet(TestCase.DeviceInformation.Browser.toString() + "_TestReport");// 指定某台裝置的TestReport
																												// sheet
				}

				if (CaseErrorList[CurrentCaseNumber].equals("Pass")) {// 取出CaseErrorList之第CurrentCaseNumber個測項中的第i台行動裝置之結果
					Sheet.getRow(CurrentCaseNumber + 1).getCell(1).setCellValue("Pass");// 填入第i台行動裝置之第CurrentCaseNumber個測項結果Pass
				}

				// 執行寫入Excel後的存檔動作
				try {
					FileOutputStream out = new FileOutputStream(
							new File("C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm"));
					workBook.write(out);
					out.close();
					workBook.close();
				} catch (Exception e) {
					System.err.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm");
				}
			}
		} catch (Exception e) {
			ErrorCheck("Quit");
		}
	}

	public void Sleep() throws IOException {
		try {
			System.out.println("[info] Executing:|Sleep|" + appInput + " second..." + "|");
			Thread.sleep((long) (Float.valueOf(appInput) * 1000));
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck("Sleep");
		}
	}

	public void ScreenShot() throws IOException {

		try {

			String filePath = "C:\\TUTK_QA_TestTool\\TestReport\\"
					+ TestCase.DeviceInformation.URL.replaceAll("//", "").replaceAll("https:", "").replaceAll("/", "")
							.replaceAll("http:", "").toString()
					+ "\\" + TestCase.CaseList.get(CurrentCase).toString() + "\\" + TestCase.DeviceInformation.Browser
					+ "\\ScreenShot\\";
			File file = new File(filePath);
			if (!file.exists()) {
				file.mkdirs();
			}

			DateFormat df = new SimpleDateFormat("yyyy_MM_dd_HH-mm-ss");
			Date today = Calendar.getInstance().getTime();
			String reportDate = df.format(today);

			File screenShotFile = (File) ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
			System.out.println("[info] Executing:|ScreenShot|");
			FileHandler.copy(screenShotFile,
					new File(filePath + TestCase.CaseList.get(CurrentCaseNumber) + "_" + reportDate + ".jpg"));
			System.out.println("[info] Executing:|ScreenShot|");
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (IOException e) {
			ErrorCheck("ScreenShot");
		}

	}

	public void Back() throws IOException {

		try {
			System.out.println("[info] Executing:|Back|");
			driver.navigate().back();
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck("Back");
		}

	}

	public void Next() throws IOException {

		try {
			System.out.println("[info] Executing:|Next|");
			driver.navigate().forward();
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck("Next");
		}

	}

	public void Refresh() throws IOException {

		try {
			System.out.println("[info] Executing:|Refresh|");
			driver.navigate().refresh();
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck("Refresh");
		}
	}

	public void Goto() throws IOException {

		try {
			System.out.println("[info] Executing:|Goto|" + appInput + "|");
			driver.navigate().to(appInput);
			CaseErrorList[CurrentCaseNumber] = "Pass";
		} catch (Exception e) {
			ErrorCheck("Goto");
		}
	}

	public void SubMethod_Result(boolean ErrorResult, boolean result) {
		// 開啟Excel
		try {
			workBook = new XSSFWorkbook(new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm"));
		} catch (Exception e) {
			System.err.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm");
		}

		if (TestCase.DeviceInformation.Browser.toString().length() > 20) {// Excel工作表名稱最常31字元因，故需判斷UDID長度是否大於31
			char[] NewUdid = new char[20];// 因需包含_TestReport字串(共11字元)，故設定20位字元陣列(31-11)
			TestCase.DeviceInformation.Browser.toString().getChars(0, 20, NewUdid, 0);// 取出UDID前20字元給NewUdid
			Sheet = workBook.getSheet(String.valueOf(NewUdid) + "_TestReport");// 根據NewUdid，指定某台裝置的TestReport
																				// sheet
		} else {
			Sheet = workBook.getSheet(TestCase.DeviceInformation.Browser.toString() + "_TestReport");// 指定某台裝置的TestReport
																										// sheet
		}

		if (ErrorResult == true) {
			Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Error");
		} else if (result == true) {
			Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Pass");
		} else if (result == false) {
			Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Fail");
		}

		// 執行寫入Excel後的存檔動作
		try {
			FileOutputStream out = new FileOutputStream(
					new File("C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm"));
			workBook.write(out);
			out.close();
			workBook.close();
		} catch (Exception e) {
			System.err.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm");
		}
	}
}