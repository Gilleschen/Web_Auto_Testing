package AutoTesting;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.safari.SafariDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class method2 {

	LoadTestCase TestCase = new LoadTestCase();
	WebDriver driver = null;
	WebDriverWait wait = null;
	static String appElemnt;// APP元件名稱
	static String appInput;// 輸入值
	public void Launch() {

		switch (TestCase.DeviceInformation.Browser.toString()) {

		case "Firefox":
			System.setProperty("webdriver.gecko.driver", TestCase.DeviceInformation.DriverPath);
			driver = new FirefoxDriver();
			break;
		case "IE":
			System.setProperty("webdriver.ie.driver", TestCase.DeviceInformation.DriverPath);
			driver = new InternetExplorerDriver();
			break;
		case "Chrome":
			System.setProperty("webdriver.chrome.driver", TestCase.DeviceInformation.DriverPath);
			driver = new ChromeDriver();
			break;
		case "Safari":
			System.setProperty("webdriver.safari.driver", TestCase.DeviceInformation.DriverPath);
			driver = new SafariDriver();
			break;
		}
		;

		wait = new WebDriverWait(driver, 15);
		driver.get(TestCase.DeviceInformation.URL);
		try {
			Thread.sleep(5000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		driver.quit();
	}

	

	public static void main(String[] args) {
		method2 m = new method2();
		m.Launch();
	}

}
