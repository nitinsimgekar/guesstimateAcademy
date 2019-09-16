package com.guesstimate.web;

import java.util.Locale;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.safari.SafariDriver;

public class SeleniumDriverManager {

	private static String driversPath = "browserDrivers";

	public static WebDriver createBrowserInstance(String browser, String mainAddress) {

		String extension = "";
		String osName = System.getProperty("os.name").toLowerCase(Locale.ENGLISH);
		if (osName.contains("windows")) {
			extension = ".exe";
		}

		WebDriver driver = null;

		if (browser.equalsIgnoreCase("Firefox")) {
			System.setProperty("webdriver.gecko.driver", driversPath + "/geckodriver" + extension);
			driver = new FirefoxDriver();
		} else if (browser.equalsIgnoreCase("IE")) {
			System.setProperty("webdriver.ie.driver", driversPath + "/IEdriverServer.exe");
			driver = new InternetExplorerDriver();
		} else if (browser.equalsIgnoreCase("Safari")) {
			driver = new SafariDriver();
		} else {
			System.setProperty("webdriver.chrome.driver", driversPath + "/chromedriver" + extension);
			System.setProperty("download.default_directory", "Download");
			System.setProperty("requireWindowFocus", "true");
			System.setProperty("enablePersistentHover", "false");
			driver = new ChromeDriver();
		}

		driver.get(mainAddress);

		return driver;
	}
}
