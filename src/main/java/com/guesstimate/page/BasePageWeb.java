package com.guesstimate.page;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.WebDriverWait;

public class BasePageWeb {

	public WebDriver driver;
	public WebDriverWait wait;
	public JavascriptExecutor executor;

	public BasePageWeb(WebDriver driver) {
		this.driver = driver;
		this.initPage();
		wait = new WebDriverWait(driver, 50);
		executor = (JavascriptExecutor) driver;
	}

	public void initPage() {
		PageFactory.initElements(this.driver, this);

	}

	public void clickUsingJavaScript(WebElement webElement) {
		executor.executeScript("arguments[0].click();", webElement);

	}

	public void sleepFor(long secounds) {
		try {
			Thread.sleep(secounds * 1000);
		} catch (InterruptedException e) {
			System.out.println("ERROR: execption during sleep.");
		}

	}

	public void scrollUsingJavaScript(WebElement webElement) {
		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", webElement);

	}

}