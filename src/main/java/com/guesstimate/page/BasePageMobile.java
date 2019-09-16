package com.guesstimate.page;

import java.util.HashMap;

import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import io.appium.java_client.pagefactory.AppiumFieldDecorator;

public abstract class BasePageMobile {

	public AppiumDriver<?> driver;
	public WebDriverWait wait;

	public BasePageMobile(AppiumDriver<?> driver) {
		this.driver = driver;
		this.initPage();
		wait = new WebDriverWait(driver, 7);
	}

	public void initPage() {
		PageFactory.initElements(new AppiumFieldDecorator(driver), this);

	}

	public void sleepFor(long secounds) {
		try {
			Thread.sleep(secounds * 1000);
		} catch (InterruptedException e) {
			System.out.println("ERROR: execption during sleep.");
		}

	}

	public void clickUsingJavaScript(MobileElement element) {
		String elementID = element.getId();
		HashMap<String, String> clickObject = new HashMap<String, String>();
		clickObject.put("element", elementID);
		clickObject.put("x", String.valueOf(element.getSize().getWidth() / 2));
		clickObject.put("y", String.valueOf(element.getSize().getHeight() / 2));
		driver.executeScript("mobile:tap", clickObject);

	}

}
