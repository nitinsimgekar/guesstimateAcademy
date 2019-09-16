package com.guesstimate.mobile;

import java.net.MalformedURLException;
import java.net.URL;

import org.openqa.selenium.remote.DesiredCapabilities;

import io.appium.java_client.AppiumDriver;

public class Android {

	public AppiumDriver<?> forAndroid(DeviceProp emulator) throws MalformedURLException {

		System.out.println("Creating Session for '" + emulator.getEmulatorDeviceName() + "' with port '"
				+ emulator.getPort() + "'");

		DesiredCapabilities capabilities = new DesiredCapabilities();

		capabilities.setCapability("deviceName", emulator.getEmulatorDeviceName());
		capabilities.setCapability("appPackage", emulator.getAppEnv().getAppPackage());
		capabilities.setCapability("appActivity", emulator.getAppEnv().getAppActivity());
		capabilities.setCapability("platformName", "Android");
		capabilities.setCapability("noReset", false);

		AppiumDriver<?> driver = new AppiumDriver<>(new URL("http://127.0.0.1:" + emulator.getPort() + "/wd/hub"),
				capabilities);

		return driver;
	}

	public AppiumDriver<?> settings(DeviceProp emulator) throws MalformedURLException {

		System.out.println("Creating Session for '" + emulator.getEmulatorDeviceName() + "' with port '"
				+ emulator.getPort() + "'");

		DesiredCapabilities capabilities = new DesiredCapabilities();

		capabilities.setCapability("deviceName", emulator.getEmulatorDeviceName());
		capabilities.setCapability("appPackage", emulator.getAppEnv().getAppPackage());
		capabilities.setCapability("appActivity", emulator.getAppEnv().getAppActivity());
		capabilities.setCapability("platformName", "Android");
		capabilities.setCapability("noReset", false);

		AppiumDriver<?> driver = new AppiumDriver<>(new URL("http://127.0.0.1:" + emulator.getPort() + "/wd/hub"),
				capabilities);

		return driver;
	}

}
