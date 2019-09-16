package com.guesstimate.mobile;

import java.net.MalformedURLException;
import java.net.URL;

import org.openqa.selenium.remote.DesiredCapabilities;

import io.appium.java_client.AppiumDriver;

public class Apple {

    public AppiumDriver<?> forIOS(DeviceProp emulator) throws MalformedURLException {

        System.out.println("Creating Session for '" + emulator.getName() + "' with port '" + emulator.getPort() + "'");

        DesiredCapabilities capabilities = new DesiredCapabilities();

        capabilities.setCapability("platformName", "ios");
        capabilities.setCapability("platformVersion", emulator.getVersion());
        capabilities.setCapability("deviceName", emulator.getName());
        capabilities.setCapability("udid", emulator.getEmulatorDeviceName());
        capabilities.setCapability("bundleId", emulator.getAppEnv().getOprBundleId());
        capabilities.setCapability("automationName", "XCUITest");
        capabilities.setCapability("fullReset", false);
        capabilities.setCapability("noReset", true);
        capabilities.setCapability("newCommandTimeout", 120);

        // for barric
        //capabilities.setCapability("xcodeOrgId", "F4G6G6KK28");
        //capabilities.setCapability("xcodeSigningId", "iPhone Developer");
        AppiumDriver<?> driver = new AppiumDriver<>(new URL("http://127.0.0.1:" + emulator.getPort() + "/wd/hub"),
                capabilities);

        return driver;

    }

    public AppiumDriver<?> settings(DeviceProp emulator) throws MalformedURLException {

        System.out.println("Creating Session for '" + emulator.getName() + "' with port '" + emulator.getPort() + "'");

        DesiredCapabilities capabilities = new DesiredCapabilities();

        capabilities.setCapability("platformName", "ios");
        capabilities.setCapability("platformVersion", emulator.getVersion());
        capabilities.setCapability("deviceName", emulator.getName());
        capabilities.setCapability("udid", emulator.getEmulatorDeviceName());
        capabilities.setCapability("bundleId", "com.apple.Preferences");
        capabilities.setCapability("automationName", "XCUITest");
        capabilities.setCapability("newCommandTimeout", 120);

        AppiumDriver<?> driver = new AppiumDriver<>(new URL("http://127.0.0.1:" + emulator.getPort() + "/wd/hub"),
                capabilities);

        return driver;

    }

}
