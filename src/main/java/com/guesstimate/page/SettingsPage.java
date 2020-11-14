package com.guesstimate.page;

import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import io.appium.java_client.pagefactory.AndroidFindBy;
import io.appium.java_client.pagefactory.iOSFindBy;
import org.openqa.selenium.support.ui.ExpectedConditions;

import java.util.List;

public class SettingsPage extends BasePageMobile {

    public SettingsPage(AppiumDriver<?> driver) {
        super(driver);
    }

    @iOSFindBy(xpath = "//XCUIElementTypeCell[@name='Wi-Fi']")
    @AndroidFindBy(xpath = "//*[text()='Connections']")
    private MobileElement wifiOption;

    @AndroidFindBy(id = "switch_widget")
    private MobileElement switchWidget;

    @iOSFindBy(xpath = "//*[@label='Wi-Fi']")
    private List<MobileElement> labels;

    @iOSFindBy(xpath = "//XCUIElementTypeSwitch[@name='Wi-Fi']")
    @AndroidFindBy(xpath = "//*[@id='list']/android.widget.LinearLayout/android.widget.LinearLayout")
    private MobileElement wifiSwitch;

    public SettingsPage clickWifiOption() {
        wait.until(ExpectedConditions.visibilityOf(wifiOption));
        wifiOption.click();
        return this;
    }

    public SettingsPage clickToggleBtnAndroid() {
        wait.until(ExpectedConditions.visibilityOf(switchWidget));
        switchWidget.click();
        return this;
    }

    public SettingsPage turnOnWifi() {
        sleepFor(2);
        wifiSwitch.click();
        return this;
    }

    public SettingsPage turnOffWifiAndroid() {
        wait.until(ExpectedConditions.visibilityOf(wifiSwitch));
        wifiSwitch.click();
        return this;
    }

    public SettingsPage turnOffWifiIos() {
        sleepFor(2);
        for (MobileElement label :labels) {
            label.click();
        }
        return this;
    }

}
