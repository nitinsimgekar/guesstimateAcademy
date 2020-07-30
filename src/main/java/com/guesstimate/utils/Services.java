package com.guesstimate.utils;

import io.appium.java_client.AppiumDriver;
import io.appium.java_client.MobileElement;
import io.appium.java_client.TouchAction;
import io.appium.java_client.touch.WaitOptions;
import io.appium.java_client.touch.offset.PointOption;
import org.openqa.selenium.WebDriverException;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.Instant;
import java.time.ZoneId;
import java.util.Date;
import java.util.HashMap;
import java.util.TimeZone;

public class Services {

    public void hideKeyboard(AppiumDriver<?> driver) {
        try {
            driver.hideKeyboard();
            driver.findElementByAccessibilityId("Toolbar Done Button").click();
        } catch (WebDriverException e) {
        }
    }

    public void hideKeyboard(AppiumDriver<?> apDriver, DeviceType deviceType) {
        try {
            if (deviceType == DeviceType.IOS) {
                apDriver.findElementByAccessibilityId("Toolbar Done Button").click();
            } else {
                apDriver.hideKeyboard();
            }
        } catch (WebDriverException e) {
        }
    }

    public void hideKeyboard(AppiumDriver<?> apDriver, DeviceType deviceType, String element) {
        try {
            if (deviceType == DeviceType.IOS) {
                apDriver.findElementByAccessibilityId(element).click();
            } else {
                apDriver.hideKeyboard();
            }
        } catch (WebDriverException e) {
        }
    }

    public void scrollRight(AppiumDriver<?> apDriver, MobileElement element) {

        int leftXOffSet = (int) (element.getSize().getWidth() * 0.20);
        int rightX = (int) (element.getSize().getWidth() * 0.80);

        int middleY = (element.getSize().getHeight() / 2) + element.getLocation().getY();

        TouchAction swipe = new TouchAction(apDriver).press(PointOption.point(rightX, middleY))
                .waitAction(WaitOptions.waitOptions(Duration.ofMillis(1000)))
                .moveTo(PointOption.point(leftXOffSet, middleY)).release();
        swipe.perform();

    }

    public void scrollUp(AppiumDriver<?> apDriver, MobileElement element) {

        int upperY = (int) (element.getSize().getHeight() * 0.20);
        int lowerY = (int) (element.getSize().getHeight() * 0.80);

        int middleX = element.getSize().getWidth() / 2;

        TouchAction swipe = new TouchAction(apDriver).press(PointOption.point(middleX, lowerY))
                .waitAction(WaitOptions.waitOptions(Duration.ofMillis(1000))).moveTo(PointOption.point(middleX, upperY))
                .release();
        swipe.perform();

    }

    public void scrollDown(AppiumDriver<?> apDriver, MobileElement element) {

        int upperY = (int) (element.getSize().getHeight() * 0.20);
        int lowerY = (int) (element.getSize().getHeight() * 0.80);

        int middleX = element.getSize().getWidth() / 2;

        TouchAction swipe = new TouchAction(apDriver).press(PointOption.point(middleX, upperY))
                .waitAction(WaitOptions.waitOptions(Duration.ofMillis(1000))).moveTo(PointOption.point(middleX, lowerY))
                .release();
        swipe.perform();

    }

    public void scrollDownUsingJavaScript(AppiumDriver<?> apDriver, MobileElement element) {

        String elementID = element.getId();
        HashMap<String, String> scrollObject = new HashMap<String, String>();
        scrollObject.put("element", elementID);
        scrollObject.put("direction", "down");
        apDriver.executeScript("mobile: scroll", scrollObject);

    }

    public void scrollUpUsingJavaScript(AppiumDriver<?> apDriver, MobileElement element) {

        String elementID = element.getId();
        HashMap<String, String> scrollObject = new HashMap<String, String>();
        scrollObject.put("element", elementID);
        scrollObject.put("direction", "up");
        apDriver.executeScript("mobile: scroll", scrollObject);

    }

    public String formatDateToString(Date date, String format, String timeZone) {

        DateFormat df = new SimpleDateFormat(format);
        df.setTimeZone(TimeZone.getTimeZone(timeZone));
        return df.format(date);

    }

    //ZoneId systemDate = ZoneId.systemDefault();
    //String actual = "2018-11-02T07:30:00Z";
    public boolean isDateRecordedNearBy(long systemEpochMili, String couchBaseDate) throws ParseException {

        long lR = Instant.parse(couchBaseDate).atZone(ZoneId.of("UTC")).minusMinutes(3).toInstant().toEpochMilli();
        long hR = Instant.parse(couchBaseDate).atZone(ZoneId.of("UTC")).plusMinutes(3).toInstant().toEpochMilli();

        boolean isMatch = false;

        if (hR > systemEpochMili && systemEpochMili > lR)
            isMatch = true;

        return isMatch;

    }

    public boolean isEitherOf(String actual, String expected) {

        boolean isMatch = false;

        String[] exp = expected.split(",");

        for (String each : exp) {

            isMatch = actual.equalsIgnoreCase(each.trim());
            if (isMatch) {
                return isMatch;
            }
        }


        return isMatch;
    }

    public String readFile(String filepath) {

        StringBuilder stringBuilder = new StringBuilder();

        try {
            BufferedReader reader = new BufferedReader(new FileReader(filepath));

            String line = null;
            String ls = System.getProperty("line.separator");
            while ((line = reader.readLine()) != null) {
                stringBuilder.append(line.trim());
                stringBuilder.append(ls.trim());
            }

            reader.close();

        } catch (IOException e) {

        }

        return stringBuilder.toString();
    }

    public String readFileWithSpaces(String filepath) {

        StringBuilder stringBuilder = new StringBuilder();

        try {
            BufferedReader reader = new BufferedReader(new FileReader(filepath));

            String line = null;
            while ((line = reader.readLine()) != null) {
                stringBuilder.append(line + "\n");
            }

            reader.close();

        } catch (IOException e) {

        }

        return stringBuilder.toString();
    }

}
