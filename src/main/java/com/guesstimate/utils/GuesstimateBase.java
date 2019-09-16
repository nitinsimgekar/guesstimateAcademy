package com.guesstimate.utils;

import com.guesstimate.mobile.AppEnv;
import com.guesstimate.mobile.DeviceProp;
import com.guesstimate.mobile.Initialize;
import com.guesstimate.web.SeleniumDriverManager;
import io.appium.java_client.AppiumDriver;
import org.ini4j.Wini;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;
import org.testng.asserts.SoftAssert;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class GuesstimateBase {

    public static String name, webUsername, webPassword, mobileUsername, mobilePassword, baseUrl, browser, server, buildNumber, jiraHost, jenkinsBaseUrl;

    public static WebDriver sedriver;
    public static AppiumDriver<?> apdriver;

    public static boolean logJira;

    public static ArrayList<DeviceProp> deviceProps;

    public Initialize initialize = new Initialize();
    public SoftAssert softAssert = new SoftAssert();
    public Services services = new Services();
    public static AppEnv appEnv = new AppEnv();


    @Parameters({"server", "webUsername", "webPassword", "mobileUsername", "mobilePassword", "browser", "baseUrl", "logJira", "jiraHost", "jenkinsBaseUrl", "projectName", "oprBundleId", "appActivity", "appPackage"})
    @BeforeSuite(alwaysRun = true)
    public void setUp(@Optional String server, @Optional String webUsername, @Optional String webPassword,
                      @Optional String mobileUsername, @Optional String mobilePassword,
                      @Optional String browser, @Optional String baseUrl,
                      @Optional String logJira, @Optional String jiraHost, @Optional String jenkinsBaseUrl, @Optional String projectName, @Optional String oprBundleId, @Optional String appActivity, @Optional String appPackage) throws IOException {

        String operatorbuildNbrIos = (System.getProperty("oprI") == null) ? "" : System.getProperty("oprI");
        String operatorbuildNbrAndroid = (System.getProperty("oprA") == null) ? "" : System.getProperty("oprA");

        String supervisorbuildNbrIos = (System.getProperty("suprI") == null) ? "" : System.getProperty("suprI");
        String supervisorbuildNbrAndroid = (System.getProperty("suprA") == null) ? "" : System.getProperty("suprA");

        String device = (System.getProperty("device") == null) ? "ALL" : System.getProperty("device");

        GuesstimateBase.buildNumber = (System.getProperty("buildNumber") == null) ? "11111" : System.getProperty("buildNumber");

        GuesstimateBase.logJira = Boolean.parseBoolean((System.getProperty("logJira") == null) ? logJira : System.getProperty("logJira"));
        GuesstimateBase.jiraHost = jiraHost;
        GuesstimateBase.jenkinsBaseUrl = jenkinsBaseUrl;

        GuesstimateBase.name = projectName + "_" + server.toUpperCase();
        GuesstimateBase.server = server;
        GuesstimateBase.webUsername = webUsername;
        GuesstimateBase.webPassword = webPassword;
        GuesstimateBase.mobileUsername = mobileUsername;
        GuesstimateBase.mobilePassword = mobilePassword;
        GuesstimateBase.browser = browser;
        GuesstimateBase.baseUrl = baseUrl;

        appEnv.setName(projectName).setServer(server).setOprBundleId(oprBundleId).setAppActivity(appActivity).setAppPackage(appPackage)
                .setDeviceType(DeviceType.IOS).setOprHockeyIosPath("")
                .setOperatorbuildNbrIos(server + "_O_" + operatorbuildNbrIos).setOperatorbuildNbrAndroid(server + "_O_" + operatorbuildNbrAndroid);

        sedriver = SeleniumDriverManager.createBrowserInstance(browser, baseUrl);

        deviceProps = initialize.bringEnvironmentUp(DeviceEnvironment.valueOf(device), appEnv, false);

        int appiumPort = 4725;

        List<DeviceProp> temp = new ArrayList<>();
        boolean isRemoved = false;

        for (DeviceProp deviceProp : deviceProps) {
            deviceProp.setPort(appiumPort++);

            try {

                apdriver = initialize.startApplication(deviceProp);
                apdriver.resetApp();

            } catch (NullPointerException e) {
                temp.add(deviceProp);
                isRemoved = true;
            }
        }

        deviceProps.removeAll(temp);

        if (isRemoved) {

            System.out.println(
                    "Your test cases will now be executed on the following environments, as the other enviroments had problem with communicating with appium:");
            for (DeviceProp deviceProp : deviceProps) {
                System.out.println("'" + deviceProp.getName() + "' with version '" + deviceProp.getVersion() + "'");

            }
        }

    }


    @AfterSuite(alwaysRun = true)
    public void killEmulators() {

        Wini ini = null;

        try {
            ini = new Wini(new File("./res/runningEmulators.ini"));

            for (DeviceProp deviceProp : deviceProps) {
                if (!deviceProp.isRealDevice()) {

                    System.out.println("Bringing the emulator '" + deviceProp.getName() + "' down. Please wait...");

                    BufferedReader r = null;

                    switch (deviceProp.getDeviceType()) {
                        case ANDROID:
                            Process p1 = Runtime.getRuntime()
                                    .exec("adb -s " + deviceProp.getEmulatorDeviceName() + " emu kill");
                            r = new BufferedReader(new InputStreamReader(p1.getInputStream()));
                            break;

                        case IOS:
                            Process p2 = Runtime.getRuntime().exec("killall Simulator");
                            r = new BufferedReader(new InputStreamReader(p2.getInputStream()));
                            break;

                    }

                    while (true) {

                        String emulatorline = r.readLine();
                        if (emulatorline == null) {
                            ini.put(deviceProp.getName(), "isEmulatorUp", false);
                            ini.store();
                            break;
                        }
                    }

                }

            }

        } catch (IOException | NullPointerException e1) {
        }

        if (sedriver != null) {
            sedriver.quit();
            if (sedriver != null) {
                sedriver = null;
            }
        }

    }

    public void resetAfterClassWeb() {
        if (sedriver != null) {
            sedriver.quit();
            if (sedriver != null) {
                sedriver = null;
            }
        }
    }

    public void beforeClassResetMobile() throws MalformedURLException {

        for (DeviceProp deviceProp : deviceProps) {

            apdriver = initialize.startApplication(deviceProp);

            if (deviceProp.getDeviceType() == DeviceType.IOS)
                initialize.installIPA(deviceProp);
            else
                apdriver.resetApp();

        }

    }

    public WebDriver getSeDriver() {

        return sedriver;
    }

    public WebDriver getApDriver() {

        return apdriver;
    }

    public boolean isLogJira() {
        return logJira;
    }

    public String getBuildNumber() {
        return buildNumber;
    }

    public String getEnvironmnet() {

        if (server.equalsIgnoreCase("dev"))
            return "Development";
        else
            return "Test";
    }

}
