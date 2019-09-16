package com.guesstimate.Listeners;

import com.guesstimate.ExtentReports.ExtentManager;
import com.guesstimate.ExtentReports.ExtentTestManager;
import com.guesstimate.jira.JiraActivities;
import com.guesstimate.utils.GuesstimateBase;
import com.relevantcodes.extentreports.LogStatus;
import org.ini4j.Profile;
import org.ini4j.Wini;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.testng.ITestContext;
import org.testng.ITestListener;
import org.testng.ITestResult;

import java.io.*;
import java.util.HashSet;
import java.util.Set;

public class TestListener extends GuesstimateBase implements ITestListener {

    private static String getTestMethodName(ITestResult iTestResult) {
        return iTestResult.getMethod().getConstructorOrMethod().getName();
    }

    private static String getTestClassName(ITestResult iTestResult) {
        return iTestResult.getClass().getName();
    }

    @Override
    public void onStart(ITestContext iTestContext) {
        iTestContext.setAttribute("WebDriver", sedriver);
    }

    @Override
    public void onFinish(ITestContext iTestContext) {
        ExtentTestManager.endTest();
        ExtentManager.getReporter().flush();
    }

    @Override
    public void onTestStart(ITestResult iTestResult) {
        System.out.println("\nTest case Started: " + getTestMethodName(iTestResult));
        String packName = (iTestResult.getTestClass().getName() + " -- " + iTestResult.getMethod().getMethodName());
        ExtentTestManager.startTest(packName.substring(packName.indexOf("testcases."), packName.length()), "");
    }

    @Override
    public void onTestSuccess(ITestResult iTestResult) {
        if (getTestClassName(iTestResult).startsWith("w"))
            ExtentTestManager.getTest().log(LogStatus.PASS, "Wow !!! Test passed.");
    }

    @Override
    public void onTestFailure(ITestResult iTestResult) {

        Object testClass = iTestResult.getInstance();

        WebDriver webDriver;

        String desc = ExtentTestManager.getTest().getDescription();
        if (desc.equalsIgnoreCase("mobile"))
            webDriver = ((GuesstimateBase) testClass).getApDriver();
        else
            webDriver = ((GuesstimateBase) testClass).getSeDriver();

        String base64Screenshot = "data:image/png;base64,"
                + ((TakesScreenshot) webDriver).getScreenshotAs(OutputType.BASE64);

        File screenPrintJira = ((TakesScreenshot) webDriver).getScreenshotAs(OutputType.FILE);

        String trace;

        StringWriter sw = new StringWriter();
        PrintWriter pw = new PrintWriter(sw);
        Throwable cause = iTestResult.getThrowable();

        if (null != cause) {

            if (cause.getClass().getName().equalsIgnoreCase("java.lang.AssertionError")) {
                trace = cause.getMessage();

                if (isLogJira()) {

                    Wini ini = null;
                    String loggedBugRef = "";


                    JiraActivities jiraActivities = new JiraActivities();

                    try {
                        ini = new Wini(new File("./res/jira/reportedBugs.ini"));

                        Set<String> loggedBugs = new HashSet<>();
                        Set<String> sections = ini.keySet();

                        for (String sec : sections) {

                            Profile.Section section = ini.get(sec);


                            if (section != null) {
                                loggedBugs = section.keySet();
                            }
                        }

                        String[] bugs = trace.split("\n");

                        String barrerAuthToken = "";


                        for (String bug : bugs) {

                            try {

                                if (!bug.trim().equalsIgnoreCase("null") && !bug.trim().isEmpty() && !bug.trim().equalsIgnoreCase("The following asserts failed:") && !bug.trim().equalsIgnoreCase("This feature can be tested only on real device. emulators support is not yet provided.")) {

                                    String summary = "Data validation: " + iTestResult.getMethod().getMethodName();

                                    if (bug.trim().endsWith(",")) {
                                        loggedBugRef = bug.trim();
                                        loggedBugRef = loggedBugRef.substring(0, bug.trim().length() - 1);
                                    } else
                                        loggedBugRef = bug.trim();

                                    if (bug.contains("doesn't match")) {
                                        summary = summary + ": " + bug.substring(0, bug.indexOf("doesn't")).trim();
                                    }


                                    if (!loggedBugs.contains(loggedBugRef)) {


                                        String description = iTestResult.getTestClass().getName() + " -- " + iTestResult.getMethod().getMethodName() + "\\n\\n" + bug;
                                        String body = services.readFile("./res/jira/createBug.json").replace("{{ Environment }}", getEnvironmnet()).replace("\"{{ buildNo }}\"", getBuildNumber()).replace("{{ summary }}", summary).replace("{{ description }}", description);

                                        String issueId = jiraActivities.createBug(jiraHost, barrerAuthToken, body);

                                        ini.put(iTestResult.getTestClass().getName() + "." + iTestResult.getMethod().getMethodName(), loggedBugRef, issueId);
                                        ini.store();

                                        jiraActivities.addAttachmentToIssue(jiraHost, barrerAuthToken, issueId, screenPrintJira);


                                    } else {

                                        String issueId = ini.get(iTestResult.getTestClass().getName() + "." + iTestResult.getMethod().getMethodName(), loggedBugRef);


                                        String status = jiraActivities.getStatus(jiraHost, barrerAuthToken, issueId);

                                        if (status.equalsIgnoreCase("Done")) {

                                            if (!jiraActivities.reopenBug(jiraHost, barrerAuthToken, issueId, "281")) {

                                                System.out.println("bugsFailedToReopen " + jiraActivities.getKey(jiraHost, barrerAuthToken, issueId));
                                                BufferedWriter writer = new BufferedWriter(new FileWriter("./res/jira/bugsFailedToReopen.text", true));
                                                writer.write("\n" + jiraActivities.getKey(jiraHost, barrerAuthToken, issueId));
                                                writer.close();
                                            }

                                        }

                                        String comment = "This issue was also observed in buildNo: " + getBuildNumber() + "\\nReopening the issue.";

                                        jiraActivities.addComments(jiraHost, barrerAuthToken, issueId, comment);


                                    }
                                }

                            } catch (IOException e) {

                                System.out.println("bugsFailedToLog " + bug);

                                try {
                                    BufferedWriter writer = new BufferedWriter(new FileWriter("./res/jira/bugsFailedToLog.text", true));
                                    writer.write("\n" + bug);
                                    writer.close();
                                } catch (IOException ignored) {

                                }
                            }
                        }

                    } catch (IOException e) {
                        System.out.println("something went wrong with opening file reportedBugs.ini");
                    }


                }
            } else if (cause.getClass().getName().equalsIgnoreCase("org.openqa.selenium.TimeoutException")) {


                cause.printStackTrace(pw);
                trace = sw.getBuffer().toString();

            } else {

                cause.printStackTrace(pw);
                trace = sw.getBuffer().toString();
            }
        } else {
            trace = "Test failed.";
        }

        System.out.println("trace: " + trace);

        ExtentTestManager.getTest().log(LogStatus.FAIL, trace.replace("\n", "<br>"),
                ExtentTestManager.getTest().addBase64ScreenShot(base64Screenshot));
    }

    @Override
    public void onTestSkipped(ITestResult iTestResult) {
        ExtentTestManager.getTest().log(LogStatus.SKIP, "I don't know why but test skipped.");
    }

    @Override
    public void onTestFailedButWithinSuccessPercentage(ITestResult iTestResult) {
        System.out.println("Test failed but it is in defined success ratio " + getTestMethodName(iTestResult));
    }

}