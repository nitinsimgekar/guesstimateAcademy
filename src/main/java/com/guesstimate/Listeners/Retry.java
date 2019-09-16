package com.guesstimate.Listeners;

import org.testng.IRetryAnalyzer;
import org.testng.ITestResult;

public class Retry implements IRetryAnalyzer {

    private int count = 0;
    private static int maxTry = 2;

    @Override
    public boolean retry(ITestResult iTestResult) {
        if (!iTestResult.isSuccess() && !iTestResult.getThrowable().getClass().getName().equalsIgnoreCase("java.lang.AssertionError")) {
            if (count < maxTry) {
                count++;
                iTestResult.setStatus(ITestResult.FAILURE);
                return true;
            } else {
                iTestResult.setStatus(ITestResult.FAILURE);
            }
        } else {
            iTestResult.setStatus(ITestResult.SUCCESS);
        }
        return false;
    }
}
