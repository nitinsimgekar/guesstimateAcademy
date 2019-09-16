package com.guesstimate.ExtentReports;

import com.relevantcodes.extentreports.ExtentReports;

import java.text.SimpleDateFormat;
import java.util.Date;


public class ExtentManager {

    private static ExtentReports extent;

    public synchronized static ExtentReports getReporter() {
        if (extent == null) {

            SimpleDateFormat ft =
                    new SimpleDateFormat("yyyy.MM.dd 'at' hh:mm:ss");

            extent = new ExtentReports("./reports/" + ft.format(new Date()) + ".html", true);
        }
        return extent;
    }

}