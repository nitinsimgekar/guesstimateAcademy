package com.guesstimate.mobile;

import java.util.ArrayList;
import java.util.List;

import com.guesstimate.utils.DeviceCategory;
import com.guesstimate.utils.DeviceType;

public class AppEnv {

    private String name;
    private String operatorbuildNbrAndroid;
    private String supervisorbuildNbrAndroid;
    private String operatorbuildNbrIos;
    private String supervisorbuildNbrIos;
    private String oprJenkinsIosPath;
    private String oprHockeyIosPath;
    private String suprJenkinsIosPath;
    private String suprJenkinsAndroidPath;
    private String oprJenkinsAndroidPath;
    private String oprHockeyAndroidPath;
    private String appPackage;
    private String appActivity;
    private String oprBundleId;
    private String supBundleId;
    private String server;
    private DeviceType deviceType = DeviceType.ANDROID;
    private DeviceCategory deviceCategory;
    private List<String> priorityEmulator = new ArrayList<>();

    public String getOprHockeyIosPath() {
        return oprHockeyIosPath;
    }

    public AppEnv setOprHockeyIosPath(String oprHockeyIosPath) {
        this.oprHockeyIosPath = oprHockeyIosPath;
        return this;
    }

    public String getOprHockeyAndroidPath() {
        return oprHockeyAndroidPath;
    }

    public AppEnv setOprHockeyAndroidPath(String oprHockeyAndroidPath) {
        this.oprHockeyAndroidPath = oprHockeyAndroidPath;
        return this;
    }

    public String getServer() {
        return server;
    }

    public DeviceType getDeviceType() {
        return deviceType;
    }

    public AppEnv setDeviceType(DeviceType deviceType) {
        this.deviceType = deviceType;
        return this;
    }

    public AppEnv setServer(String server) {
        this.server = server;
        return this;
    }

    public String getOperatorbuildNbrAndroid() {
        return operatorbuildNbrAndroid;
    }

    public String getSupervisorbuildNbrAndroid() {
        return supervisorbuildNbrAndroid;
    }

    public String getOperatorbuildNbrIos() {
        return operatorbuildNbrIos;
    }

    public String getSupervisorbuildNbrIos() {
        return supervisorbuildNbrIos;
    }

    public AppEnv setOperatorbuildNbrAndroid(String operatorbuildNbrAndroid) {
        this.operatorbuildNbrAndroid = operatorbuildNbrAndroid;
        return this;
    }

    public AppEnv setSupervisorbuildNbrAndroid(String supervisorbuildNbrAndroid) {
        this.supervisorbuildNbrAndroid = supervisorbuildNbrAndroid;
        return this;
    }

    public AppEnv setOperatorbuildNbrIos(String operatorbuildNbrIos) {
        this.operatorbuildNbrIos = operatorbuildNbrIos;
        return this;
    }

    public AppEnv setSupervisorbuildNbrIos(String supervisorbuildNbrIos) {
        this.supervisorbuildNbrIos = supervisorbuildNbrIos;
        return this;
    }

    public DeviceCategory getDeviceCategory() {
        return deviceCategory;
    }

    public AppEnv setDeviceCategory(DeviceCategory deviceCategory) {
        this.deviceCategory = deviceCategory;
        return this;
    }

    public String getAppPackage() {
        return appPackage;
    }

    public AppEnv setAppPackage(String appPackage) {
        this.appPackage = appPackage;
        return this;
    }

    public String getAppActivity() {
        return appActivity;
    }

    public AppEnv setAppActivity(String appActivity) {
        this.appActivity = appActivity;
        return this;
    }

    public String getOprBundleId() {
        return oprBundleId;
    }

    public String getSupBundleId() {
        return supBundleId;
    }

    public AppEnv setOprBundleId(String oprBundleId) {
        this.oprBundleId = oprBundleId;
        return this;
    }

    public AppEnv setSupBundleId(String supBundleId) {
        this.supBundleId = supBundleId;
        return this;
    }

    public String getOprJenkinsAndroidPath() {
        return oprJenkinsAndroidPath;
    }

    public AppEnv setOprJenkinsAndroidPath(String oprJenkinsAndroidPath) {
        this.oprJenkinsAndroidPath = oprJenkinsAndroidPath;
        return this;
    }

    public String getOprJenkinsIosPath() {
        return oprJenkinsIosPath;
    }

    public AppEnv setOprJenkinsIosPath(String oprJenkinsIosPath) {
        this.oprJenkinsIosPath = oprJenkinsIosPath;
        return this;
    }

    public String getSuprJenkinsIosPath() {
        return suprJenkinsIosPath;
    }

    public AppEnv setSuprJenkinsIosPath(String suprJenkinsIosPath) {
        this.suprJenkinsIosPath = suprJenkinsIosPath;
        return this;
    }

    public String getSuprJenkinsAndroidPath() {
        return suprJenkinsAndroidPath;
    }

    public AppEnv setSuprJenkinsAndroidPath(String suprJenkinsAndroidPath) {
        this.suprJenkinsAndroidPath = suprJenkinsAndroidPath;
        return this;
    }

    public String getName() {
        return name;
    }

    public AppEnv setName(String name) {
        this.name = name;
        return this;
    }

    public List<String> getPriorityEmulator() {
        return priorityEmulator;
    }

    public AppEnv setPriorityEmulator(List<String> priorityEmulator) {
        this.priorityEmulator = priorityEmulator;
        return this;
    }

}
