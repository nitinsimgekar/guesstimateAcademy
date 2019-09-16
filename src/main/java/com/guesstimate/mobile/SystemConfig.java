package com.guesstimate.mobile;

public class SystemConfig {

	private String osName;
	private String androidHome;

	public SystemConfig() {
		osName = System.getProperty("os.name");
		androidHome = System.getenv("ANDROID_HOME");

	}

	public String getOsName() {
		return osName;
	}

	public void setOsName(String osName) {
		this.osName = osName;
	}

	public String getAndroidHome() {
		return androidHome;
	}

	public void setAndroidHome(String androidHome) {
		this.androidHome = androidHome;
	}

}
