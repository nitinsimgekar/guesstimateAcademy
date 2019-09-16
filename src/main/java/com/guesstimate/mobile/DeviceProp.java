package com.guesstimate.mobile;

import com.guesstimate.utils.DeviceType;

public class DeviceProp {

	// name field holds 'emulator name' for emulator's and
	// 'model' for real devices.
	private String name;
	private String configFilePath;
	private String folderPath;
	private String emulatorDeviceName;
	private String version;
	private boolean isRealDevice = false;
	private DeviceType deviceType;
	private AppEnv appEnv;
	private int port;

	public int getPort() {
		return port;
	}

	public AppEnv getAppEnv() {
		return appEnv;
	}

	public void setAppEnv(AppEnv appEnv) {
		this.appEnv = appEnv;
	}

	public DeviceType getDeviceType() {
		return deviceType;
	}

	public void setDeviceType(DeviceType deviceType) {
		this.deviceType = deviceType;
	}

	public boolean isRealDevice() {
		return isRealDevice;
	}

	public void setRealDevice(boolean isRealDevice) {
		this.isRealDevice = isRealDevice;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getVersion() {
		return version;
	}

	public void setVersion(String version) {
		this.version = version;
	}

	public String getConfigFilePath() {
		return configFilePath;
	}

	public void setConfigFilePath(String configFilePath) {
		this.configFilePath = configFilePath;
	}

	public String getFolderPath() {
		return folderPath;
	}

	public void setFolderPath(String folderPath) {
		this.folderPath = folderPath;
	}

	public String getEmulatorDeviceName() {
		return emulatorDeviceName;
	}

	public void setEmulatorDeviceName(String emulatorDeviceName) {
		this.emulatorDeviceName = emulatorDeviceName;
	}

	public void setPort(int port) {
		this.port = port;

	}

}
