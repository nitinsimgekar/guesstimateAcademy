package com.guesstimate.utils;

public enum DeviceEnvironment {

	Oreo_v1("27"), Oreo("26"), Nougat("24"), Nougat_v1("25"), Marshmallow("23"), lollipop_v1("22"), lollipop(
			"21"), ios_90("9.0"), ios_91("9.1"), ios_92("9.2"), ios_93("9.3"), ios_10("10.0"), ios_101("10.1"), ios_102(
					"10.2"), ios_103("10.3"), ios_11(
							"11.0"), ios_111("11.1"), ios_112("11.2"), ios_113("11.3"), ios_114("11.4"), ios_12("12.0"),ALL("");

	private String version;

	private DeviceEnvironment(String version) {
		this.version = version;
	}

	public String getVersion() {
		return version;
	}

}
