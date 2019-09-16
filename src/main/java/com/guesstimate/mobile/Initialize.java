package com.guesstimate.mobile;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.MalformedURLException;
import java.util.*;

import org.ini4j.Wini;

import com.guesstimate.utils.DeviceCategory;
import com.guesstimate.utils.DeviceEnvironment;
import com.guesstimate.utils.DeviceType;
import com.guesstimate.utils.RestServices;
import com.guesstimate.utils.ServiceGenerator;

import io.appium.java_client.AppiumDriver;
import okhttp3.ResponseBody;
import retrofit2.Call;
import retrofit2.Response;

public class Initialize {

    private Wini ini;
    private static Integer port = 5554;

    public ArrayList<DeviceProp> bringEnvironmentUp(DeviceEnvironment deviceEnvironment, AppEnv appEnv,
                                                    boolean runOnRealDeviceOnly) {

        // Read the ANDROID_HOME environmental variable from local system.
        SystemConfig getEnvironment = new SystemConfig();
        String androidHome = getEnvironment.getAndroidHome();

        if (appEnv.getDeviceType() == DeviceType.ANDROID) {

            if (androidHome == null) {
                System.err.println(
                        "ERROR: ANDROID_HOME is not set under environmental variable. HENCE TERMINATING THE SUITE.");

                System.exit(0);
            }
        }

        // Check if system confirmations are defined as expected
        checkSystemConfig();

        // Verify if emulators are already running.
        //checkIfEmulatorsRunning();

        // Get the connected real devices properties
        ArrayList<DeviceProp> devicesList = getAllRealDevices(appEnv.getDeviceType());

        // Get the existing emulators properties
        if (!runOnRealDeviceOnly) {
            devicesList = getAllEmulators(androidHome, devicesList, getEnvironment.getOsName(),
                    appEnv.getDeviceCategory(), appEnv);
        }

        // filter the emulators and real real device based on supported versions
        devicesList = filterSupportedEmulators(devicesList);

        // set env
        devicesList = setEnv(appEnv, devicesList);

        // Verify if required emulator's or real device is available in the
        // server. if real device and emulators are present for the same
        // version, real devices will be picked up
        devicesList = getRequiredEnvironment(deviceEnvironment, devicesList);

        // Bring emulators up and running.
        startEmulator(androidHome, devicesList);

        // Download apk and ipa builds from web
        if (appEnv.getServer().equals("qa")) {
            downloadBuilds(appEnv, devicesList, runOnRealDeviceOnly);
        }

        // Validate if emulators are up and running.
        validateRunningDevices(devicesList);

        // returns the running emulator list
        devicesList = returnRunningEmulators(devicesList);

        // install builds
        installBuilds(devicesList);

        return devicesList;

    }

    private void installBuilds(ArrayList<DeviceProp> devicesList) {

        for (DeviceProp deviceProp : devicesList) {

            switch (deviceProp.getDeviceType()) {

                case ANDROID:
                    installAPK(deviceProp);
                    break;

                case IOS:
                    installIPA(deviceProp);
                    break;

            }
        }

    }

    private ArrayList<DeviceProp> filterSupportedEmulators(ArrayList<DeviceProp> devicesList) {

        for (int i = 0; i < devicesList.size(); i++) {

            DeviceProp deviceProp = devicesList.get(i);

            switch (deviceProp.getDeviceType()) {
                case ANDROID:

                    if (Integer.parseInt(deviceProp.getVersion()) < 23) {
                        devicesList.remove(deviceProp);

                        if (deviceProp.isRealDevice())
                            System.out.println("THE REAL DEVICE IS NOT SUPPORTED. '" + deviceProp.getName()
                                    + "' with version '" + deviceProp.getVersion() + "'");
                    }
                    break;

                case IOS:
                    if (Integer.parseInt(deviceProp.getVersion().substring(0, deviceProp.getVersion().indexOf("."))) < 9) {
                        devicesList.remove(deviceProp);

                        if (deviceProp.isRealDevice())
                            System.out.println("THE REAL DEVICE IS NOT SUPPORTED. '" + deviceProp.getName()
                                    + "' with version '" + deviceProp.getVersion() + "'");
                    }
                    break;

                default:
                    break;
            }

        }

        return devicesList;

    }

    private boolean writeResponseBodyToFile(File file, okhttp3.ResponseBody body) {
        try {

            InputStream inputStream = null;
            OutputStream outputStream = null;

            try {
                byte[] fileReader = new byte[4096];

                long fileSize = body.contentLength();
                long fileSizeDownloaded = 0;

                inputStream = body.byteStream();
                outputStream = new FileOutputStream(file);

                while (true) {
                    int read = inputStream.read(fileReader);

                    if (read == -1) {
                        break;
                    }

                    outputStream.write(fileReader, 0, read);

                    fileSizeDownloaded += read;

                }

                outputStream.flush();

                return true;
            } catch (IOException e) {
                return false;
            } finally {
                if (inputStream != null) {
                    inputStream.close();
                }

                if (outputStream != null) {
                    outputStream.close();
                }
            }
        } catch (IOException e) {
            return false;
        }
    }

    private ArrayList<DeviceProp> setEnv(AppEnv appEnv, ArrayList<DeviceProp> devicesList) {

        for (DeviceProp deviceProp : devicesList) {

            switch (deviceProp.getDeviceType()) {
                case ANDROID:
                    deviceProp.setAppEnv(appEnv);
                    break;

                case IOS:
                    deviceProp.setAppEnv(appEnv);
                    break;
            }

        }

        return devicesList;

    }

    private void downloadBuilds(AppEnv appEnv, ArrayList<DeviceProp> devicesList, boolean runOnRealDeviceOnly) {

        File dir = new File("./res/builds");
        if (!dir.exists())
            dir.mkdir();

        boolean isAndroidBuildRequired = false;
        boolean isIosBuildRequired = false;

        for (DeviceProp deviceProp : devicesList) {

            switch (deviceProp.getDeviceType()) {
                case ANDROID:
                    isAndroidBuildRequired = true;
                    break;

                case IOS:
                    isIosBuildRequired = true;
                    break;
            }

        }

        String operatorbuildNbrIos = appEnv.getOperatorbuildNbrIos();
        String supervisorbuildNbrIos = appEnv.getSupervisorbuildNbrIos();

        String operatorbuildNbrAndroid = appEnv.getOperatorbuildNbrAndroid();
        String supervisorbuildNbrAndroid = appEnv.getSupervisorbuildNbrAndroid();
        if (runOnRealDeviceOnly && isIosBuildRequired && operatorbuildNbrIos != null) {

            File file = new File("./res/builds/" + appEnv.getName() + "_ios_" + operatorbuildNbrIos + ".ipa");

            if (!file.exists()) {

                RestServices restServices = ServiceGenerator.createService(RestServices.class, "nsimgekar", "123456");
                Call call = restServices.getIpa(appEnv.getOprJenkinsIosPath() + ".ipa");
                try {
                    Response response = call.execute();
                    if (response.isSuccessful()) {
                        ResponseBody responseBody = (ResponseBody) response.body();

                        writeResponseBodyToFile(file, responseBody);
                    } else {
                        System.err
                                .println("ERROR: WHILE DOWNLOADING THE BUILD: '" + response.message() + "' from path '"
                                        + appEnv.getOprJenkinsIosPath() + ".ipa' and file '" + file.getAbsolutePath() + "'. HENCE TERMINATING THE SUITE.");
                        System.exit(0);
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }

            }

        }

        if (!runOnRealDeviceOnly && isAndroidBuildRequired && operatorbuildNbrAndroid != null) {

            File file = new File("./res/builds/" + appEnv.getName() + "_android_" + operatorbuildNbrAndroid + ".apk");

            if (!file.exists()) {

                RestServices restServices = ServiceGenerator.createService(RestServices.class, "nsimgekar", "123456");
                Call call = restServices.getIpa(appEnv.getOprJenkinsAndroidPath() + ".apk");
                try {
                    Response response = call.execute();
                    if (response.isSuccessful()) {
                        ResponseBody responseBody = (ResponseBody) response.body();

                        writeResponseBodyToFile(file, responseBody);
                    } else {
                        System.err
                                .println("ERROR: WHILE DOWNLOADING THE BUILD: '" + response.message() + "' from path '"
                                        + appEnv.getOprJenkinsAndroidPath() + ".apk' and file '" + file.getAbsolutePath() + "'. HENCE TERMINATING THE SUITE.");
                        System.exit(0);
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }

            }
        }

        if (runOnRealDeviceOnly && supervisorbuildNbrIos != null && !supervisorbuildNbrIos.isEmpty()) {
            File file = new File("./res/builds/" + appEnv.getName() + "_ios_" + supervisorbuildNbrIos + ".ipa");

            if (!file.exists()) {

                RestServices restServices = ServiceGenerator.createService(RestServices.class, "nsimgekar", "123456");
                Call call = restServices.getIpa(appEnv.getSuprJenkinsIosPath() + ".ipa");
                try {
                    Response response = call.execute();
                    if (response.isSuccessful()) {
                        ResponseBody responseBody = (ResponseBody) response.body();

                        writeResponseBodyToFile(file, responseBody);
                    } else {
                        System.err
                                .println("ERROR: WHILE DOWNLOADING THE BUILD: '" + response.message() + "' from path '"
                                        + appEnv.getSuprJenkinsIosPath() + ".ipa' and file '" + file.getAbsolutePath() + "'. HENCE TERMINATING THE SUITE.");
                        System.exit(0);
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }

            }
        }

        if (runOnRealDeviceOnly && supervisorbuildNbrAndroid != null && !supervisorbuildNbrAndroid.isEmpty()) {
            File file = new File("./res/builds/" + appEnv.getName() + "_android_" + supervisorbuildNbrAndroid + ".apk");

            if (!file.exists()) {

                RestServices restServices = ServiceGenerator.createService(RestServices.class, "nsimgekar", "123456");
                Call call = restServices.getIpa(appEnv.getSuprJenkinsAndroidPath() + ".apk");
                try {
                    Response response = call.execute();
                    if (response.isSuccessful()) {
                        ResponseBody responseBody = (ResponseBody) response.body();

                        writeResponseBodyToFile(file, responseBody);
                    } else {
                        System.err
                                .println("ERROR: WHILE DOWNLOADING THE BUILD: '" + response.message() + "' from path '"
                                        + appEnv.getSuprJenkinsAndroidPath() + ".apk' and file '" + file.getAbsolutePath() + "'. HENCE TERMINATING THE SUITE.");
                        System.exit(0);
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }

            }
        }

    }

    private ArrayList<DeviceProp> getRequiredEnvironment(DeviceEnvironment deviceEnvironment,
                                                         ArrayList<DeviceProp> devicesList) {

        for (int i = 0; i < devicesList.size() - 1; i++) {

            DeviceProp deviceProp_i = devicesList.get(i);

            for (int k = devicesList.size() - 1; k > i; k--) {

                DeviceProp deviceProp_k = devicesList.get(k);

                if (deviceProp_i.getVersion().startsWith(deviceProp_k.getVersion())) {

                    if (!deviceProp_k.isRealDevice()) {
                        devicesList.remove(deviceProp_k);
                    } else if (!deviceProp_i.isRealDevice()) {
                        devicesList.remove(deviceProp_i);
                    }

                }

            }
        }

        boolean isEmulatorAvailable = false;

        if (deviceEnvironment.getVersion().equalsIgnoreCase("")) {

            isEmulatorAvailable = true;

            System.out.println("Your test cases will be executed on the following environments:");
            for (DeviceProp deviceProp : devicesList) {
                System.out.println("'" + deviceProp.getName() + "' with version '" + deviceProp.getVersion() + "'");

            }

            if (devicesList.isEmpty()) {
                System.err.println(
                        "ERROR: USER SPECIFIC EMULATORS NOT FOUND. PLEASE CREATE ONE USING AVD MANAGER/XCODE. HENCE TERMINATING THE SUITE.");

                System.exit(0);
            }

            return devicesList;
        }

        ArrayList<DeviceProp> specificEmulatorsList = new ArrayList<DeviceProp>();

        for (DeviceProp deviceProp : devicesList) {

            if (deviceProp.getVersion().startsWith(deviceEnvironment.getVersion())) {
                isEmulatorAvailable = true;
                specificEmulatorsList.add(deviceProp);
            }

        }

        if (!isEmulatorAvailable) {
            System.err.println(
                    "ERROR: USER SPECIFIC EMULATORS NOT FOUND. PLEASE CREATE ONE USING AVD MANAGER/XCODE. HENCE TERMINATING THE SUITE.");

            System.exit(0);
        }

        if (specificEmulatorsList.size() == 0) {
            System.err.println(
                    "ERROR: USER SPECIFIC EMULATORS NOT FOUND. PLEASE CREATE ONE USING AVD MANAGER. HENCE TERMINATING THE SUITE.");

            System.exit(0);
        }

        return specificEmulatorsList;

    }

    private void checkSystemConfig() {

        try {

            File file = new File("./res/runningEmulators.ini");
            if (file.exists()) {
                file.delete();
            }

            file.createNewFile();
            System.out.println("Emulator related data is stored under '" + file.getName() + "' for your referance.");

            ini = new Wini(file);
            ini.add("root");
            ini.store();

        } catch (IOException e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);

            System.out.println("ERROR: SOMETHING WENT WRONG DUE TO " + sw.toString());
            System.exit(0);
        }
    }

    private void startEmulator(String androidHome, ArrayList<DeviceProp> emulatorsList) {

        try {

            for (DeviceProp emulator : emulatorsList) {

                if (!emulator.isRealDevice()) {

                    switch (emulator.getDeviceType()) {
                        case ANDROID:
                            bringAndroidEmulatorUp(androidHome, emulator);
                            break;

                        case IOS:
                            bringIosEmulatorUp(emulator);
                            break;

                    }

                }

            }

        } catch (IOException e1) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e1.printStackTrace(pw);

            System.out.println("ERROR: SOMETHING WENT WRONG DUE TO " + sw.toString());
            System.exit(0);
        }
    }

    private void bringIosEmulatorUp(DeviceProp emulator) throws IOException {

        System.out.println("Bringing the ios simulator '" + emulator.getName() + "' up. Please wait...");

        Process p = Runtime.getRuntime().exec("instruments -w " + emulator.getEmulatorDeviceName());
        BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));

        while (true) {

            String emulatorline = r.readLine();
            if (emulatorline == null) {
                break;
            }

            if (emulatorline.contains("Waiting for device to boot...")) {
                System.out.println("Waiting for device to boot...");
                break;
            }

        }
    }

    private void bringAndroidEmulatorUp(String androidHome, DeviceProp emulator) throws IOException {

        System.out.println("Bringing the android emulator '" + emulator.getName() + "' up. Please wait...");
        Runtime.getRuntime().exec("emulator -avd " + emulator.getName() + " -port " + port, null,
                new File(androidHome + "/tools"));

        emulator.setEmulatorDeviceName("emulator-" + port);
        System.out.println("'" + "emulator-" + port + "' is building up.");

    }

    private void validateRunningDevices(ArrayList<DeviceProp> emulatorsList) {

        try {
            Thread.sleep(10000);
        } catch (InterruptedException e2) {

            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e2.printStackTrace(pw);

            System.err.println("WARN: SOMETHING WENT WRONG DUE TO " + sw.toString());
        }

        for (DeviceProp deviceProp : emulatorsList) {

            if (!deviceProp.isRealDevice()) {

                switch (deviceProp.getDeviceType()) {
                    case ANDROID:
                        verifyAndroidEmulator(deviceProp);
                        break;
                    case IOS:
                        verifyIOSEmulator(deviceProp);
                }

            }
        }
    }

    private void verifyIOSEmulator(DeviceProp deviceProp) {

        boolean isDeviceUp = false;

        try {

            Thread.sleep(10000);

            System.out.println("Validating emulator. Please wait...");
            Process validateProcess = Runtime.getRuntime().exec("xcrun simctl list devices");

            BufferedReader r = new BufferedReader(new InputStreamReader(validateProcess.getInputStream()));
            String ValidateLine;
            while (true) {
                ValidateLine = r.readLine();
                if (ValidateLine == null || ValidateLine.equalsIgnoreCase("")) {
                    break;
                }
                if (ValidateLine.contains(deviceProp.getEmulatorDeviceName())) {

                    String status = ValidateLine
                            .substring(ValidateLine.lastIndexOf("(") + 1, ValidateLine.lastIndexOf(")")).trim();
                    if (status.equalsIgnoreCase("Booted")) {
                        isDeviceUp = true;
                    }

                }

            }

            if (!isDeviceUp) {
                System.err.println(
                        "ERROR: SOMETHING WENT WRONG DURING INITILIZATION OF EMULATORS. PLEASE VERIFY USING AVD MANAGER.");

            } else {
                System.out.println("'" + deviceProp.getEmulatorDeviceName() + "' is now up and running.");

                updateEmulatorRes(deviceProp, true, false);

            }

        } catch (IOException | InterruptedException e1) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e1.printStackTrace(pw);

            System.err.println("ERROR: SOMETHING WENT WRONG DUE TO " + sw.toString());
            System.exit(0);

        }

        isDeviceUp = false;
    }

    private void verifyAndroidEmulator(DeviceProp deviceProp) {

        boolean isDeviceUp = false;

        try {

            System.out.println("Validating emulator. Please wait...");
            Process validateProcess = Runtime.getRuntime().exec("adb devices");

            BufferedReader r = new BufferedReader(new InputStreamReader(validateProcess.getInputStream()));
            String ValidateLine;
            while (true) {
                ValidateLine = r.readLine();
                if (ValidateLine == null || ValidateLine.equalsIgnoreCase("")) {
                    break;
                }
                if (!ValidateLine.equalsIgnoreCase("List of devices attached")) {
                    if (ValidateLine.contains("offline")) {
                        Thread.sleep(5000);
                        Process validateProcess2 = Runtime.getRuntime().exec("adb devices");
                        r = new BufferedReader(new InputStreamReader(validateProcess2.getInputStream()));

                    } else if (ValidateLine.replace("device", "").trim()
                            .equalsIgnoreCase(deviceProp.getEmulatorDeviceName())) {
                        isDeviceUp = true;

                    }
                }

            }

            if (!isDeviceUp) {
                System.err.println(
                        "ERROR: SOMETHING WENT WRONG DURING INITILIZATION OF EMULATORS. PLEASE VERIFY USING AVD MANAGER.");

            } else {
                System.out.println("'" + deviceProp.getEmulatorDeviceName() + "' is now up and running.");

                updateEmulatorRes(deviceProp, true, false);

            }

        } catch (IOException | InterruptedException e1) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e1.printStackTrace(pw);

            System.err.println("ERROR: SOMETHING WENT WRONG DUE TO " + sw.toString());
            System.exit(0);

        }

        isDeviceUp = false;
    }

    private ArrayList<DeviceProp> returnRunningEmulators(ArrayList<DeviceProp> emulatorsList) {

        Set<DeviceProp> ids = new HashSet<DeviceProp>();

        for (int i = 0; i < emulatorsList.size(); i++) {

            DeviceProp deviceProp = emulatorsList.get(i);

            if (deviceProp.getEmulatorDeviceName() == null || deviceProp.getEmulatorDeviceName().equalsIgnoreCase("")) {
                ids.add(deviceProp);
                updateEmulatorRes(deviceProp, false, false);
            }

        }

        for (DeviceProp id : ids) {
            emulatorsList.remove(id);
        }

        return emulatorsList;
    }

    private void updateEmulatorRes(DeviceProp emulator, boolean isEmulatorUp, boolean isAppInstalled) {

        try {

            ini.put(emulator.getName(), "name", emulator.getName());
            ini.put(emulator.getName(), "emulatorDeviceName", emulator.getEmulatorDeviceName());
            ini.put(emulator.getName(), "isEmulatorUp", isEmulatorUp);
            ini.put(emulator.getName(), "isAppInstalled", isAppInstalled);

            ini.store();

        } catch (IOException e1) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e1.printStackTrace(pw);

            System.out.println("ERROR: failed to update RUNNINGEMULATORS.INI due to " + sw.toString());
            System.exit(0);
        }
    }

    private ArrayList<DeviceProp> getAllRealDevices(DeviceType deviceType) {

        ArrayList<DeviceProp> realDevicesList = new ArrayList<DeviceProp>();

        try {

            System.out.println("Looking for real devices. Please wait...");

            Process p;
            BufferedReader r;
            String line;

            DeviceProp prop = new DeviceProp();
            List<String> unconfiguredDevices = new ArrayList<String>();

            if (deviceType == DeviceType.ANDROID) {

                p = Runtime.getRuntime().exec("adb devices");
                r = new BufferedReader(new InputStreamReader(p.getInputStream()));

                while (true) {
                    line = r.readLine();
                    if (line == null) {
                        break;
                    }
                    if (!line.contains("List of devices attached")) {
                        if (line.contains("device")) {
                            if (!line.startsWith("emulator-")) {
                                String emuName = line.replace("device", "").trim();
                                prop.setEmulatorDeviceName(emuName);
                                prop.setRealDevice(true);

                                p = Runtime.getRuntime().exec("adb -s " + emuName + " shell getprop ro.build.version.sdk");
                                r = new BufferedReader(new InputStreamReader(p.getInputStream()));

                                while (true) {
                                    line = r.readLine();
                                    if (line == null || line.equals("")) {
                                        unconfiguredDevices.add(emuName);
                                        break;
                                    }

                                    prop.setVersion(line);
                                    prop.setDeviceType(DeviceType.ANDROID);

                                }

                                p = Runtime.getRuntime().exec("adb -s " + emuName + " shell getprop ro.product.model");
                                r = new BufferedReader(new InputStreamReader(p.getInputStream()));

                                while (true) {
                                    line = r.readLine();
                                    if (line == null || line.equals("")) {
                                        break;
                                    }

                                    prop.setName(line);

                                }

                                realDevicesList.add(prop);
                                prop = new DeviceProp();
                            }

                        }

                    }
                }

            }

            p = Runtime.getRuntime().exec("instruments -s devices");
            r = new BufferedReader(new InputStreamReader(p.getInputStream()));

            while (true) {
                line = r.readLine();
                if (line == null || line.contains("Known Templates")) {
                    break;
                }

                if (line.contains("error: tool 'instruments' requires Xcode")) {

                    System.err.println(
                            "ERROR: Root permissions required. execute the below cmd in terminal. 'sudo xcode-select -s /Applications/Xcode.app/Contents/Developer'");
                    System.exit(0);

                }

                if (!line.contains("Simulator") && line.contains("(") && !line.contains("Apple")
                        && !line.contains("Mac Mini") && !line.contains("MacBook") && !line.contains("(null) [")
                        && !line.contains("Jenkins")) {

                    prop.setName(line.substring(0, line.indexOf("(")).trim());
                    prop.setVersion(line.substring(line.indexOf("(") + 1, line.indexOf(")")).trim());
                    prop.setEmulatorDeviceName(line.substring(line.indexOf("[") + 1, line.indexOf("]")).trim());
                    prop.setRealDevice(true);
                    prop.setDeviceType(DeviceType.IOS);
                    realDevicesList.add(prop);
                    prop = new DeviceProp();

                }

            }

            if (unconfiguredDevices.size() > 0) {
                System.out.println("The following devices were not configured properly: " + unconfiguredDevices);

                System.out.println("Please execute the following command to know more, 'instruments -s devices'");
            }

            if (realDevicesList.size() == 0) {
                System.out.println("WARN: NO REAL DEVICES FOUND.");
            }

            r.close();
            p.destroy();

        } catch (IOException e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);

            System.out.println("ERROR: SOMETHING WENT WRONG DUE TO " + sw.toString());
            System.exit(0);
        }

        return realDevicesList;
    }

    private ArrayList<DeviceProp> getAllEmulators(String androidHome, ArrayList<DeviceProp> devicesList, String osName,
                                                  DeviceCategory deviceCat, AppEnv appEnv) {

        try {

            System.out.println("Looking for emulators. Please wait...");

            Process p;
            DeviceProp prop = new DeviceProp();
            List<String> priorityEmulator = appEnv.getPriorityEmulator();
            String deviceCategory = "";

            if (appEnv.getDeviceType() == DeviceType.ANDROID) {

                File dir = new File(androidHome + "/tools/bin");
                String avdmanager = "avdmanager.bat";
                boolean check = new File(dir, avdmanager).exists();

                Runtime runtime = Runtime.getRuntime();

                if (osName.contains("Windows")) {

                    if (check) {
                        p = runtime.exec("cmd /c " + avdmanager + " list avd", null, dir);
                    } else {
                        p = runtime.exec("cmd /c android list avd", null, new File(androidHome + "/tools"));
                    }

                } else {

                    if (check) {
                        p = runtime.exec(avdmanager + " list avd");
                    } else {
                        p = runtime.exec("android list avd");
                    }
                }

                BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));
                String line;

                while (true) {
                    line = r.readLine();
                    if (line == null) {
                        break;
                    }

                    if (line.contains("The 'android' command is no longer available.")) {

                        System.err.println(
                                "ERROR: ANDROID COMMAND IS NOT YET INSTALLED. PLEASE SETUP SDK TOOLS FROM SDK MANAGER/ANDROID STUDIO. HENCE TERMINATING THE SUITE.");
                        System.exit(0);

                    }

                    if (line.contains("Name:")) {
                        prop.setName(line.replace("Name:", "").trim());

                    } else if (line.contains("Path:")) {
                        prop.setFolderPath(line.replace("Path:", "").trim());
                        String path = line.replace("Path:", "").trim().replace(".avd", ".ini");
                        prop.setConfigFilePath(path);

                        Wini ini = new Wini(new File(path.substring(0, path.lastIndexOf(".")) + ".avd/config.ini"));

                        boolean isMatch = ini.get("?", "hw.device.name").contains("Tablet");

                        ini = new Wini(new File(path));

                        if (deviceCat == DeviceCategory.TABLET && isMatch) {
                            prop.setVersion(ini.get("?", "target").replace("android-", ""));
                            prop.setDeviceType(DeviceType.ANDROID);
                        } else if (deviceCat == DeviceCategory.PHONE && !isMatch) {
                            prop.setVersion(ini.get("?", "target").replace("android-", ""));
                            prop.setDeviceType(DeviceType.ANDROID);
                        } else {
                            prop = new DeviceProp();
                        }

                    } else if (line.contains("---------")) {
                        if (prop.getName() != null) {
                            prop.setRealDevice(false);
                            devicesList.add(prop);
                        }
                        prop = new DeviceProp();

                    }

                }

                if (prop.getName() != null) {
                    prop.setRealDevice(false);
                    devicesList.add(prop);
                }
                prop = new DeviceProp();

                if (deviceCat == DeviceCategory.TABLET) {
                    deviceCategory = "iPad";
                } else if (deviceCat == DeviceCategory.PHONE) {
                    deviceCategory = "iPhone";
                }

            }

            p = Runtime.getRuntime().exec("instruments -s devices");
            BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));

            while (true) {
                String line = r.readLine();
                if (line == null) {
                    break;
                }

                if (line.contains("error: tool 'instruments' requires Xcode")) {

                    System.err.println(
                            "ERROR: Root permissions required. execute the below cmd in terminal. 'sudo xcode-select -s /Applications/Xcode.app/Contents/Developer'");
                    System.exit(0);

                }

                if (line.contains("Simulator") && !line.contains("Apple") && !line.contains("Mac Mini")
                        && !line.contains("MacBook") && line.contains(deviceCategory) && !line.contains("Jenkins")) {

                    prop.setEmulatorDeviceName(line.substring(line.indexOf("[") + 1, line.indexOf("]")).trim());

                    line = line.substring(0, line.indexOf("["));

                    if (!priorityEmulator.isEmpty()
                            && !priorityEmulator.contains(line.substring(0, line.lastIndexOf("(")).trim())) {
                        continue;

                    }

                    prop.setName(line.substring(0, line.lastIndexOf("(")).trim());
                    prop.setVersion(line.substring(line.lastIndexOf("(") + 1, line.lastIndexOf(")")).trim());

                    prop.setRealDevice(false);
                    prop.setDeviceType(DeviceType.IOS);
                    devicesList.add(prop);
                    prop = new DeviceProp();

                }

            }

            r.close();
            p.destroy();

            if (devicesList.size() == 0) {
                System.err.println(
                        "ERROR: NO EMULATORS FOUND. PLEASE CREATE ONE USING AVD MANAGER WITH A MIN ANDROID MARSHMALLOW CONFIG. HENCE TERMINATING THE SUITE.");
                System.exit(0);
            }

        } catch (

                IOException e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);

            System.out.println("ERROR: SOMETHING WENT WRONG DUE TO " + sw.toString());
            System.exit(0);
        }

        return devicesList;
    }

    public void installIPA(DeviceProp emulator) {

        try {

            if (emulator.isRealDevice()) {

                System.out.println("Installing IPA file. Please wait...");

                if (emulator.getAppEnv().getOperatorbuildNbrIos() != null) {

                    File ipa = new File("./res/builds/" + emulator.getAppEnv().getName() + "_ios_"
                            + emulator.getAppEnv().getOperatorbuildNbrIos() + ".ipa");

                    if (!ipa.exists()) {
                        System.err.println("ERROR: NO IPA FILE FOUND '" + ipa
                                + "'. PLEASE MENTION THE JENKINS BUILD NUMBER. HENCE TERMINATING THE SUITE.");
                        System.exit(0);
                    }

                    System.out.println("Verifying if " + emulator.getAppEnv().getName()
                            + " ipa is already installed on '" + emulator.getName() + "'. Please wait...");

                    Runtime.getRuntime().exec("ideviceinstaller -u " + emulator.getEmulatorDeviceName() + " -U "
                            + emulator.getAppEnv().getOprBundleId());

                    Process p = Runtime.getRuntime()
                            .exec("ideviceinstaller -u " + emulator.getEmulatorDeviceName() + " -i " + "./res/builds/"
                                    + emulator.getAppEnv().getName() + "_ios_"
                                    + emulator.getAppEnv().getOperatorbuildNbrIos() + ".ipa");

                    BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));

                    while (true) {
                        String line = r.readLine();
                        if (line == null) {
                            break;
                        }

                        if (line.contains("Could not connect to lockdownd")) {

                            System.err
                                    .println("Could not connect to lockdownd. Exiting. Refer to README file for fix.");
                            System.exit(0);

                        }
                    }

                    r.close();
                    p.destroy();

                }

                if (emulator.getAppEnv().getSupervisorbuildNbrIos() != null && !emulator.getAppEnv().getSupervisorbuildNbrIos().isEmpty()) {

                    File ipa = new File("./res/builds/" + emulator.getAppEnv().getName() + "_ios_"
                            + emulator.getAppEnv().getSupervisorbuildNbrIos() + ".ipa");

                    if (!ipa.exists()) {
                        System.err.println("ERROR: NO IPA FILE FOUND '" + ipa
                                + "'. PLEASE MENTION THE JENKINS BUILD NUMBER. HENCE TERMINATING THE SUITE.");
                        System.exit(0);
                    }

                    System.out.println("Verifying if " + emulator.getAppEnv().getName()
                            + " ipa is already installed on '" + emulator.getName() + "'. Please wait...");

                    Runtime.getRuntime().exec("ideviceinstaller -u " + emulator.getEmulatorDeviceName() + " -U "
                            + emulator.getAppEnv().getSupBundleId());

                    Process p = Runtime.getRuntime()
                            .exec("ideviceinstaller -u " + emulator.getEmulatorDeviceName() + " -i " + "./res/builds/"
                                    + emulator.getAppEnv().getName() + "_ios_"
                                    + emulator.getAppEnv().getSupervisorbuildNbrIos() + ".ipa");

                    BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));

                    while (true) {
                        String line = r.readLine();
                        if (line == null) {
                            break;
                        }

                        if (line.contains("Could not connect to lockdownd")) {

                            System.err
                                    .println("Could not connect to lockdownd. Exiting. Refer to README file for fix.");
                            System.exit(0);

                        }
                    }

                    r.close();
                    p.destroy();

                }


            } else {

                System.out.println("Installing APP file. Please wait...");

                if (emulator.getAppEnv().getOperatorbuildNbrIos() != null) {

                    File ipa = new File("./res/builds/" + emulator.getAppEnv().getName() + "_ios_"
                            + emulator.getAppEnv().getOperatorbuildNbrIos() + ".app");

                    if (!ipa.exists()) {
                        System.err.println("ERROR: NO APP FILE FOUND '" + ipa
                                + "'. PLEASE MENTION THE JENKINS BUILD NUMBER. HENCE TERMINATING THE SUITE.");
                        System.exit(0);
                    }

                    System.out.println("Verifying if " + emulator.getAppEnv().getName()
                            + " app is already installed on '" + emulator.getName() + "'. Please wait...");

                    Runtime.getRuntime().exec("xcrun simctl uninstall " + emulator.getEmulatorDeviceName() + " -U "
                            + emulator.getAppEnv().getOprBundleId());

                    Process p = Runtime.getRuntime()
                            .exec("xcrun simctl install " + emulator.getEmulatorDeviceName() + " ./res/builds/"
                                    + emulator.getAppEnv().getName() + "_ios_"
                                    + emulator.getAppEnv().getOperatorbuildNbrIos() + ".app");

                    BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));

                    while (true) {
                        String line = r.readLine();
                        if (line == null) {
                            break;
                        }

                        if (line.contains("Could not connect to lockdownd")) {

                            System.err
                                    .println("Could not connect to lockdownd. Exiting. Refer to README file for fix.");
                            System.exit(0);

                        } else if (line.contains("The bundle identifier of the application could not be determined.")) {

                            System.err
                                    .println("The bundle identifier of the application could not be determined.");
                            System.exit(0);

                        }
                    }

                    r.close();
                    p.destroy();

                }

                if (emulator.getAppEnv().getSupervisorbuildNbrIos() != null && !emulator.getAppEnv().getSupervisorbuildNbrIos().isEmpty()) {

                    File ipa = new File(" ./res/builds/" + emulator.getAppEnv().getName() + "_ios_"
                            + emulator.getAppEnv().getSupervisorbuildNbrIos() + ".app");

                    if (!ipa.exists()) {
                        System.err.println("ERROR: NO APP FILE FOUND '" + ipa
                                + "'. PLEASE MENTION THE JENKINS BUILD NUMBER. HENCE TERMINATING THE SUITE.");
                        System.exit(0);
                    }

                    System.out.println("Verifying if " + emulator.getAppEnv().getName()
                            + " app is already installed on '" + emulator.getName() + "'. Please wait...");

                    Runtime.getRuntime().exec("xcrun simctl uninstall " + emulator.getEmulatorDeviceName() + " -U "
                            + emulator.getAppEnv().getOprBundleId());

                    Process p = Runtime.getRuntime()
                            .exec("xcrun simctl install " + emulator.getEmulatorDeviceName() + " ./res/builds/"
                                    + emulator.getAppEnv().getName() + "_ios_"
                                    + emulator.getAppEnv().getSupervisorbuildNbrIos() + ".app");

                    BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));

                    while (true) {
                        String line = r.readLine();
                        if (line == null) {
                            break;
                        }

                        if (line.contains("Could not connect to lockdownd")) {

                            System.err
                                    .println("Could not connect to lockdownd. Exiting. Refer to README file for fix.");
                            System.exit(0);

                        }
                    }

                    r.close();
                    p.destroy();

                }

            }

        } catch (IOException e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);

            System.out.println("ERROR: SOMETHING WENT WRONG DUE TO " + sw.toString());
            System.exit(0);
        }

    }

    private void installAPK(DeviceProp emulatorsList) {

        try {
            Thread.sleep(30000);
        } catch (InterruptedException e1) {
            e1.printStackTrace();
        }

        String apk = "";

        try {
            apk = new File("./res/builds/" + emulatorsList.getAppEnv().getName() + "_android_"
                    + emulatorsList.getAppEnv().getOperatorbuildNbrAndroid() + ".apk").getAbsolutePath();
        } catch (NullPointerException e) {
            System.err.println(
                    "ERROR: .APK file not found. please download one from the build applications. Hence terminating the suite.");
            System.exit(0);
        }

        try {

            System.out.println("Verifying if " + emulatorsList.getAppEnv().getName() + " app is already installed on '"
                    + emulatorsList.getName() + "'. Please wait...");

            Process p = Runtime.getRuntime().exec("adb -s " + emulatorsList.getEmulatorDeviceName()
                    + " shell pm list packages " + emulatorsList.getAppEnv().getAppPackage());
            BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));

            String line;
            while (true) {
                line = r.readLine();
                if (line == null) {
                    break;
                }

                if (line.contains("device offline") || line.contains("no devices/emulators found")) {
                    Thread.sleep(10000);
                    Process validateProcess2 = Runtime.getRuntime()
                            .exec("adb -s " + emulatorsList.getEmulatorDeviceName() + " shell pm list packages "
                                    + emulatorsList.getAppEnv().getAppPackage());
                    r = new BufferedReader(new InputStreamReader(validateProcess2.getInputStream()));

                }

                if (line.contains("package")) {

                    System.out.println(emulatorsList.getAppEnv().getName() + " is already installed. Uninstalling...");

                    p = Runtime.getRuntime().exec("adb -s " + emulatorsList.getEmulatorDeviceName()
                            + " shell pm uninstall " + emulatorsList.getAppEnv().getAppPackage());
                    r = new BufferedReader(new InputStreamReader(p.getInputStream()));

                }

                if (line.trim().equalsIgnoreCase("success")) {

                    System.out.println(emulatorsList.getAppEnv().getName() + " uninstalled successfully.");

                }

                if (line.trim().equalsIgnoreCase("failed")) {
                    System.err.println("FAILED TO INSTALL " + emulatorsList.getAppEnv().getName()
                            + " APP. HENCE TERMINATING THE SUITE.");

                }

            }

            System.out.println("Installing  " + emulatorsList.getAppEnv().getName() + " app '" + apk + "' on '"
                    + emulatorsList.getName() + "'. Please wait...");

            p = Runtime.getRuntime().exec("adb -s " + emulatorsList.getEmulatorDeviceName() + " install " + apk);
            r = new BufferedReader(new InputStreamReader(p.getInputStream()));

            while (true) {
                line = r.readLine();

                if (line == null) {
                    break;
                }

                if (line.contains("waiting for device")) {
                    Thread.sleep(5000);
                    p = Runtime.getRuntime()
                            .exec("adb -s " + emulatorsList.getEmulatorDeviceName() + " install " + apk);
                    r = new BufferedReader(new InputStreamReader(p.getInputStream()));

                }

                if (line.contains("INSTALL_FAILED_ALREADY_EXISTS")) {
                    System.out.println("Running test on the existing app.");
                }

            }

            System.out.println("Verifying if " + emulatorsList.getAppEnv().getName()
                    + " app is installed sucessfully on '" + emulatorsList.getName() + "'. Please wait...");

            boolean isAppInstalled = false;

            p = Runtime.getRuntime().exec("adb -s " + emulatorsList.getEmulatorDeviceName() + " shell pm list packages "
                    + emulatorsList.getAppEnv().getAppPackage());
            r = new BufferedReader(new InputStreamReader(p.getInputStream()));

            while (true) {
                line = r.readLine();
                if (line == null) {
                    break;
                }

                if (line.contains("device offline") || line.contains("no devices/emulators found")) {
                    Thread.sleep(10000);
                    Process validateProcess2 = Runtime.getRuntime()
                            .exec("adb -s " + emulatorsList.getEmulatorDeviceName() + " shell pm list packages "
                                    + emulatorsList.getAppEnv().getAppPackage());
                    r = new BufferedReader(new InputStreamReader(validateProcess2.getInputStream()));

                }

                if (line.contains("error: more than one device/emulator")) {

                    System.err.println("ERROR: SOMETHING WENT WRONG: MORE THAN ONE DEVICE/EMULATOR");
                    System.exit(0);

                }

                if (line.contains("package")) {

                    System.out.println(emulatorsList.getAppEnv().getName() + " app installed.");
                    isAppInstalled = true;

                }
            }

            if (!isAppInstalled) {

                System.err.println(
                        "ERROR: SOMETHING WENT WRONG: FAILED TO INSTALL THE APP ON '" + emulatorsList.getName() + "'");
                System.exit(0);
            }

            r.close();
            p.destroy();

        } catch (IOException | InterruptedException e) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e.printStackTrace(pw);

            System.err.println("ERROR: SOMETHING WENT WRONG DUE TO " + sw.toString());
            System.exit(0);
        }
    }

    private void checkIfEmulatorsRunning() {

        ArrayList<String> runningEmulatorsList = new ArrayList<String>();

        try {

            Process p = Runtime.getRuntime().exec("adb devices");

            BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream()));

            while (true) {
                String line = r.readLine();
                if (line == null) {
                    break;
                }
                if (!line.equalsIgnoreCase("List of devices attached")) {
                    if (line.contains("device") || line.contains("offline")) {
                        if (line.startsWith("emulator")) {
                            runningEmulatorsList.add(line);
                        }

                    }
                }

            }

            p = Runtime.getRuntime().exec("xcrun simctl list devices");

            r = new BufferedReader(new InputStreamReader(p.getInputStream()));

            while (true) {
                String line = r.readLine();
                if (line == null) {
                    break;
                }
                if (line.contains("Booted")) {
                    runningEmulatorsList.add(line);
                }

            }

            r.close();
            p.destroy();

        } catch (IOException e1) {
            StringWriter sw = new StringWriter();
            PrintWriter pw = new PrintWriter(sw);
            e1.printStackTrace(pw);

            System.out.println("ERROR: SOMETHING WENT WRONG DUE TO " + sw.toString());
            System.exit(0);
        }

        if (runningEmulatorsList.size() != 0) {

            System.err.println("ERROR: FOLLOWING EMULATORS ARE STILL RUNNING. PLEASE CLOSE/KILL IF ITS NOT IN USE.");

            for (String string : runningEmulatorsList) {
                System.out.println(string);
            }

            System.exit(0);
        }
    }

    public AppiumDriver<?> startApplication(DeviceProp emulator) throws MalformedURLException {

        switch (emulator.getDeviceType()) {
            case ANDROID:
                return new Android().forAndroid(emulator);

            case IOS:
                return new Apple().forIOS(emulator);

        }
        return null;

    }

    public AppiumDriver<?> startSettings(DeviceProp emulator) throws MalformedURLException {

        switch (emulator.getDeviceType()) {
            case ANDROID:
                return new Android().settings(emulator);

            case IOS:
                return new Apple().settings(emulator);

        }
        return null;

    }

}
