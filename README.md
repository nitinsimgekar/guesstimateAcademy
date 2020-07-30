## For TestLink: ##

* mvn test -Putc -Dexec.args="path to xls file in your local"
* mvn test -Pxlx -Dexec.args="path to xml file in your local"

## For Automation: ##

###	Pre Conditions: ###

* Android home has to be set under environmental variables.
* Screen timeout should be set to max for android/iOS
* Split keyboard is not supported for iOS
* iDevice -> Settings -> Developer - Enable UI Automation should be enabled.
* Add the following to bash_profile

vim ~/.bash_profile
export ANDROID_HOME=~/Library/Android/sdk
export PATH=$ANDROID_HOME/emulator:$ANDROID_HOME/tools:$PATH
source ~/.bash_profile

###	Tips & Troubleshooting: ###
* Start eclipse from cmd line to revoke environmental properties in mac machines

###	Environment setup installations: ###
* brew install ideviceinstaller
* brew install carthage
* npm install -g ios-deploy
* npm install -g deviceconsole
* sudo gem install xcpretty
* brew install --HEAD libimobiledevice -g
* brew cask install android-sdk


###	Webdriver Config: ###
* navigate to /usr/bin/lib/node_modules/appium/node_modules/appium-xcuitest-driver/WebDriverAgent/  (OR)
			/usr/local/lib/node_modules/appium/node_modules/appium-xcuitest-driver/WebDriverAgent/  (OR)
			/usr/local/lib/node_modules/appium/node_modules/appium-webdriveragent/ (For versions above 1.15.0)
* mkdir -p Resources/WebDriverAgent.bundle
* sh ./Scripts/bootstrap.sh -d
* open WebDriverAgent.xcodeproj using xcode
* build webDriverAgent project with sign (change 'product bundle identifier' from 'build settings' if required)

###	execute this in terminal (if required) ###
* if ! grep -q "DEVELOPMENT_TEAM = <teamId>" Configurations/ProjectSettings.xcconfig; then
		echo "DEVELOPMENT_TEAM = <teamId>" >> Configurations/ProjectSettings.xcconfig
	fi
	if ! grep -q "CODE_SIGN_IDENTITY = iPhone Developer" Configurations/ProjectSettings.xcconfig; then
		echo "CODE_SIGN_IDENTITY = iPhone Developer" >> Configurations/ProjectSettings.xcconfig
	fi

###	To verify: ###
* xcodebuild -project WebDriverAgent.xcodeproj -scheme WebDriverAgentRunner -destination 'id=<udid>' test

### Error: Unable to launch WebDriverAgent because of xcodebuild failure ###
* Open the WebDriverAgent.xcodeproj
* Select 'Targets' -> 'WebDriverAgentRunner'
* Open 'Build Phases' -> 'Copy frameworks'
* Click '+' -> add 'RoutingHTTPServer'

###	To install appium: ###
* npm install -g appium
* npm install -g appium-doctor

###	To uninstall appium: ###
* npm uninstall -g appium

###	To upgrade node: ###
* npm update -g

###	To know app package and app activity: ###
* adb shell
* dumpsys window windows | grep -E 'mCurrentFocus'

###	Could not connect to lockdownd: ###
* brew uninstall ideviceinstaller
* brew uninstall libimobiledevice
* brew install --HEAD libimobiledevice
* brew link --overwrite libimobiledevice
* brew install --HEAD  ideviceinstaller
* brew link --overwrite ideviceinstaller
* sudo chmod -R 777 /var/db/lockdown/



###	To Inspect ios elements (by macaca) ###
* npm install macaca-cli -g
* macaca doctor
* npm install app-inspector -g
* open /usr/local/lib/node_modules/app-inspector/node_modules/xctestwd/XCTestWD/XCTestWD.xcodeproj in xcode
* signin and build the project
* xcodebuild -project XCTestWD.xcodeproj -scheme XCTestWDUITests -destination 'platform=iOS,id=<udid>' XCTESTWD_PORT=8001 clean test 
* app-inspector -u <udid>


### Building your iOS app for use in the Appium GUI ###
* https://aaronpresley.com/building-ios-app-use-appium-gui/

### Generate bearer token for jira authentication
*  mvn test -Pgbt -Dexec.arguments="'username','password'"


### Custom Arch type: ###
* Go to the folder where the new project has to be created using terminal
* Execute the cmd: mvn archetype:generate -B -DarchetypeArtifactId=maven-archetype-archetype -DgroupId=<group id> -DartifactId=<artifact id> -Dversion=<version> -Dpackage=<package>
eg:  mvn archetype:generate -B -DarchetypeArtifactId=maven-archetype-archetype -DgroupId=com.guesstimate  -DartifactId=automation  -Dversion=1.0 -Dpackage=com.guesstimate
* Import this project in eclipse
* Open archetype.xml under -- src -- main -- resources -- META-INF -- maven -- archetype.xml
* Rename the archetype.xml to archetype-metadata.xml
* Add the fileset details in the archetype-metadata.xml
* Run as -> maven build in eclipse, a pop up will appear, under goals field enter "install archetype:update-local-catalog"
* Go to comand prompt and run "mvn archetype:crawl"
* Newly archetype xml can be found at .m2/repository folder
* Add this archetype in your eclipse
	Go to ... eclispe -> preferances -> maven - > archetype
	Click on "add local catelog" button
	Enter newly created archetype xml file path and description
	Click ok

* Now you will see this archetype when you create a new maven project.

### Generate Actions Methods ###
* Follow the template of:
* add ```edit``` before or after the variable for which needs to enter text
* add ```lbl``` before or after the variable for which needs we need to read text
* add ```btn``` before or after the variable for which needs to click

### For More Details ###
Contact 
<br>Nitin Simgekar
<br><b>+91 9160908000</b>
<br>nitinsimgekar@gmail.com
