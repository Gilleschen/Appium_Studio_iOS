package Appium_Studio;

import java.io.File;
import java.io.FileWriter;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.logging.Level;

import org.openqa.selenium.logging.LogEntry;
import org.openqa.selenium.remote.DesiredCapabilities;

import io.appium.java_client.ios.IOSDriver;
import io.appium.java_client.ios.IOSElement;
import io.appium.java_client.remote.IOSMobileCapabilityType;
import io.appium.java_client.remote.MobileCapabilityType;

public class test {

	public static void main(String[] args) throws MalformedURLException, InterruptedException {
		IOSDriver<IOSElement> driver = null;
		DesiredCapabilities cap = new DesiredCapabilities();
		cap.setCapability(MobileCapabilityType.UDID, "7e9529ce5c4ee1c7de7598e7ac26e25e2d1f8700");
		cap.setCapability(MobileCapabilityType.PLATFORM_VERSION, "10.1.1");
		cap.setCapability(MobileCapabilityType.NO_RESET, true);
		cap.setCapability(MobileCapabilityType.AUTOMATION_NAME, "XCUITest");
		cap.setCapability(IOSMobileCapabilityType.BUNDLE_ID, "com.tutk.kalayapp.jenkins");
		driver = new IOSDriver<>(new URL("http://localhost:4725/wd/hub"), cap);

		Thread.sleep(8000);

		String filePath = "C:\\TUTK_QA_TestTool\\TestReport\\" + "com.tutk.kalayapp.jenkins" + "\\" + "log_test" + "\\"
				+ "7e9529ce5c4ee1c7de7598e7ac26e25e2d1f8700" + "\\log\\";

		File file = new File(filePath);
		if (!file.exists()) {
			file.mkdirs();
		}

		DateFormat df = new SimpleDateFormat("yyyy_MM_dd_HH-mm-ss");
		Date today = Calendar.getInstance().getTime();
		String reportDate = df.format(today);
		
		List<LogEntry> logEntries = driver.manage().logs().get("logcat").filter(Level.ALL);
		try {
			FileWriter fw = new FileWriter(filePath + reportDate + "_log" + ".txt");
			for (int i = 0; i < logEntries.size(); i++) {
				fw.write(logEntries.get(i).toString() + "\n");
			}
			fw.flush();
			fw.close();

		} catch (Exception e) {
			;
		}

	}

}
