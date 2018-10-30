package Appium_Studio;

import static java.time.Duration.ofSeconds;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.logging.Level;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.ScreenOrientation;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.logging.LogEntry;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.google.common.base.Stopwatch;
import io.appium.java_client.TouchAction;

import io.appium.java_client.android.AndroidKeyCode;
import io.appium.java_client.android.Connection;
import io.appium.java_client.ios.IOSDriver;
import io.appium.java_client.remote.AndroidMobileCapabilityType;
import io.appium.java_client.remote.IOSMobileCapabilityType;
import io.appium.java_client.remote.MobileCapabilityType;

public class method {
	// int systemPort = 8200;// System port
	static int port = 4725;// Appium port
	static int StartPort = port;
	static int device_timeout = 1200;// 120sec
	int command_timeout = 30;// 30sec
	LoadExpectResult ExpectResult = new LoadExpectResult();
	static LoadTestCase TestCase = new LoadTestCase();
	static ArrayList<IOSDriver> driver = new ArrayList<IOSDriver>();// �x�s
																	// IOSDriver
																	// ��ArrayList
	// driverBK�ƥ�driver ��ArrayList�A�ت����^�_�쥻driver
	// list����driver��T�A�]��Command�X����|�]�w��driver list��null�A�G�����U�Ӯרҫe�ݳ]�w�^��driver ��T
	static ArrayList<IOSDriver> driverBK = new ArrayList<IOSDriver>();
	static String CaseErrorList[][] = new String[TestCase.CaseList.size()][TestCase.DeviceInformation.deviceName
			.size()];// �����U�רҩ�U�˸m�����O���G (2���}�C)CaseErrorList[CaseList][Devices]
	static IOSDriver IOSDriver = null;
	static String ErrorList[] = new String[TestCase.DeviceInformation.deviceName.size()];// �����U�˸m�����O���G
	WebDriverWait[] wait = new WebDriverWait[TestCase.DeviceInformation.deviceName.size()];
	static XSSFWorkbook workBook;
	static String appElemnt;// APP����W��
	static String appInput;// ��J��
	static String appInputXpath;// ��J�Ȫ�Xpath�榡
	static String toElemnt;// APP����W��
	static int startx, starty, endx, endy;// Swipe���ʮy��
	static int iterative;// �e���ưʦ���
	static String scroll;// �e�����ʤ�V
	static String appElemntarray;// �j�M���h����������
	static String checkVerifyText;// �T�{����QuitAPP�e�O�_����LVerifyText
	static String switchWiFi;// �Ұ�wifi������wifi
	String element[] = new String[driver.size()];
	static int CurrentCaseNumber = -1;// �ثe�����ĴX�Ӵ��ծצC
	static Boolean CommandError = true;// �P�w���檺���O�O�_�X�{���~�Fture�����T�Ffalse�����~
	static int CurrentErrorDevice = 0;// �έp�ثe�X�����]�Ƽƶq
	XSSFSheet Sheet;
	static long totaltime;// �έp�Ҧ��רҴ��ծɶ�
	static int location;// ����driver arraylist��null��index
	static int CurrentCase;

	public static void main(String[] args) throws IOException, NoSuchMethodException, SecurityException,
			IllegalAccessException, IllegalArgumentException, InvocationTargetException, InstantiationException {
		initial();
		CreateAppiumSession();// �إ� Appium �q�D
		invokeFunction();
		EndAppiumSession();// ���_ Appium �q�D
		System.out.println("���յ���!!!" + "(" + totaltime + " s)");
		Process proc = Runtime.getRuntime().exec("explorer C:\\TUTK_QA_TestTool\\TestReport");// �}��TestReport��Ƨ�

	}

	public static void initial() {// ��l��CaseErrorList�x�}
		for (int i = 0; i < CaseErrorList.length; i++) {
			for (int j = 0; j < CaseErrorList[i].length; j++) {
				CaseErrorList[i][j] = "";// ��J�Ŧr��A�קK���ȮɡA�X�{���~
			}
		}
		for (int i = 0; i < TestCase.DeviceInformation.deviceName.size(); i++) {
			ErrorList[i] = "";
		}
	}

	public static void invokeFunction() throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException {
		Object obj = new method();
		Class c = obj.getClass();
		String methodName = null;

		for (CurrentCase = 0; CurrentCase < TestCase.StepList.size(); CurrentCase++) {
			Stopwatch timer = Stopwatch.createStarted();// �}�l�p��
			System.out.println("[info] CaseName:|" + TestCase.CaseList.get(CurrentCase).toString() + "|");
			CommandError = true;// �w�]CommandError��True
			CurrentErrorDevice = 0;// �w�]�X�����]�ƼƬ�0�x
			for (int CurrentCaseStep = 0; CurrentCaseStep < TestCase.StepList.get(CurrentCase)
					.size(); CurrentCaseStep++) {
				if (!CommandError & CurrentErrorDevice == TestCase.DeviceInformation.deviceName.size()) {
					break;// �Y�ثe���ծרҥX�{CommandError=false�A�h���X�ثe�רҨð���U�@�Ӯר�
				}
				switch (TestCase.StepList.get(CurrentCase).get(CurrentCaseStep).toString()) {

				case "LaunchAPP":
					methodName = "LaunchAPP";
					break;

//				case "Byid_SendKey":
//					methodName = "Byid_SendKey";
//					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
//					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
//					CurrentCaseStep = CurrentCaseStep + 2;
//					break;
//
//				case "Byid_Click":
//					methodName = "Byid_Click";
//					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
//					CurrentCaseStep = CurrentCaseStep + 1;
//					break;
//
//				case "Byid_Swipe":
//					methodName = "Byid_Swipe";
//					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
//					toElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
//					CurrentCaseStep = CurrentCaseStep + 2;
//					break;

				case "ByXpath_SendKey":
					methodName = "ByXpath_SendKey";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

//				case "Byid_Clear":
//					methodName = "Byid_Clear";
//					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
//					CurrentCaseStep = CurrentCaseStep + 1;
//					break;

				case "ByXpath_Clear":
					methodName = "ByXpath_Clear";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Click":
					methodName = "ByXpath_Click";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Swipe":
					methodName = "ByXpath_Swipe";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					toElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

//				case "Byid_Wait":
//					methodName = "Byid_Wait";
//					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
//					CurrentCaseStep = CurrentCaseStep + 1;
//					break;

				case "ByXpath_Wait":
					methodName = "ByXpath_Wait";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "HideKeyboard":
					methodName = "HideKeyboard";
					break;

				case "Byid_VerifyText":
					methodName = "Byid_VerifyText";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_VerifyText":
					methodName = "ByXpath_VerifyText";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Sleep":
					methodName = "Sleep";
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ScreenShot":
					methodName = "ScreenShot";
					break;

				case "Orientation":
					methodName = "Orientation";
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Swipe":
					methodName = "Swipe";
					startx = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1));
					starty = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2));
					endx = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3));
					endy = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 4));
					iterative = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 5));
					CurrentCaseStep = CurrentCaseStep + 5;
					break;

				case "ByXpath_Swipe_Vertical":
					methodName = "ByXpath_Swipe_Vertical";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					scroll = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					iterative = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3));
					CurrentCaseStep = CurrentCaseStep + 3;
					break;

				case "ByXpath_Swipe_Horizontal":
					methodName = "ByXpath_Swipe_Horizontal";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					scroll = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					iterative = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3));
					CurrentCaseStep = CurrentCaseStep + 3;
					break;

				case "ByXpath_Swipe_FindText_Click_Android":
					methodName = "ByXpath_Swipe_FindText_Click_Android";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					scroll = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					appElemntarray = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3);
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 4);
					appInputXpath = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 5);
					CurrentCaseStep = CurrentCaseStep + 5;
					break;

//				case "Byid_LongPress":
//					methodName = "Byid_LongPress";
//					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
//					CurrentCaseStep = CurrentCaseStep + 1;
//					break;

				case "ByXpath_LongPress":
					methodName = "ByXpath_LongPress";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

//				case "Byid_invisibility":
//					methodName = "Byid_invisibility";
//					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
//					CurrentCaseStep = CurrentCaseStep + 1;
//					break;

				case "ByXpath_invisibility":
					methodName = "ByXpath_invisibility";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

//				case "ByXpath_Swipe_FindText_Click_iOS":
//					methodName = "ByXpath_Swipe_FindText_Click_iOS";
//					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
//					scroll = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
//					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3);
//					CurrentCaseStep = CurrentCaseStep + 3;
//					break;

				case "Back":
					methodName = "Back";
					break;

				case "Home":
					methodName = "Home";
					break;

				case "Power":
					methodName = "Power";
					break;

				case "Recent":
					methodName = "Recent";
					break;

				case "ResetAPP":
					methodName = "ResetAPP";
					break;

				case "QuitAPP":
					methodName = "QuitAPP";
					checkVerifyText = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep - 2).toString();
					break;

				case "WiFi":
					methodName = "WiFi";
					switchWiFi = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				}

				Method method;
				method = c.getMethod(methodName, new Class[] {});
				method.invoke(c.newInstance());

			}
			System.out.println("[info] Time:|" + timer.stop() + "|");
			totaltime += timer.elapsed(TimeUnit.SECONDS);
			System.out.println("");
			ResetDriverArrayList();// reset driver arraylist
		}
		// System.out.println("���յ���!!!" + "(" + totaltime + " s)");
	}

	public static void ErrorCheck(Object... elements) throws IOException {
		DateFormat df = new SimpleDateFormat("MMM dd, yyyy h:mm:ss a");
		Date today = Calendar.getInstance().getTime();
		String reportDate = df.format(today);

		// APP�e���W�䤣����w����
		if ((elements.length - 1) > 1) {
			String APPElement = "";
			int i = 0;
			for (Object element : elements) {
				APPElement = APPElement + element;
				if (i != (elements.length - 2)) {// �P�_�O�_�B�z��˼ƲĤG��element
					APPElement = APPElement + " or ";// �զ�" A���� or B���� or
														// C����"�r��
				} else {
					break;// �O�˼ƲĤG��break for
				}
				i++;
			}
			System.err.print("[Error] Can't find " + APPElement + " on screen.");
		} else {
			int j = 0;
			for (Object element : elements) {
				if (j != (elements.length - 1)) {// �P�_�O�_�B�z��̫�@��element

					if (element.equals("On")) {// ok
						System.err.print("[Error] Can't turn on WiFi. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Off")) {// ok
						System.err.print("[Error] Can't turn off WiFi. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("HideKeyboard")) {// ok
						System.out.print("[Error] Can't hide keyboard. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Sleep")) {// ok
						System.err.print("[Error] Fail to sleep. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("ScreenShot")) {// ok
						System.err.print("[Error] Fail to ScreenShot. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Portrait")) {// ok
						System.err.print("[Error] Can't rotate to portrait. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Landscape")) {// ok
						System.err.print("[Error] Can't rotate to landscape. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("EndAppiumSession")) {// ok
						System.err.print("[Error] Can't end session. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("QuitAPP")) {// ok
						System.err.print("[Error] Can't close APP. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("ResetAPP")) {// ok
						System.err.print("[Error] Can't reset APP. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("CeateAppiumSession")) {// ok
						System.out.print("[Error] Can't create new session. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("LaunchAPP")) {// ok
						System.err.print("[Error] Can't launch APP. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("BACK") || element.equals("Home") || element.equals("Power")
							|| element.equals("Recent")) {// ok
						System.err.print("[Error] Can't execute " + element + " button. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Customized")) {// ok
						System.err.print("[Error] Can't execute Customized_Method. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Swipe")) {// ok

						System.err.print("[Error] Can't swipe " + "(" + startx + "," + starty + ")" + " to (" + endx
								+ "," + endy + "). (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else {
						System.err.print("[Error] Can't find " + element + " on screen. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					}
				}
				j++;
			}
		}
		System.err.println(" " + reportDate);
		String FilePath = MakeErrorFolder(Integer.parseInt(String.valueOf(elements[elements.length - 1])));// �إߦU�רҤ���ʸ˸m��Ƨ��s��log��T��Screenshot��T
		// logcat(FilePath, Integer.parseInt(String.valueOf(elements[elements.length -
		// 1])));// ����log
		ErrorScreenShot(FilePath, Integer.parseInt(String.valueOf(elements[elements.length - 1])));// Screenshot
																									// Error�e��
		ErrorList[Integer.parseInt(String.valueOf(elements[elements.length - 1]))] = "Error";// �x�s��i�x�]��command���ѵ��G
		CaseErrorList[CurrentCaseNumber] = ErrorList;// �x�s��i�x�]�ư����CurrentCaseNumber�ӮרҤ�command���ѵ��G
		CommandError = false;// �Y�䤣����w����A�h�]�wCommandError=false
		driver.set(Integer.parseInt(String.valueOf(elements[elements.length - 1])), null);// �N�X����driver�]�w��null
		CurrentErrorDevice++;// �έp�X�����]�Ƽ�
	}

//	public static void logcat(String FilePath, int DeviceNum) throws IOException {
//		// ����log
//		// System.out.println(
//		// "[info] Executing:|Saving device log...|" +
//		// TestCase.DeviceInformation.deviceName.get(DeviceNum) + "|");
//		DateFormat df = new SimpleDateFormat("yyyy_MM_dd_HH-mm-ss");
//		Date today = Calendar.getInstance().getTime();
//		String reportDate = df.format(today);
//		List<LogEntry> logEntries = driver.get(DeviceNum).manage().logs().get("logcat").filter(Level.ALL);
//		try {
//			FileWriter fw = new FileWriter(FilePath + reportDate + "_log" + ".txt");
//			for (int i = 0; i < logEntries.size(); i++) {
//				fw.write(logEntries.get(i).toString() + "\n");
//			}
//			fw.flush();
//			fw.close();
//			System.out.println("[info] Executing:|Saving device log - Done.|"
//					+ TestCase.DeviceInformation.deviceName.get(DeviceNum) + "|");
//		} catch (Exception e) {
//			System.err.println("[Error] Executing:|Fail to save device log.|"
//					+ TestCase.DeviceInformation.deviceName.get(DeviceNum) + "|");
//		}
//	}

	public static void ErrorScreenShot(String FilePath, int DeviceNum) {
		try {
			System.out.println("[info] Executing:|Taking a screenshot of error.|"
					+ TestCase.DeviceInformation.deviceName.get(DeviceNum) + "|");
			DateFormat df = new SimpleDateFormat("yyyy_MM_dd_HH-mm-ss");
			Date today = Calendar.getInstance().getTime();
			String reportDate = df.format(today);
			File screenShotFile = (File) driver.get(DeviceNum).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(screenShotFile, new File(FilePath + reportDate + ".jpg"));
		} catch (IOException e) {
			System.err.println("[Error] Executing:|Fail to ErrorScreenShot.|"
					+ TestCase.DeviceInformation.deviceName.get(DeviceNum) + "|");
		}
	}

	public static String MakeErrorFolder(int i) {
		// ��Ƨ����c
		// C:\TUTK_QA_TestTool\TestReport\appPackage\CaseName\DeviceUdid\log\
		String filePath = "C:\\TUTK_QA_TestTool\\TestReport\\" + TestCase.DeviceInformation.bundleID.toString() + "\\"
				+ TestCase.CaseList.get(CurrentCase).toString() + "\\" + TestCase.DeviceInformation.deviceName.get(i)
				+ "\\log\\";
		File file = new File(filePath);
		if (!file.exists()) {
			file.mkdirs();
		}
		return filePath;
	}

	public static void ResetDriverArrayList() {
		// ��@Case������A�Ndriver arraylist��null�������^�_���쥻driver���(��driverBK
		// arraylist�פJdriver arraylist)
		boolean arraycheck;
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) == null) {
				location = i;// ����driver arraylist��null index��}
				for (int k = 0; k < driverBK.size(); k++) {
					arraycheck = false;
					for (int l = 0; l < driver.size(); l++) {
						// �T�{driverBK arraylist��ƬO�_�s�b��driver arraylist��
						if (driverBK.get(k).equals(driver.get(l))) {
							arraycheck = false;
							break;
						} else {
							arraycheck = true;
						}
					}
					if (arraycheck) {
						driver.set(location, driverBK.get(k));// �л\driver
																// arraylist����null
					}
				}
			}
		}

	}


	public void ByXpath_VerifyText() throws IOException {
		boolean result[] = new boolean[driver.size()];// �����wBoolean�ȡA�w�]��False
		boolean ErrorResult[] = new boolean[driver.size()];

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_VerifyText|" + appElemnt + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					element[i] = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)))
							.getText();
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
					// System.err.println("[Error] Can't find " + appElemnt);
					// element[i] = "ERROR";// �䤣��Ӫ���A�^��Error
					// driver.set(i, null);
					// CurrentErrorDevice++;// �έp�]�X�����]�Ƽ�
				}

				if (element[i].equals("ERROR")) {
					ErrorResult[i] = true;

				} else {
					// �^�Ǵ��ծרҲM�檺�W�ٵ�ExpectResult.LoadExpectResult�A�æs����浲�G��ResultList�M��
					ExpectResult.LoadExpectResult(TestCase.CaseList.get(CurrentCaseNumber).toString());
					for (int j = 0; j < ExpectResult.ResultList.size(); j++) {
						if (element[i].equals(ExpectResult.ResultList.get(j)) == true) {
							result[i] = true;
							break;
						} else {
							result[i] = false;
						}
					}
				}
			}
		}
		SubMethod_Result(ErrorResult, result);

	}

	public void ByXpath_Wait() throws IOException {

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_Wait|" + appElemnt + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.presenceOfElementLocated(By.xpath(appElemnt)));
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void ByXpath_SendKey() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_SendKey|" + appElemnt + "|" + appInput + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					// wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).clear();
					wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)))
							.sendKeys(appInput);
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
					HideKeyboard();
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void ByXpath_Click() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_Click|" + appElemnt + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).click();
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void ByXpath_Clear() throws IOException {

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Byid_Clear|" + appElemnt + "|Clear|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).clear();
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
					HideKeyboard();
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}

	}

	public void HideKeyboard() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {

					System.out.println(
							"[info] Executing:|HideKeyboard|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					driver.get(i).hideKeyboard();
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;

				} catch (Exception e) {
					ErrorCheck("HideKeyboard", i);
				}
			}
		}
	}

	public void Sleep() throws IOException {
		// String NewString = "";// �s�r��
		// char[] r = { '.' };// �p���I�r��
		// char[] c = appInput.toCharArray();// �N�r���ন�r���}�C
		// for (int i = 0; i < c.length; i++) {
		// if (c[i] != r[0]) {// �P�_�r���O�_���p���I
		// NewString = NewString + c[i];// �_�A�N�r���զX���s�r��
		// } else {
		// break;// �O�A���X�j��
		// }
		// }
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {

					System.out.println("[info] Executing:|Sleep|" + appInput + " second..." + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					Thread.sleep((long) (Float.valueOf(appInput) * 1000));
					// Thread.sleep(Integer.valueOf(NewString) * 1000);//
					// �N�r���ন���
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;

				} catch (Exception e) {
					ErrorCheck("Sleep", i);
				}
			}
		}
	}

	public void ScreenShot() throws IOException {
		DateFormat df = new SimpleDateFormat("yyyy_MM_dd_HH-mm-ss");
		Date today = Calendar.getInstance().getTime();
		String reportDate = df.format(today);

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					File screenShotFile = (File) driver.get(i).getScreenshotAs(OutputType.FILE);
					System.out.println(
							"[info] Executing:|ScreenShot|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					FileUtils.copyFile(screenShotFile, new File("C:\\TUTK_QA_TestTool\\TestReport\\" + "#" + i + 1
							+ TestCase.CaseList.get(CurrentCaseNumber) + "_" + reportDate + ".jpg"));
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (IOException e) {
					ErrorCheck("ScreenShot", i);
				}
			}
		}
	}

	public void Orientation() throws IOException {

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					if (appInput.equals("Landscape")) {
						System.out.println("[info] Executing:|Orientation|Landscape|"
								+ TestCase.DeviceInformation.deviceName.get(i) + "|");
						driver.get(i).rotate(ScreenOrientation.LANDSCAPE);
					} else if (appInput.equals("Portrait")) {
						System.out.println("[info] Executing:|Orientation|Portrait|");
						driver.get(i).rotate(ScreenOrientation.PORTRAIT);
					}
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					if (appInput.equals("Landscape")) {
						ErrorCheck("Landscape", i);
					} else {
						ErrorCheck("Portrait", i);
					}
				}
			}
		}
	}

	public static void EndAppiumSession() throws IOException {
		for (int j = 0; j < driver.size(); j++) {
			if (driver.get(j) != null) {
				try {
					System.out.println("[info] Executing:|End Session|Server Port:" + (port - 2 - (j * 2)) + "|");
					driver.get(j).quit();
					ErrorList[j] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck("EndAppiumSession", j);
				}
			}
		}
	}

	public void QuitAPP() throws IOException {
		boolean state[] = new boolean[driver.size()];
		for (int j = 0; j < driver.size(); j++) {
			if (driver.get(j) != null) {
				try {
					System.out
							.println("[info] Executing:|QuitAPP|" + TestCase.DeviceInformation.deviceName.get(j) + "|");
					driver.get(j).closeApp();

					// ���槹QuitAPP��A�N���Case�y�{�ާ@���`�����A�]���|�ATestReport��JPass
					// ;���Y���e����LVerifyText�ʧ@��A�]�|�bTestReport��J��Case���GPass/Fail/Error�A�]�����B�ݧP�_�O�_����LVerify�A�קKVerify���G�Q�걼�A�p�U�P�_��
					for (int i = 0; i < TestCase.StepList.get(CurrentCaseNumber).size(); i++) {
						if (TestCase.StepList.get(CurrentCaseNumber).get(i).equals("Byid_VerifyText")
								|| TestCase.StepList.get(CurrentCaseNumber).get(i).equals("ByXpath_VerifyText")
								|| TestCase.StepList.get(CurrentCaseNumber).get(i).equals("Byid_VerifyRadioButton")
								|| TestCase.StepList.get(CurrentCaseNumber).get(i)
										.equals("ByXpath_VerifyRadioButton")) {
							state[j] = true;// true�N����Verify
							break;
						}
					}
					if (!state[j]) {// �Ystate[j]=false���case������Verify���O�A�G����H�U�y�{

						try {
							workBook = new XSSFWorkbook(
									new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\TestReport_iOS.xlsm"));
						} catch (Exception e) {
							System.err.println(
									"[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport_iOS.xlsm");
						}

						if (TestCase.DeviceInformation.deviceName.get(j).toString().length() > 20) {// Excel�u�@��W�ٳ̱`31�r���]�A�G�ݧP�_UDID���׬O�_�j��31
							char[] NewUdid = new char[20];// �]�ݥ]�t_TestReport�r��(�@11�r��)�A�G�]�w20��r���}�C(31-11)
							TestCase.DeviceInformation.deviceName.get(j).toString().getChars(0, 20, NewUdid, 0);// ���XUDID�e20�r����NewUdid
							Sheet = workBook.getSheet(String.valueOf(NewUdid) + "_TestReport");// �ھ�NewUdid�A���w�Y�x�˸m��TestReport
																								// sheet
						} else {
							Sheet = workBook
									.getSheet(TestCase.DeviceInformation.deviceName.get(j).toString() + "_TestReport");// ���w�Y�x�˸m��TestReport
																														// sheet
						}
						if (CaseErrorList[CurrentCaseNumber][j].equals("Pass")) {// ���XCaseErrorList����CurrentCaseNumber�Ӵ���������i�x��ʸ˸m�����G
							Sheet.getRow(CurrentCaseNumber + 1).getCell(1).setCellValue("Pass");// ��J��i�x��ʸ˸m����CurrentCaseNumber�Ӵ������GPass
						}
						// ����g�JExcel�᪺�s�ɰʧ@
						try {
							FileOutputStream out = new FileOutputStream(
									new File("C:\\TUTK_QA_TestTool\\TestReport\\TestReport_iOS.xlsm"));
							workBook.write(out);
							out.close();
							workBook.close();
						} catch (Exception e) {
							System.err.println(
									"[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport_iOS.xlsm");
						}
					}

				} catch (Exception e) {
					ErrorCheck("QuitAPP", j);
				}
			}
		}
	}

	public void ResetAPP() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println(
							"[info] Executing:|ResetAPP|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					driver.get(i).resetApp();
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck("ResetAPP", i);
				}
			}
		}
	}

	public static void CreateAppiumSession() throws IOException {
		DesiredCapabilities cap[] = new DesiredCapabilities[TestCase.DeviceInformation.deviceName.size()];
		driver.clear();// �M��driver arraylist

		for (int i = 0; i < TestCase.DeviceInformation.deviceName.size(); i++) {
			cap[i] = new DesiredCapabilities();
		}

		for (int i = 0; i < TestCase.DeviceInformation.deviceName.size(); i++) {
			System.out.println("DEVICE_NAME:" + TestCase.DeviceInformation.deviceName.get(i).toString());
			cap[i].setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, device_timeout);
			cap[i].setCapability(MobileCapabilityType.DEVICE_NAME, TestCase.DeviceInformation.deviceName.get(i));
			cap[i].setCapability(MobileCapabilityType.UDID, TestCase.DeviceInformation.deviceName.get(i));
			cap[i].setCapability(MobileCapabilityType.PLATFORM_VERSION,
					TestCase.DeviceInformation.platformVersion.get(i));
			cap[i].setCapability(IOSMobileCapabilityType.BUNDLE_ID, TestCase.DeviceInformation.bundleID);
			cap[i].setCapability(MobileCapabilityType.NO_RESET, TestCase.DeviceInformation.ResetAPP);// �C����@Case����e�A�O�_�M��APP�֨����;�O��true;�_��false
			cap[i].setCapability("autoLaunch", false);
			// cap[i].setCapability(MobileCapabilityType.AUTOMATION_NAME, "XCUITest");

			try {

				System.out.println("[info] Executing:|Create New Session|Server Port:" + port + "|");
				IOSDriver = new IOSDriver<>(new URL("http://127.0.0.1:" + port + "/wd/hub"), cap[i]);
				driverBK.add(IOSDriver);
				driver.add(IOSDriver);
			} catch (Exception e) {
				ErrorCheck("CeateAppiumSession", i);
			}
			port = port + 2;
			// systemPort = systemPort + 4;
		}
		System.out.println("");
	}

	public void LaunchAPP() throws IOException {
		CurrentCaseNumber = CurrentCaseNumber + 1;
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|LaunchAPP|" + TestCase.DeviceInformation.bundleID + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					driver.get(i).launchApp();
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck("LaunchAPP", i);
				}
			}
		}

	}

	public void Home() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Home|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					driver.get(i).executeScript("seetest:client.deviceAction(\"Home\")");
					// driver.get(i).pressKeyCode(AndroidKeyCode.HOME);
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck("Home", i);
				}
			}
		}
	}

	public void Power() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Power|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					// driver.get(i).pressKeyCode(AndroidKeyCode.KEYCODE_POWER);
					driver.get(i).executeScript("seetest:client.deviceAction(\"Power\")");
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck("Power", i);
				}
			}
		}
	}

	public void Recent() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out
							.println("[info] Executing:|Recent|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					// driver.get(i).pressKeyCode(AndroidKeyCode.KEYCODE_APP_SWITCH);
					driver.get(i).executeScript("seetest:client.deviceAction(\"Recent Apps\")");
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck("Recent", i);
				}
			}
		}
	}

	public void ByXpath_invisibility() throws IOException {
		for (int i = 0; i < driver.size(); i++) {

			if (driver.get(i) != null) {

				try {
					System.out.println("[info] Executing:|ByXpath_invisibility|" + appElemnt + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(appElemnt)));
					ErrorList[i] = "Pass";// �x�s��i�x�]��command���G�A���\����Click�h�A�s�JPass||�|��
											// ���N1�GErrorList=[Pass];���N2�GErrorList=[Pass,Pass]
					CaseErrorList[CurrentCaseNumber] = ErrorList;// �x�s��i�x�]�ư����CurrentCaseNumber�ӮרҤ�command���G||�|��
																	// ���N1�GCaseErrorList=[[Pass]];���N2�GCaseErrorList=[[Pass,Pass]]
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void ByXpath_LongPress() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {

				try {
					System.out.println("[info] Executing:|ByXpath_LongPress|" + appElemnt + "|");
					TouchAction t = new TouchAction(driver.get(i));
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					t.longPress(wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))))
							.perform();

					ErrorList[i] = "Pass";// �x�s��i�x�]��command���G�A���\����Click�h�A�s�JPass||�|��
											// ���N1�GErrorList=[Pass];���N2�GErrorList=[Pass,Pass]
					CaseErrorList[CurrentCaseNumber] = ErrorList;// �x�s��i�x�]�ư����CurrentCaseNumber�ӮרҤ�command���G||�|��
																	// ���N1�GCaseErrorList=[[Pass]];���N2�GCaseErrorList=[[Pass,Pass]]
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void ByXpath_Swipe() throws IOException {
		Point p1, p2;// p1 ���_�I;p2�����I

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_Swipe|" + appElemnt + "|" + toElemnt + "|");
					TouchAction t = new TouchAction(driver.get(i));
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					WebElement ele2 = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(toElemnt)));
					WebElement ele1 = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)));

					t.press(ele1.getLocation().x, ele1.getLocation().y).waitAction(1)
							.moveTo(ele2.getLocation().x, ele2.getLocation().y).release().perform();
					// t.press(ele1).waitAction(WaitOptions.waitOptions(ofSeconds(1))).moveTo(ele2).release().perform();
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck(toElemnt, appElemnt, i);
				}
			}
		}
	}

	public void Swipe() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				for (int j = 0; j < iterative; j++) {
					try {
						System.out.println(
								"[info] Executing:|Swipe|(" + startx + "," + starty + ")|(" + endx + "," + endy + ")|");
						TouchAction t = new TouchAction(driver.get(i));

						t.press(startx, starty).waitAction(1000).moveTo(endx, endy).release().perform();
						ErrorList[i] = "Pass";
						CaseErrorList[CurrentCaseNumber] = ErrorList;
					} catch (Exception e) {
						ErrorCheck("Swipe", i);
						break;// �X����A���}iterative�^��
					}
				}
			}
		}
	}

	public void ByXpath_Swipe_Vertical() throws IOException {
		Point p;// ����y��
		Dimension s;// ����j�p
		WebElement e;

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_Swipe_Vertical|" + appElemnt + "|" + scroll + "|"
							+ iterative + "|");
					TouchAction t = new TouchAction(driver.get(i));
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					e = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)));
					s = e.getSize();
					p = e.getLocation();
					int errorX = (int) Math.round(s.width * 0.01);
					int errorY = (int) Math.round(s.height * 0.01);
					for (int j = 0; j < iterative; j++) {
						if (scroll.equals("DOWN")) {// �e���V�U����

							t.press(p.x + errorX, p.y + s.height - errorY).waitAction(1000)
									.moveTo(p.x + errorX, p.y + errorY).release().perform();

						} else if (scroll.equals("UP")) {// �e���V�W����
							t.press(p.x + errorX, p.y + errorY).waitAction(1000)
									.moveTo(p.x + errorX, p.y + s.height - errorY).release().perform();
						}
					}
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception w) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void ByXpath_Swipe_Horizontal() throws IOException {
		Point p;// ����y��
		Dimension s;// ����j�p
		WebElement e;
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_Swipe_Horizontal|" + appElemnt + "|" + scroll + "|"
							+ iterative + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					TouchAction t = new TouchAction(driver.get(i));
					e = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)));

					s = e.getSize();
					p = e.getLocation();
					int errorX = (int) Math.round(s.getWidth() * 0.01);
					int errorY = (int) Math.round(s.getHeight() * 0.01);
					for (int j = 0; j < iterative; j++) {
						if (scroll.equals("RIGHT")) {// �e���V�k���� (�[�ݵe�����褺�e)

							t.press(p.x + errorX, p.y + errorY).waitAction(1000)
									.moveTo(p.x + s.width - errorX, p.y + errorY).release().perform();
						} else if (scroll.equals("LEFT")) {// �e���V������ (�[�ݵe���k�褺�e)

							t.press(p.x + s.width - errorX, p.y + errorY).waitAction(1000)
									.moveTo(p.x + errorX, p.y + errorY).release().perform();
						}
					}
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception w) {
					ErrorCheck(appElemnt, i);
				}
			}

		}
	}

	/*
	 * �W�U�H���ư�n�� public void Swipe() { Random rand = new Random(); boolean items[] =
	 * { true, false }; for (int i = 0; i < driver.size(); i++) { for (int j = 0; j
	 * < iterative; j++) { if (items[rand.nextInt(items.size())]) {
	 * driver.get(i).swipe(startx, starty, endx, endy, 500); }else{
	 * driver.get(i).swipe(endx, endy, startx , starty , 500); } } } }
	 * 
	 */

	public void SubMethod_Result(boolean ErrorResult[], boolean result[]) {
		// �}��Excel
		try {
			workBook = new XSSFWorkbook(new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\TestReport_iOS.xlsm"));
		} catch (Exception e) {
			System.err.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport_iOS.xlsm");
		}
		for (int i = 0; i < driver.size(); i++) {

			if (TestCase.DeviceInformation.deviceName.get(i).toString().length() > 20) {// Excel�u�@��W�ٳ̱`31�r���]�A�G�ݧP�_UDID���׬O�_�j��31
				char[] NewUdid = new char[20];// �]�ݥ]�t_TestReport�r��(�@11�r��)�A�G�]�w20��r���}�C(31-11)
				TestCase.DeviceInformation.deviceName.get(i).toString().getChars(0, 20, NewUdid, 0);// ���XUDID�e20�r����NewUdid
				Sheet = workBook.getSheet(String.valueOf(NewUdid) + "_TestReport");// �ھ�NewUdid�A���w�Y�x�˸m��TestReport
																					// sheet
			} else {
				Sheet = workBook.getSheet(TestCase.DeviceInformation.deviceName.get(i).toString() + "_TestReport");// ���w�Y�x�˸m��TestReport
																													// sheet
			}

			if (ErrorResult[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Error");
			} else if (result[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Pass");
			} else if (result[i] == false) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Fail");
			}
		}
		// ����g�JExcel�᪺�s�ɰʧ@
		try {
			FileOutputStream out = new FileOutputStream(
					new File("C:\\TUTK_QA_TestTool\\TestReport\\TestReport_iOS.xlsm"));
			workBook.write(out);
			out.close();
			workBook.close();
		} catch (Exception e) {
			System.err.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport_iOS.xlsm");
		}
	}
}
