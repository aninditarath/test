package testCases;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import ComExcel.Read;
import operation.*;

import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.springframework.core.io.FileSystemResource;
import org.springframework.mail.javamail.JavaMailSenderImpl;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.mail.javamail.MimeMessagePreparator;
import org.testng.ITestResult;
import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.DataProvider;








//import com.relevantcodes.extentreports.ExtentReports;
//import com.relevantcodes.extentreports.ExtentTest;
//import com.relevantcodes.extentreports.LogStatus;
import com.aventstack.extentreports.ExtentReporter;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.*;
//import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.reporter.*;
import com.aventstack.extentreports.reporter.configuration.Theme;
import com.relevantcodes.extentreports.LogStatus;
import com.sun.jna.platform.win32.OaIdl.DATE;

import util.Log;

import java.io.File;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import javax.mail.MessagingException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

public class TestNG {

	public static ExtentReports extent;
	public static ExtentTest test;
	public ITestResult it;
	public ExtentTest parent;
	public ExtentReporter reporter;
	public ExtentHtmlReporter htmlReporter;
	public Date date;
	public String datenew;
	public String report;
	public String path;

	// ExtentReports report = ExtentReports.TestNG.Class;
	// @BeforeTest
	// @BeforeClass
	@BeforeMethod
	public void setupreport() throws UnknownHostException {

		try {
			System.out.println(" before method: setup report");
			String path = System.getProperty("user.dir");   // return project folder path

			String driverpath = path + "\\framework_lib\\chromedriver.exe";   // return driver folder path 

			System.setProperty("webdriver.chrome.driver",driverpath );
			
			//test-output
			//WebDriver driver= new ChromeDriver();
			/*
			 * String path = System.getProperty("user.dir"); // return project
			 * folder path
			 * 
			 * String driverpath = path + "\\framework_lib\\chromedriver.exe";
			 * // return driver folder path
			 * 
			 * System.setProperty("webdriver.chrome.driver",driverpath );
			 */
			Calendar calendar = Calendar.getInstance();
			SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-yyyy");
			System.out.println(formatter.format(calendar.getTime()));
			datenew = formatter.format(calendar.getTime());

			Date date = java.util.Calendar.getInstance().getTime();
			System.out.println(date);

			/*
			 * SimpleDateFormat formatter = new
			 * SimpleDateFormat("dd/MM/yyyy HH:mm:ss"); Date date = new Date();
			 * System.out.println(formatter.format(date));
			 */
			 path = System.getProperty("user.dir");
			String reportpath = ".\\Reports\\GacTestingReport " + datenew
							+ ".html";
			
			htmlReporter= new ExtentHtmlReporter(reportpath);
			/*htmlReporter = new ExtentHtmlReporter(
					"C:/New_Hybrid/Reports/GacTestingReport " + datenew
							+ ".html");*/
			/*
			 * report = htmlReporter.toString(); extent = new ExtentReports();
			 * extent.attachReporter(htmlReporter);
			 * htmlReporter.config().setTheme(Theme.STANDARD);
			 * //htmlReporter.config().getCSS().
			 */
			report = "GacTestingReport" + " " + datenew + ".html";
			// report= htmlReporter+"";
			extent = new ExtentReports();
			extent.attachReporter(htmlReporter);
			htmlReporter.config().setTheme(Theme.STANDARD);
			// htmlReporter.config().getCSS().

			extent.setSystemInfo("Windos", "Gacituser");
			extent.setSystemInfo("Host Name", InetAddress.getLocalHost()
					.getHostName());
			extent.setSystemInfo("Environment", "QA");
			extent.setSystemInfo("User Name", "Gacituser");
			
			System.out.println(" before method: setup report");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/*
	 * extent.config(); htmlReporter.start(); htmlReporter = new
	 * ExtentHtmlReporter("C:\\New_Hybrid\\Newextentreport.html");
	 * 
	 * 
	 * extent.setSystemInfo("Envornment", "QA"); //String className =
	 * this.getClass().getSimpleName(); test= extent.createTest("Testcase");
	 */

	@Test
	public void automationscript() throws Throwable {
System.out.println("automation script");
		ActionEvents act = new ActionEvents();

		XSSFSheet sheet = Read.opencasesheet(); // --------initializing test
												// case sheet
		XSSFSheet automationsteps = Read.openstepSheet(); // ---------initializing
															// test step sheet

		int LastRowNumber_testcaseSheet = sheet.getLastRowNum(); // ---------collect
																	// number of
																	// rows in
																	// testcase
																	// sheet
		int LastRowNumber_testStepSheet = automationsteps.getLastRowNum(); // ----------collect
																			// number
																			// of
																			// rows
																			// in
																			// teststep
																			// sheet
		// System.out.println("number of rows in testcase sheet  " +
		// LastRowNumber_testStepSheet);
		// System.out.println("number of rows in teststep sheet  "+LastRowNumber_testcaseSheet);
		// Log.info("Number of Test Steps Rows in sheet is :  " +
		// LastRowNumber_testStepSheet);
		// Log.info("Number of Test Cases Rows in sheet is :  "+LastRowNumber_testcaseSheet);
		// Reporter.log("Number of Test Steps Rows in sheet is :  " +
		// LastRowNumber_testStepSheet);
		// Reporter.log("Number of Test Cases Rows in sheet is :  "+LastRowNumber_testcaseSheet);
		// ----------running from 0 to last row in testcase sheet
		for (int Row_TestcaseSheet = 0; Row_TestcaseSheet <= LastRowNumber_testcaseSheet; Row_TestcaseSheet++) {
			String runmode = Read.runmode(Row_TestcaseSheet);

			try {

				if (runmode.equals("yes"))

				{

					String nor = sheet.getRow(Row_TestcaseSheet).getCell(8)
							.getStringCellValue();
					int NoOf_repeatation = Integer.valueOf(nor);
					// System.out.println("NoOf repreatation  ="+NoOf_repeatation);
					Log.info("No. of repeatation  =" + NoOf_repeatation);
					Reporter.log("No. of repeatation  =" + NoOf_repeatation);

					if (NoOf_repeatation != 1) {
						// System.out.println("no of repeatation  ="+NoOf_repeatation);
						Log.info("No. of Repeatation  =" + NoOf_repeatation);
						// Reporter.log("No. of Repeatation  ="+NoOf_repeatation);
						String ini = Read.cellreadCASE(Row_TestcaseSheet, 9);
						int RowNumber = Integer.valueOf(ini) - 1;
						int ResultRN = Integer.valueOf(ini) - 1;

						for (int repeat = 1; repeat <= NoOf_repeatation; repeat++)
						// abc

						{

							// Tid is testcase Id in test case sheet
							String Tid = sheet.getRow(Row_TestcaseSheet)
									.getCell(0).getStringCellValue();//

							for (int teststeps = 0; teststeps <= LastRowNumber_testStepSheet; teststeps++) {
								// tid is test case id in test step sheet
								String tid = automationsteps.getRow(teststeps)
										.getCell(0).getStringCellValue();
								String method = automationsteps
										.getRow(teststeps).getCell(2)
										.getStringCellValue();
								// test =
								// extent.startTest(tid+"   "+method+"_"+":"+
								// "_start_");
								// test.log(LogStatus.FAIL, it.getThrowable());
								if (tid.equals(Tid)) {
									// System.out.println("Row numer  ="+RowNumber+" and steps no  ="+teststeps);
									Log.info("Row number  =" + RowNumber
											+ " and steps no  =" + teststeps);
									Reporter.log("Row number  =" + RowNumber
											+ " and steps no  =" + teststeps);
									String data = Read.Rdata(RowNumber,
											teststeps, Tid);

									String locator = Read.locators(teststeps);

									act.ActionR(teststeps, locator, data);
								}
								Read.save();

								if (tid.equals(Tid)) {
									String RowId_InputData = Read.Rdata(
											RowNumber, 0, Tid);
									String Step_Status = Read
											.Step_status(teststeps);
									// System.out.println(Step_Status);
									Log.info(Step_Status);
									Reporter.log(Step_Status);
									if (Step_Status.equals("fail")) {
										// test.log(LogStatus.FAIL,"step failed");
										String previousStatus = Read
												.cellreadCASE(
														Row_TestcaseSheet, 12);
										String testSpetID = Read
												.openstepSheet()
												.getRow(teststeps).getCell(1)
												.getStringCellValue();
										// System.out.println("previous status is   "+previousStatus+" and ini ="+ini);
										Log.info("previous status is   "
												+ previousStatus + " and ini ="
												+ ini);
										Reporter.log("previous status is   "
												+ previousStatus + " and ini ="
												+ ini);

										if (Step_Status.equals("****")) {
											Read.Case_cellwrite(
													Row_TestcaseSheet, 10,
													"fail");
											Read.Case_cellwrite(
													Row_TestcaseSheet, 12,
													"fail for input data Row Id="
															+ RowId_InputData
															+ " At step no.="
															+ testSpetID
															+ ",.0,");
											ResultRN = ResultRN + 1;
											break;
										}

										else {
											// test.log(LogStatus.FAIL,"step failed");
											Read.Case_cellwrite(
													Row_TestcaseSheet, 10,
													"fail");
											Read.Case_cellwrite(
													Row_TestcaseSheet,
													12,
													previousStatus
															+ ","
															+ "fail for input data Row="
															+ RowId_InputData
															+ " At step no.="
															+ testSpetID);
											ResultRN = ResultRN + 1;
											break;
										}
									} else if (Step_Status.equals("pass")) {
										Read.Case_cellwrite(Row_TestcaseSheet,
												10, "pass");
										// test.log(LogStatus.PASS,"step passed");
									}

								}
							}
							RowNumber = RowNumber + 1;

						}
					}

					if (NoOf_repeatation == 1) {
						String Tid = sheet.getRow(Row_TestcaseSheet).getCell(0)
								.getStringCellValue();
						// System.out.println("testcase id ="+Tid);
						Log.info("Testcase id =  " + Tid);
						Reporter.log("Testcase id  = " + Tid);
						// parent=extent.startTest(Tid);
						// parent=extent.startTest(Tid);
						parent = extent.createTest(Tid);
						for (int testSteps = 0; testSteps <= LastRowNumber_testStepSheet; testSteps++) {

							String tid = automationsteps.getRow(testSteps)
									.getCell(0).getStringCellValue();
							String method = automationsteps.getRow(testSteps)
									.getCell(2).getStringCellValue();
							// test = extent.startTest(tid+"_"+":"+ "_start_");

							if (tid.equals(Tid)) {
								// test= extent.createTest(method);
								test = parent.createNode(method);
								// test.log(Status.INFO, tid+" :   :  "+method);

								// System.out.println(test=
								// extent.startTest(method));
								test.log(Status.INFO, tid + " :   :  " + method);
								// test =
								// extent.startTest(tid+" :   :  "+method);
								// test= extent.startTest(method);
								// test.log(LogStatus.INFO,
								// tid+" :   :  "+method);

								// parent = extent.startTest("ClassName");

								// test.appendChild(test);
								String locator = Read.locators(testSteps);
								ExtentReports extent = null;
								// ------------this object call action function
								act.Action(testSteps, locator, test);
							}
						}

					}

				} else if (runmode.equals("no")) {

					String Tid = sheet.getRow(Row_TestcaseSheet).getCell(0)
							.getStringCellValue();
					// System.out.println(Tid+"  Skipped as User said No Execution");
					// Log.info(Tid+"  Skipped as User said No Execution");
					// Reporter.log(Tid+"  Skipped as User said No Execution");
				}

			}

			catch (Exception e) {

			}

		}
		Read.save();

		// ////////////////////////////////////////////////////////////////////////////////////////////////////////////

		for (int Row_TestcaseSheet = 0; Row_TestcaseSheet <= LastRowNumber_testcaseSheet; Row_TestcaseSheet++) {
			String runmode = Read.runmode(Row_TestcaseSheet);
			try {
				if (runmode.equals("yes")) {

					String nor = sheet.getRow(Row_TestcaseSheet).getCell(8)
							.getStringCellValue();
					int NoOf_repeatation = Integer.valueOf(nor);

					if (NoOf_repeatation == 1) {
						String Tid = sheet.getRow(Row_TestcaseSheet).getCell(0)
								.getStringCellValue();
						// System.out.println("testcase id  =" + Tid);
						Log.info("Testcase id  =" + Tid);
						Reporter.log("Testcase id  =" + Tid);

						for (int testSteps = 0; testSteps <= LastRowNumber_testStepSheet; testSteps++) {
							String tid = automationsteps.getRow(testSteps)
									.getCell(0).getStringCellValue();
							if (tid.equals(Tid)) {
								String Step_Status = Read
										.Step_status(testSteps);
								// System.out.println("Test Case ID: "+ Tid +
								// " is " + Step_Status);
								Log.info("Test Step ID: " + Tid + " is "
										+ Step_Status);
								Reporter.log("Test Step ID: " + Tid + " is "
										+ Step_Status);
								if (Step_Status.equals("fail")) {
									Read.Case_cellwrite(Row_TestcaseSheet, 10,
											"fail");
									// System.out.println("Test Case ID: "+ Tid
									// + " is " + Step_Status);
									Log.info("Test Step ID: " + Tid + " is "
											+ Step_Status);
									Reporter.log("Test Step ID: " + Tid
											+ " is " + Step_Status);
									break;
								} else if (Step_Status.equals("pass")) {
									Read.Case_cellwrite(Row_TestcaseSheet, 10,
											"pass");
									Log.info("Test Step ID: " + Tid + " is "
											+ Step_Status);
									Reporter.log("Test Step ID: " + Tid
											+ " is " + Step_Status);
								}
							}
						}
					}

				} else if (runmode.equals("no")) {
					String Tid = sheet.getRow(Row_TestcaseSheet).getCell(0)
							.getStringCellValue();
					Read.Case_cellwrite(Row_TestcaseSheet, 10,
							" Skipped as User said No Execution ");
					Log.info(Tid + "Skipped as User said No Execution");
					Reporter.log(Tid + "Skipped as User said No Execution");
				}
			}

			catch (Exception e) {

			}

		}
		Read.save();

		// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

	}

	// @AfterTest
	// @AfterClass
	@AfterMethod
	public void teardown() {
		
		System.out.println("222222222222222222222222222222222222");
		// extent.(test);
		// driver.close();
		// reporter.stop();
		extent.flush();

		try {
			ITestResult result;
			WebDriver driver = null;
			 TakesScreenshot ts = (TakesScreenshot)driver;
		        File source = ts.getScreenshotAs(OutputType.FILE);
		        String dest = System.getProperty("user.dir") +"\\ErrorScreenshots\\"+"fail"+".png";
		        File destination = new File(dest);
		        FileUtils.copyFile(source, destination);  
		       /* if(result.getStatus() == ITestResult.FAILURE)
		        {
		           // String screenShotPath = GetScreenShot.capture(driver, "screenShotName");
		           // test.log(LogStatus.FAIL, "failed");
		            //test.log(LogStatus.FAIL, "Snapshot below: " + dest);
		        }*/
			String hostName = "mail.ad.crisil.com";
			String portNo = "25";

			JavaMailSenderImpl sender = new JavaMailSenderImpl();
			sender.setHost(hostName);
			if (null != portNo && !("").equals(portNo)) {
				sender.setPort(Integer.parseInt(portNo));
			}
			sender.send(new MimeMessagePreparator() {
				public void prepare(MimeMessage mimeMessage)
						throws MessagingException {
					MimeMessageHelper message = new MimeMessageHelper(
							mimeMessage, true, "UTF-8");
					// MimeMessage message1 = new MimeMessage(session);
					message.setFrom("anindita.rath@ext-crisil.com");
					// message.setTo("anindita.rath@ext-crisil.com");

					message.setCc("anindita.rath@ext-crisil.com");

					/*
					 * String mailTo="Put TO Info"; Address[] cc = new Address[]
					 * {InternetAddress.parse("abc@abc.com"),
					 * InternetAddress.parse("abc@def.com"),
					 * InternetAddress.parse("ghi@abc.com")};
					 */

					String mailtoreceipents = " anindita.rath@ext-crisil.com";
					String[] recipientList = mailtoreceipents.split(",");
					InternetAddress[] recipientAddress = new InternetAddress[recipientList.length];
					int counter = 0;
					for (String recipient : recipientList) {
						recipientAddress[counter] = new InternetAddress(
								recipient.trim());
						counter++;
					}
					message.setTo(recipientAddress);

					message.setSubject("Automation Testing Status");
					
					FileSystemResource file = new FileSystemResource(new File(".\\Reports\\GacTestingReport " + datenew+ ".html"));

					message.addAttachment(report, file);
					message.setText(".............................."
							+ "Automation Sanity/Smoke testing has been completed"
							+ '\n' + "...................");

					/*
					 * message.setSubject("Automation Testing Status");
					 * FileSystemResource file = new FileSystemResource(new
					 * File("C:/New_Hybrid/Reports/GactTestingRreport.html"));
					 * 
					 * message.addAttachment(report, file);
					 * message.setText(".............................."
					 * +"Automation Sanity/Smoke testing has been completed" +
					 * '\n' + "...................");
					 *//*
						 * MimeBodyPart messageBodyPart2 = new MimeBodyPart();
						 * 
						 * String filename =
						 * "C:/New_Hybrid/NewExtentreport.html";//change
						 * accordingly DataSource source = new
						 * FileDataSource(filename);
						 * messageBodyPart2.setDataHandler(new
						 * DataHandler(source));
						 * messageBodyPart2.setFileName(filename);
						 * 
						 * Multipart multipart = new MimeMultipart();
						 * multipart.addBodyPart(messageBodyPart2); ((Part)
						 * message).setContent(multipart );
						 * Transport.send(message);
						 */

					// test2.log(Status.PASS, "passed");
				}

			});

		} catch (Exception e) {
			// test2.log(Status.FAIL, " could not send mail::Failed");
			e.printStackTrace();
		}
	}

	@BeforeMethod
	public void beforeMethod() {
	}

	@AfterMethod
	public void afterMethod() {

	}

}
