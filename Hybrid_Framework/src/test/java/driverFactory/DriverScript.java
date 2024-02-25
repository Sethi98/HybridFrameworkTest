package driverFactory;

import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunctions.FunctionLibrary;
import utilities.ExcelFileUtil;

public class DriverScript 
{
	public static WebDriver driver;
	String inputpath ="./FileInput/DataEngine.xlsx";
	String outputpath = ".//FileOutput/HybridResult.xlsx";
	ExtentReports report;
	ExtentTest logger;
	public void startTest() throws Throwable
	{
		String Modulestatus ="";
		// Create object for ExcelFileUtil Class
		ExcelFileUtil xl = new ExcelFileUtil(inputpath);
		String TestCases = "MasterTestCases";
		// Iterate all TestCases from sheet
		for(int i=1;i<=xl.rowCount(TestCases);i++)
		{
			if(xl.getCellData(TestCases, i, 2).equalsIgnoreCase("Y"))
			{
				// Read all test case or corrosponding sheet
				String TCModules = xl.getCellData(TestCases, i, 1);
				// Define path of html 
				report = new ExtentReports("./target/Reports/"+TCModules+FunctionLibrary.generateDate()+".html");
				logger = report.startTest(TCModules);
				//Iterate all rows in TCModule sheet
				for(int j=1;j<=xl.rowCount(TCModules);j++)
				{
					// Read all cells from TCModule
					String Description = xl.getCellData(TCModules, j, 0);
					String Object_type = xl.getCellData(TCModules, j, 1);
					String Locator_type = xl.getCellData(TCModules, j, 2);
					String Locator_value = xl.getCellData(TCModules, j, 3);
					String Test_data = xl.getCellData(TCModules, j, 4);
					try {
						if(Object_type.equalsIgnoreCase("startBrowser"))
						{
							driver  = FunctionLibrary.startBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("openUrl"))
						{
							FunctionLibrary.openUrl();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("WaitForElement")) 
						{
							FunctionLibrary.waitForElement(Locator_type, Locator_value, Test_data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("typeAction"))
						{
							FunctionLibrary.typeAction(Locator_type, Locator_value, Test_data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("clickAction"))
						{
							FunctionLibrary.clickAction(Locator_type, Locator_value);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("validateTitle"))
						{
							FunctionLibrary.validateTitle(Test_data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("closeBrowser"))
						{
							FunctionLibrary.closeBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("dropDownAction"))
						{
							FunctionLibrary.dropDownAction(Locator_type, Locator_value, Test_data);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("captureStockNum"))
						{
							FunctionLibrary.captureStockNum(Locator_type, Locator_value);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("stockTable"))
						{
							FunctionLibrary.stockTable();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("capturesup"))
						{
							FunctionLibrary.capturesup(Locator_type, Locator_value);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("supplierTable"))
						{
							FunctionLibrary.supplierTable();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("capturecus"))
						{
							FunctionLibrary.capturecus(Locator_type, Locator_value);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_type.equalsIgnoreCase("customerTable"))
						{
							FunctionLibrary.customerTable();
							logger.log(LogStatus.INFO, Description);
						}
						// Write as pass into status cell in TCmodule
						xl.setCellData(TCModules, j, 5, "Pass", outputpath);
						logger.log(LogStatus.PASS, Description);
						Modulestatus = "true";
					} catch (Exception e) {
						System.out.println(e.getMessage());
						// Write as Fail in status cell in TCModules
						xl.setCellData(TCModules, j, 5, "Fail", outputpath);
						logger.log(LogStatus.FAIL, Description);
						Modulestatus = "False";
					}
					if(Modulestatus.equalsIgnoreCase("true"))
					{
						// Write as pass in TestCases sheet
						xl.setCellData(TestCases, i, 3, "Pass", outputpath);
					}else
					{
						// Write Fail into TestCases sheet
						xl.setCellData(TestCases, i, 3, "Fail", outputpath);
					}
					report.endTest(logger);
					report.flush();
				}
				}else
				{
					//write as blocked into status cell for Flag N
					xl.setCellData(TestCases, i, 3, "Blocked", outputpath);
				}
			
			
			
		
		}
	}

}
